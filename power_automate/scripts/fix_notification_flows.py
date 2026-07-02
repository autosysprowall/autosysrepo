from __future__ import annotations

import argparse
import base64
import copy
import json
from pathlib import Path
from typing import Any

import requests
from msal import PublicClientApplication, SerializableTokenCache


CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d"
AUTHORITY = "https://login.microsoftonline.com/common"
FLOW_SCOPE = "https://service.flow.microsoft.com/.default"
FLOW_API = "https://api.flow.microsoft.com"
ENVIRONMENT = "Default-ed7d4cfd-f42f-48ee-85ae-2e2be3539cd4"
API_VERSION = "2016-11-01"

RETURN_FLOW_ID = "56a4ec6a-5fd0-4c74-b9cb-7c65b0375928"
ASSIGNMENT_FLOW_ID = "e5dfdb04-f552-4de8-ad90-df91abfa862d"
SITE_URL = "https://sciprowall.sharepoint.com/sites/PROYECTOSPROWALL"
CONTROL_LIST_ID = "afe5544b-3f60-40e4-81a0-01e86920f5f2"
TEST_RECIPIENT = "auto.sys@prowallpanama.com"
RETURN_LIVE_CC = "jaime.madrid@prowallpanama.com"
BUDGET_GUIDE_NAME = "Guia_AutoSys_Comercial_Presupuesto.pdf"
ENGINEER_GUIDE_NAME = "Guia_AutoSys_Ingenieros_Residentes_Planta.pdf"


def acquire_flow_token(cache_path: Path) -> str:
    cache = SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))
    app = PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )
    for account in app.get_accounts():
        result = app.acquire_token_silent([FLOW_SCOPE], account=account)
        if result and "access_token" in result:
            if cache.has_state_changed:
                cache_path.write_text(cache.serialize(), encoding="utf-8")
            return str(result["access_token"])
    raise RuntimeError(
        "La sesión de Power Platform expiró. Renueve la autenticación por device code."
    )


def flow_url(flow_id: str) -> str:
    return (
        f"{FLOW_API}/providers/Microsoft.ProcessSimple/environments/"
        f"{ENVIRONMENT}/flows/{flow_id}?api-version={API_VERSION}"
    )


def flow_headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json; charset=utf-8",
    }


def get_flow(token: str, flow_id: str) -> dict[str, Any]:
    response = requests.get(
        flow_url(flow_id),
        headers=flow_headers(token),
        timeout=60,
    )
    response.raise_for_status()
    return response.json()


def patch_flow(
    token: str,
    flow_id: str,
    properties: dict[str, Any],
) -> None:
    payload = {
        "properties": {
            "displayName": properties["displayName"],
            "definition": properties["definition"],
            "connectionReferences": properties["connectionReferences"],
        }
    }
    response = requests.patch(
        flow_url(flow_id),
        headers=flow_headers(token),
        data=json.dumps(payload, ensure_ascii=False),
        timeout=60,
    )
    if response.status_code >= 400:
        raise RuntimeError(
            f"Power Automate PATCH {response.status_code}: {response.text[:2000]}"
        )


def corrected_return_definition(
    definition: dict[str, Any],
    budget_guide_base64: str = "",
) -> dict[str, Any]:
    result = copy.deepcopy(definition)
    parameters = result["actions"]["Send_an_email_(V2)"]["inputs"]["parameters"]
    parameters["emailMessage/To"] = TEST_RECIPIENT
    parameters["emailMessage/Cc"] = ""
    parameters["emailMessage/Subject"] = (
        "@concat('[PRUEBA] Presupuesto No Válido Proyecto ', "
        "if(empty(triggerBody()?['ProyectoID']), "
        "coalesce(triggerBody()?['Title'], triggerBody()?['Filename'], "
        "'sin identificar'), triggerBody()?['ProyectoID']))"
    )
    parameters["emailMessage/Body"] = (
        "@concat("
        "'<p>El presupuesto del proyecto ', "
        "if(empty(triggerBody()?['ProyectoID']), "
        "coalesce(triggerBody()?['Title'], triggerBody()?['Filename'], "
        "'sin identificar'), triggerBody()?['ProyectoID']), "
        "' no puede ser procesado por AutoSys. Por favor procurar mantener "
        "el formato sugerido en la guía adjunta.</p>', "
        "'<p><strong>Motivo técnico:</strong> ', "
        "coalesce(triggerBody()?['UltimoError'], 'No especificado'), '</p>', "
        "'<p><strong>Archivo:</strong> ', "
        "coalesce(triggerBody()?['FileLink'], 'Sin enlace disponible'), '</p>', "
        "'<hr><p><strong>MODO PRUEBA</strong><br>', "
        "'Destinatario real previsto: ', "
        "if(empty(triggerBody()?['CreatedByEmail']), '(vacío)', "
        "triggerBody()?['CreatedByEmail']), '<br>', "
        f"'CC real previsto: {RETURN_LIVE_CC}<br>', "
        f"'Este correo fue redirigido exclusivamente a {TEST_RECIPIENT}."
        "</p>')"
    )
    attachments = [
        attachment
        for attachment in parameters.get("emailMessage/Attachments", [])
        if attachment.get("Name") != BUDGET_GUIDE_NAME
    ]
    if budget_guide_base64:
        attachments.append(
            {
                "Name": BUDGET_GUIDE_NAME,
                "ContentBytes": {
                    "$content-type": "application/pdf",
                    "$content": budget_guide_base64,
                },
            }
        )
    parameters["emailMessage/Attachments"] = attachments
    return result


def corrected_assignment_definition(
    definition: dict[str, Any],
    engineer_guide_base64: str = "",
) -> dict[str, Any]:
    result = copy.deepcopy(definition)
    parameters = result["actions"]["Send_an_email_(V2)"]["inputs"]["parameters"]
    parameters["emailMessage/To"] = "@triggerBody()?['To']"
    parameters["emailMessage/Cc"] = "@triggerBody()?['Cc']"
    parameters["emailMessage/Subject"] = "@triggerBody()?['Subject']"
    parameters["emailMessage/Body"] = "@triggerBody()?['Body']"
    parameters["emailMessage/Attachments"] = (
        [
            {
                "Name": ENGINEER_GUIDE_NAME,
                "ContentBytes": {
                    "$content-type": "application/pdf",
                    "$content": engineer_guide_base64,
                },
            }
        ]
        if engineer_guide_base64
        else []
    )
    result["actions"]["Confirm_delivery_in_control"] = {
        "runAfter": {"Update_item": ["Succeeded"]},
        "cases": {
            "Assignment": {
                "case": "AsignacionGantt",
                "actions": {
                    "Confirm_assignment": control_update_action(
                        {
                            "item/CorreoAsignacionEnviado": True,
                            "item/FechaCorreoAsignacion": "@utcNow()",
                        }
                    )
                },
            },
            "Warning_1": {
                "case": "Advertencia1",
                "actions": {
                    "Confirm_warning_1": control_update_action(
                        {
                            "item/Advertencia1Enviada": True,
                            "item/FechaAdvertencia1": "@utcNow()",
                        }
                    )
                },
            },
            "Warning_2": {
                "case": "Advertencia2",
                "actions": {
                    "Confirm_warning_2": control_update_action(
                        {
                            "item/Advertencia2Enviada": True,
                            "item/FechaAdvertencia2": "@utcNow()",
                        }
                    )
                },
            },
            "Expiration": {
                "case": "Vencimiento",
                "actions": {
                    "Confirm_expiration": control_update_action(
                        {
                            "item/VencimientoNotificado": True,
                            "item/FechaVencimientoNotificado": "@utcNow()",
                        }
                    )
                },
            },
        },
        "default": {"actions": {}},
        "expression": "@triggerBody()?['TipoNotificacion']",
        "type": "Switch",
    }
    return result


def control_update_action(updates: dict[str, Any]) -> dict[str, Any]:
    parameters: dict[str, Any] = {
        "dataset": SITE_URL,
        "table": CONTROL_LIST_ID,
        "id": "@int(triggerBody()?['RelatedControlItemID'])",
    }
    parameters.update(updates)
    return {
        "type": "OpenApiConnection",
        "inputs": {
            "parameters": parameters,
            "host": {
                "apiId": (
                    "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
                ),
                "connectionName": "shared_sharepointonline",
                "operationId": "PatchItem",
            },
            "authentication": "@parameters('$authentication')",
        },
        "runAfter": {},
    }


def describe_changes(
    return_definition: dict[str, Any],
    assignment_definition: dict[str, Any],
) -> None:
    return_email = return_definition["actions"]["Send_an_email_(V2)"]["inputs"][
        "parameters"
    ]
    assignment_email = assignment_definition["actions"]["Send_an_email_(V2)"][
        "inputs"
    ]["parameters"]
    print("PA_S1_DevolverPresupuestoInvalido")
    print(f"- Subject: {return_email['emailMessage/Subject']}")
    print(f"- Body dinámico: {'UltimoError' in return_email['emailMessage/Body']}")
    print("PA_S2_EnviarNotificacionesGantt")
    print(f"- To: {assignment_email['emailMessage/To']}")
    print(f"- Cc: {assignment_email['emailMessage/Cc']}")
    print(f"- Subject: {assignment_email['emailMessage/Subject']}")
    print(f"- Body: {assignment_email['emailMessage/Body']}")
    print(
        "- Guía de ingenieros adjunta: "
        f"{len(assignment_email.get('emailMessage/Attachments', [])) == 1}"
    )
    print(
        "- Guía de presupuesto adjunta: "
        f"{len(return_email.get('emailMessage/Attachments', [])) >= 2}"
    )
    print("- Confirmación en Control_Gantt_Asignaciones: 4 tipos")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Corrige bindings dinámicos de los flujos de notificación."
    )
    parser.add_argument("--token-cache", default=".dataverse_token_cache.json")
    parser.add_argument(
        "--assets-dir",
        default="power_automate/assets",
    )
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    assets_dir = Path(args.assets_dir)
    budget_guide = base64.b64encode(
        (assets_dir / BUDGET_GUIDE_NAME).read_bytes()
    ).decode("ascii")
    engineer_guide = base64.b64encode(
        (assets_dir / ENGINEER_GUIDE_NAME).read_bytes()
    ).decode("ascii")
    token = acquire_flow_token(Path(args.token_cache))
    return_flow = get_flow(token, RETURN_FLOW_ID)
    assignment_flow = get_flow(token, ASSIGNMENT_FLOW_ID)
    return_properties = return_flow["properties"]
    assignment_properties = assignment_flow["properties"]
    return_definition = corrected_return_definition(
        return_properties["definition"],
        budget_guide,
    )
    assignment_definition = corrected_assignment_definition(
        assignment_properties["definition"],
        engineer_guide,
    )
    describe_changes(return_definition, assignment_definition)

    if not args.apply:
        print("Dry-run terminado. Use --apply para actualizar Power Automate.")
        return 0

    return_properties = dict(return_properties)
    assignment_properties = dict(assignment_properties)
    return_properties["definition"] = return_definition
    assignment_properties["definition"] = assignment_definition
    patch_flow(token, RETURN_FLOW_ID, return_properties)
    patch_flow(token, ASSIGNMENT_FLOW_ID, assignment_properties)
    print("Flujos actualizados.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
