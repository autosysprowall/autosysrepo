from __future__ import annotations

import argparse
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


def corrected_return_definition(definition: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(definition)
    parameters = result["actions"]["Send_an_email_(V2)"]["inputs"]["parameters"]
    parameters["emailMessage/Subject"] = (
        "@concat('[BORRADOR] Presupuesto devuelto - ', "
        "coalesce(triggerBody()?['Filename'], triggerBody()?['Title'], 'sin nombre'))"
    )
    parameters["emailMessage/Body"] = (
        "@concat('<p class=\"editor-paragraph\">"
        "BORRADOR PENDIENTE DE REDACCIÓN. Motivo técnico: ', "
        "coalesce(triggerBody()?['UltimoError'], 'No especificado'), '</p>')"
    )
    return result


def corrected_assignment_definition(definition: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(definition)
    parameters = result["actions"]["Send_an_email_(V2)"]["inputs"]["parameters"]
    parameters["emailMessage/To"] = "@triggerBody()?['To']"
    parameters["emailMessage/Cc"] = "@triggerBody()?['Cc']"
    parameters["emailMessage/Subject"] = "@triggerBody()?['Subject']"
    parameters["emailMessage/Body"] = "@triggerBody()?['Body']"
    return result


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


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Corrige bindings dinámicos de los flujos de notificación."
    )
    parser.add_argument("--token-cache", default=".dataverse_token_cache.json")
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    token = acquire_flow_token(Path(args.token_cache))
    return_flow = get_flow(token, RETURN_FLOW_ID)
    assignment_flow = get_flow(token, ASSIGNMENT_FLOW_ID)
    return_properties = return_flow["properties"]
    assignment_properties = assignment_flow["properties"]
    return_definition = corrected_return_definition(return_properties["definition"])
    assignment_definition = corrected_assignment_definition(
        assignment_properties["definition"]
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
