from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import requests
from msal import PublicClientApplication, SerializableTokenCache


CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d"
TENANT = "common"
DATAVERSE_ORG_URL = "https://org9621eba7.crm.dynamics.com"
DATAVERSE_API = f"{DATAVERSE_ORG_URL}/api/data/v9.2"

FLOW_NAME = "PA_S1_RegistrarPresupuestoAprobado"
SITE_URL = "https://sciprowall.sharepoint.com/sites/PROYECTOSPROWALL"
DOCUMENT_LIBRARY_ID = "8d884441-daa4-45b5-beb4-f09abdd1433a"
QUEUE_LIST_ID = "91e1c898-b8fb-44a9-8236-a600c6e38340"
INBOX_FOLDER_PATH = "/Documentos compartidos/Proyectos/Presupuestos Aprobados"
INBOX_PATH_MATCH = "/Proyectos/Presupuestos Aprobados/"


def load_dataverse_token(cache_path: Path) -> str:
    cache = SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))
    app = PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT}",
        token_cache=cache,
    )
    scopes = [f"{DATAVERSE_ORG_URL}/.default"]
    for account in app.get_accounts():
        result = app.acquire_token_silent(scopes, account=account)
        if result and "access_token" in result:
            cache_path.write_text(cache.serialize(), encoding="utf-8")
            return str(result["access_token"])
    raise RuntimeError(
        "No hay token Dataverse en cache. Ejecute primero el device login y guarde .dataverse_token_cache.json."
    )


def dataverse_headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json; charset=utf-8",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    }


def dataverse_get(token: str, path: str) -> dict[str, Any]:
    response = requests.get(
        f"{DATAVERSE_API}{path}",
        headers=dataverse_headers(token),
        timeout=60,
    )
    response.raise_for_status()
    return response.json()


def dataverse_post(token: str, path: str, payload: dict[str, Any]) -> requests.Response:
    response = requests.post(
        f"{DATAVERSE_API}{path}",
        headers=dataverse_headers(token),
        data=json.dumps(payload, ensure_ascii=False),
        timeout=60,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"Dataverse POST {response.status_code}: {response.text[:2000]}")
    return response


def dataverse_patch(token: str, path: str, payload: dict[str, Any]) -> requests.Response:
    headers = dataverse_headers(token)
    headers["If-Match"] = "*"
    response = requests.patch(
        f"{DATAVERSE_API}{path}",
        headers=headers,
        data=json.dumps(payload, ensure_ascii=False),
        timeout=60,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"Dataverse PATCH {response.status_code}: {response.text[:2000]}")
    return response


def build_flow_definition() -> dict[str, Any]:
    return {
        "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
        "contentVersion": "1.0.0.0",
        "parameters": {
            "$connections": {"defaultValue": {}, "type": "Object"},
            "$authentication": {"defaultValue": {}, "type": "SecureObject"},
        },
        "triggers": {
            "When_a_file_is_created_properties_only": {
                "type": "OpenApiConnection",
                "inputs": {
                    "host": {
                        "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                        "connectionName": "shared_sharepointonline",
                        "operationId": "GetOnNewFileItems",
                    },
                    "parameters": {
                        "dataset": SITE_URL,
                        "table": DOCUMENT_LIBRARY_ID,
                        "folderPath": INBOX_FOLDER_PATH,
                    },
                    "authentication": "@parameters('$authentication')",
                },
                "recurrence": {"frequency": "Minute", "interval": 1},
                "splitOn": "@triggerOutputs()?['body/value']",
            }
        },
        "actions": {
            "Condition_archivo_xlsx_en_inbox": {
                "type": "If",
                "expression": {
                    "and": [
                        {
                            "endsWith": [
                                "@toLower(triggerOutputs()?['body/{FilenameWithExtension}'])",
                                ".xlsx",
                            ]
                        },
                        {
                            "not": {
                                "startsWith": [
                                    "@triggerOutputs()?['body/{FilenameWithExtension}']",
                                    "~$",
                                ]
                            }
                        },
                        {
                            "contains": [
                                "@triggerOutputs()?['body/{Path}']",
                                INBOX_PATH_MATCH,
                            ]
                        },
                    ]
                },
                "actions": {
                    "Create_item_Cola_Automatizacion_Proyectos": {
                        "type": "OpenApiConnection",
                        "inputs": {
                            "host": {
                                "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                                "connectionName": "shared_sharepointonline",
                                "operationId": "PostItem",
                            },
                            "parameters": {
                                "dataset": SITE_URL,
                                "table": QUEUE_LIST_ID,
                                "item/Title": "@triggerOutputs()?['body/{FilenameWithExtension}']",
                                "item/EventType/Value": "presupuesto_aprobado",
                                "item/Estado/Value": "Pendiente",
                                "item/Filename": "@triggerOutputs()?['body/{FilenameWithExtension}']",
                                "item/FileID": "@triggerOutputs()?['body/{Identifier}']",
                                "item/FolderPath": "@triggerOutputs()?['body/{Path}']",
                                "item/FileLink": "@triggerOutputs()?['body/{Link}']",
                                "item/CreatedByEmail": "@triggerOutputs()?['body/Author/Email']",
                                "item/CreatedTime": "@triggerOutputs()?['body/Created']",
                                "item/Intentos": "0",
                                "item/Notas": "Registrado por Power Automate",
                            },
                            "authentication": "@parameters('$authentication')",
                        },
                        "runAfter": {},
                    }
                },
                "else": {"actions": {}},
                "runAfter": {},
            }
        },
        "outputs": {},
    }


def build_clientdata() -> str:
    return json.dumps(
        {
            "properties": {
                "connectionReferences": {
                    "shared_sharepointonline": {
                        "runtimeSource": "embedded",
                        "connection": {},
                        "api": {"name": "shared_sharepointonline"},
                    }
                },
                "definition": build_flow_definition(),
            },
            "schemaVersion": "1.0.0.0",
        },
        ensure_ascii=False,
        separators=(",", ":"),
    )


def find_existing_flow(token: str) -> dict[str, Any] | None:
    query_name = FLOW_NAME.replace("'", "''")
    data = dataverse_get(
        token,
        "/workflows"
        f"?$select=workflowid,name,category,statecode,type"
        f"&$filter=category eq 5 and name eq '{query_name}'",
    )
    values = data.get("value") or []
    return values[0] if values else None


def save_local_definition(output_dir: Path, clientdata: str) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / f"{FLOW_NAME}.clientdata.json").write_text(
        json.dumps(json.loads(clientdata), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    (output_dir / f"{FLOW_NAME}.definition.json").write_text(
        json.dumps(build_flow_definition(), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Crea o actualiza el cloud flow de registro de presupuestos.")
    parser.add_argument("--token-cache", default=".dataverse_token_cache.json")
    parser.add_argument("--output-dir", default="power_automate/solution/generated")
    parser.add_argument("--apply", action="store_true", help="Crea/actualiza el workflow en Dataverse.")
    args = parser.parse_args()

    clientdata = build_clientdata()
    save_local_definition(Path(args.output_dir), clientdata)
    if not args.apply:
        print("Dry-run: definicion local generada. Use --apply para subir a Dataverse.")
        return 0

    token = load_dataverse_token(Path(args.token_cache))
    existing = find_existing_flow(token)
    payload = {
        "category": 5,
        "name": FLOW_NAME,
        "type": 1,
        "description": (
            "Sistema 1 Autosys: registra presupuestos .xlsx nuevos de "
            "Proyectos/Presupuestos Aprobados en Cola_Automatizacion_Proyectos. "
            "El flujo queda en borrador/off hasta configurar la conexion SharePoint."
        ),
        "primaryentity": "none",
        "clientdata": clientdata,
    }
    if existing:
        workflow_id = existing["workflowid"]
        dataverse_patch(token, f"/workflows({workflow_id})", payload)
        print(f"UPDATED {FLOW_NAME} workflowid={workflow_id}")
    else:
        response = dataverse_post(token, "/workflows", payload)
        entity_id = response.headers.get("OData-EntityId", "")
        workflow_id = entity_id.rsplit("(", 1)[-1].rstrip(")") if "(" in entity_id else ""
        print(f"CREATED {FLOW_NAME} workflowid={workflow_id or '(ver Dataverse)'}")
    print("Estado esperado: Draft/Off. Configure conexion SharePoint y active manualmente en Power Automate.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
