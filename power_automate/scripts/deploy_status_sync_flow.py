from __future__ import annotations

import argparse
import json
import sys
import uuid
from pathlib import Path
from typing import Any

import requests

from fix_notification_flows import (
    API_VERSION,
    ENVIRONMENT,
    FLOW_API,
    acquire_flow_token,
    flow_headers,
)


FLOW_NAME = "PA_S2_SincronizarEstadoGanttExcel"
SITE_URL = "https://sciprowall.sharepoint.com/sites/PROYECTOSPROWALL"
CONTROL_LIST_ID = "afe5544b-3f60-40e4-81a0-01e86920f5f2"
SHAREPOINT_CONNECTION = (
    "shared-sharepointonl-1473fe25-8e08-4ffa-86c4-606465f96830"
)
EXCEL_CONNECTION = "a607013d36214a7288b7cec8cf4d955f"
WORKBOOK_DRIVE_ID = (
    "b!gT--atFwrEiNN1Dm9ygu7gAK2MbJi0RKuVaPM52nyYJBRIiNpNq1Rb608Jq90UM6"
)
OFFICE_SCRIPT_ID = (
    "ms-officescript%3A%2F%2Fonedrive_business_itemlink%2F"
    "01DOEBQSOGBUUS4J3DDNEZEYWXPWEC4VXH"
)


def open_api_action(
    *,
    connection: str,
    operation: str,
    parameters: dict[str, Any],
    run_after: dict[str, list[str]] | None = None,
) -> dict[str, Any]:
    return {
        "runAfter": run_after or {},
        "type": "OpenApiConnection",
        "inputs": {
            "parameters": parameters,
            "host": {
                "apiId": (
                    "/providers/Microsoft.PowerApps/apis/"
                    f"{connection}"
                ),
                "connectionName": connection,
                "operationId": operation,
            },
            "authentication": "@parameters('$authentication')",
        },
    }


def sharepoint_update(
    fields: dict[str, Any],
    *,
    run_after: dict[str, list[str]] | None = None,
) -> dict[str, Any]:
    parameters: dict[str, Any] = {
        "dataset": SITE_URL,
        "table": CONTROL_LIST_ID,
        "id": "@triggerBody()?['ID']",
        "item/Title": "@triggerBody()?['Title']",
    }
    parameters.update({f"item/{name}": value for name, value in fields.items()})
    return open_api_action(
        connection="shared_sharepointonline",
        operation="PatchItem",
        parameters=parameters,
        run_after=run_after,
    )


def build_definition() -> dict[str, Any]:
    success_expression = (
        "@or("
        "equals(outputs('Run_script')?['body/result/code'],'ACTUALIZADO'),"
        "equals(outputs('Run_script')?['body/result/code'],'YA_SINCRONIZADO')"
        ")"
    )
    final_status = "@outputs('Run_script')?['body/result/finalStatus']"
    result_code = "@outputs('Run_script')?['body/result/code']"
    return {
        "$schema": (
            "https://schema.management.azure.com/providers/Microsoft.Logic/"
            "schemas/2016-06-01/workflowdefinition.json#"
        ),
        "contentVersion": "1.0.0.0",
        "parameters": {
            "$connections": {"defaultValue": {}, "type": "Object"},
            "$authentication": {"defaultValue": {}, "type": "SecureObject"},
        },
        "triggers": {
            "When_an_item_is_created_or_modified": {
                "recurrence": {"frequency": "Minute", "interval": 1},
                "evaluatedRecurrence": {"frequency": "Minute", "interval": 1},
                "splitOn": "@triggerOutputs()?['body/value']",
                "type": "OpenApiConnection",
                "inputs": {
                    "parameters": {
                        "dataset": SITE_URL,
                        "table": CONTROL_LIST_ID,
                    },
                    "host": {
                        "apiId": (
                            "/providers/Microsoft.PowerApps/apis/"
                            "shared_sharepointonline"
                        ),
                        "connectionName": "shared_sharepointonline",
                        "operationId": "GetOnUpdatedItems",
                    },
                    "authentication": "@parameters('$authentication')",
                },
                "conditions": [
                    {
                        "expression": (
                            "@and("
                            "equals(triggerBody()?['EstadoSyncExcel'],'Pendiente'),"
                            "not(empty(triggerBody()?['GanttWorkingIdentifier'])),"
                            "or("
                            "equals(triggerBody()?['StatusExcelDeseado'],'Actual'),"
                            "equals(triggerBody()?['StatusExcelDeseado'],'En Progreso')"
                            ")"
                            ")"
                        )
                    }
                ],
                "runtimeConfiguration": {"concurrency": {"runs": 1}},
            }
        },
        "actions": {
            "Mark_sync_processing": sharepoint_update(
                {"EstadoSyncExcel": "Procesando"}
            ),
            "Run_script": open_api_action(
                connection="shared_excelonlinebusiness",
                operation="RunScriptProd",
                parameters={
                    "source": SITE_URL,
                    "drive": WORKBOOK_DRIVE_ID,
                    "file": "@triggerBody()?['GanttWorkingIdentifier']",
                    "scriptId": OFFICE_SCRIPT_ID,
                    "ScriptParameters/desiredStatus": (
                        "@triggerBody()?['StatusExcelDeseado']"
                    ),
                    "ScriptParameters/expectedCurrentStatus": (
                        "@triggerBody()?['StatusExcel']"
                    ),
                    "ScriptParameters/allowReplaceEntregar": (
                        "@and("
                        "equals(triggerBody()?['StatusExcelDeseado'],'Actual'),"
                        "equals("
                        "coalesce("
                        "triggerBody()?['EstadoGantt']?['Value'],"
                        "triggerBody()?['EstadoGannt']?['Value'],"
                        "triggerBody()?['EstadoGantt'],"
                        "triggerBody()?['EstadoGannt']"
                        "),"
                        "'Actual'"
                        "),"
                        "not(empty(coalesce("
                        "triggerBody()?['VersionActual'],"
                        "triggerBody()?['VersionadoActual']"
                        ")))"
                        ")"
                    ),
                },
                run_after={"Mark_sync_processing": ["Succeeded"]},
            ),
            "Handle_script_result": {
                "runAfter": {"Run_script": ["Succeeded"]},
                "type": "If",
                "expression": success_expression,
                "actions": {
                    "Wait_for_SharePoint_metadata": {
                        "runAfter": {},
                        "type": "Wait",
                        "inputs": {
                            "interval": {"count": 30, "unit": "Second"}
                        },
                    },
                    "Get_file_metadata": open_api_action(
                        connection="shared_sharepointonline",
                        operation="GetFileMetadata",
                        parameters={
                            "dataset": SITE_URL,
                            "id": "@triggerBody()?['GanttWorkingIdentifier']",
                        },
                        run_after={
                            "Wait_for_SharePoint_metadata": ["Succeeded"]
                        },
                    ),
                    "Mark_sync_success": sharepoint_update(
                        {
                            "StatusExcel": final_status,
                            "EstadoSyncExcel": "Sincronizado",
                            "FechaUltimoSyncExcel": "@utcNow()",
                            "UltimoETagAutomatizacion": (
                                "@body('Get_file_metadata')?['ETag']"
                            ),
                            "GanttWorkingETag": (
                                "@body('Get_file_metadata')?['ETag']"
                            ),
                            "FechaUltimaModificacionGantt": (
                                "@body('Get_file_metadata')?['Modified']"
                            ),
                            "UltimoErrorSyncExcel": "",
                            "ProximoIntentoSyncExcel": None,
                        },
                        run_after={"Get_file_metadata": ["Succeeded"]},
                    ),
                },
                "else": {
                    "actions": {
                        "Mark_sync_omitted": sharepoint_update(
                            {
                                "StatusExcel": final_status,
                                "EstadoSyncExcel": "Omitido",
                                "FechaUltimoSyncExcel": "@utcNow()",
                                "UltimoErrorSyncExcel": result_code,
                            }
                        )
                    }
                },
            },
            "Mark_sync_error": sharepoint_update(
                {
                    "EstadoSyncExcel": "Error",
                    "IntentosSyncExcel": (
                        "@add(coalesce(triggerBody()?['IntentosSyncExcel'],0),1)"
                    ),
                    "UltimoErrorSyncExcel": (
                        "@string(outputs('Run_script'))"
                    ),
                    "ProximoIntentoSyncExcel": "@addMinutes(utcNow(),15)",
                },
                run_after={
                    "Run_script": ["Failed", "TimedOut"],
                },
            ),
        },
        "outputs": {},
    }


def connection_references() -> dict[str, Any]:
    return {
        "shared_sharepointonline": {
            "connectionName": SHAREPOINT_CONNECTION,
            "source": "Embedded",
            "id": (
                "/providers/Microsoft.PowerApps/apis/"
                "shared_sharepointonline"
            ),
            "displayName": "SharePoint",
            "tier": "Standard",
            "apiName": "sharepointonline",
        },
        "shared_excelonlinebusiness": {
            "connectionName": EXCEL_CONNECTION,
            "source": "Embedded",
            "id": (
                "/providers/Microsoft.PowerApps/apis/"
                "shared_excelonlinebusiness"
            ),
            "displayName": "Excel Online (Business)",
            "tier": "Standard",
            "apiName": "excelonlinebusiness",
        },
    }


def flow_collection_url() -> str:
    return (
        f"{FLOW_API}/providers/Microsoft.ProcessSimple/environments/"
        f"{ENVIRONMENT}/flows?api-version={API_VERSION}"
    )


def flow_url(flow_id: str) -> str:
    return (
        f"{FLOW_API}/providers/Microsoft.ProcessSimple/environments/"
        f"{ENVIRONMENT}/flows/{flow_id}?api-version={API_VERSION}"
    )


def find_existing(token: str) -> dict[str, Any] | None:
    response = requests.get(
        flow_collection_url(),
        headers=flow_headers(token),
        timeout=60,
    )
    response.raise_for_status()
    for flow in response.json().get("value", []):
        if flow.get("properties", {}).get("displayName") == FLOW_NAME:
            return flow
    return None


def deploy(token: str, *, start: bool) -> tuple[str, str]:
    existing = find_existing(token)
    flow_id = str(existing.get("name")) if existing else str(uuid.uuid4())
    payload = {
        "properties": {
            "displayName": FLOW_NAME,
            "definition": build_definition(),
            "connectionReferences": connection_references(),
            "state": "Started" if start else "Stopped",
        }
    }
    if existing:
        response = requests.patch(
            flow_url(flow_id),
            headers=flow_headers(token),
            json=payload,
            timeout=120,
        )
    else:
        payload["name"] = flow_id
        response = requests.post(
            flow_collection_url(),
            headers=flow_headers(token),
            json=payload,
            timeout=120,
        )
    if not response.ok:
        raise RuntimeError(
            f"Power Automate {response.request.method} {response.status_code}: "
            f"{response.text[:4000]}"
        )
    deployed = response.json()
    return flow_id, str(deployed.get("properties", {}).get("state") or "")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Crea o actualiza el flujo que sincroniza Gantt!B6."
    )
    parser.add_argument("--token-cache", default=".dataverse_token_cache.json")
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--start", action="store_true")
    parser.add_argument(
        "--output",
        default=(
            "power_automate/solution/generated/"
            "PA_S2_SincronizarEstadoGanttExcel.definition.json"
        ),
    )
    args = parser.parse_args()
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(build_definition(), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Definición generada: {output}")
    if not args.apply:
        print("Dry-run. Use --apply para desplegar detenido.")
        return 0
    token = acquire_flow_token(Path(args.token_cache))
    flow_id, state = deploy(token, start=args.start)
    print(f"DEPLOYED flow={FLOW_NAME} id={flow_id} state={state}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
