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


FLOW_NAME = "PA_S2_GanttWorkingModificado_A_Cola"
SITE_URL = "https://sciprowall.sharepoint.com/sites/PROYECTOSPROWALL"
DOCUMENT_LIBRARY_ID = "8d884441-daa4-45b5-beb4-f09abdd1433a"
CONTROL_LIST_ID = "afe5544b-3f60-40e4-81a0-01e86920f5f2"
QUEUE_LIST_ID = "91e1c898-b8fb-44a9-8236-a600c6e38340"
SHAREPOINT_CONNECTION = (
    "shared-sharepointonl-1473fe25-8e08-4ffa-86c4-606465f96830"
)
EXCEL_CONNECTION = "a607013d36214a7288b7cec8cf4d955f"
WORKBOOK_DRIVE_ID = (
    "b!gT--atFwrEiNN1Dm9ygu7gAK2MbJi0RKuVaPM52nyYJBRIiNpNq1Rb608Jq90UM6"
)
GET_STATUS_SCRIPT_ID = (
    "ms-officescript%3A%2F%2Fonedrive_business_itemlink%2F"
    "01DOEBQSLKZQFHHXCTNJFY6QSKQ4VYIDDD"
)

CONTROL_ITEM = "first(body('Filter_control_items'))"
CONTROL_ID = f"@{CONTROL_ITEM}?['ID']"
CONTROL_STATE = (
    "coalesce("
    f"{CONTROL_ITEM}?['EstadoGantt']?['Value'],"
    f"{CONTROL_ITEM}?['EstadoGannt']?['Value'],"
    f"{CONTROL_ITEM}?['EstadoGantt'],"
    f"{CONTROL_ITEM}?['EstadoGannt']"
    ")"
)
FILE_NAME = "@triggerBody()?['{FilenameWithExtension}']"
PROJECT_ID = (
    "@first(split(triggerBody()?['{FilenameWithExtension}'],'_gantt_'))"
)
READ_STATUS = "outputs('Read_Gantt_status')?['body/result/status']"
READ_VALID = "outputs('Read_Gantt_status')?['body/result/valid']"


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


def update_control(
    fields: dict[str, Any],
    *,
    run_after: dict[str, list[str]] | None = None,
) -> dict[str, Any]:
    parameters: dict[str, Any] = {
        "dataset": SITE_URL,
        "table": CONTROL_LIST_ID,
        "id": CONTROL_ID,
        "item/Title": f"@{CONTROL_ITEM}?['Title']",
    }
    parameters.update({f"item/{name}": value for name, value in fields.items()})
    return open_api_action(
        connection="shared_sharepointonline",
        operation="PatchItem",
        parameters=parameters,
        run_after=run_after,
    )


def mark_in_progress_action() -> dict[str, Any]:
    return update_control(
        {
            "EstadoGannt/Value": "En Progreso",
            "FechaInicioEnProgreso": (
                "@if("
                f"empty({CONTROL_ITEM}?['FechaInicioEnProgreso']),"
                "utcNow(),"
                f"{CONTROL_ITEM}?['FechaInicioEnProgreso']"
                ")"
            ),
            "MinutosEnProgresoActual": (
                f"@coalesce({CONTROL_ITEM}?['MinutosEnProgresoActual'],0)"
            ),
            "StatusExcel": f"@{READ_STATUS}",
            "StatusExcelDeseado": "En Progreso",
            "EstadoSyncExcel": "Pendiente",
            "IntentosSyncExcel": 0,
            "ProximoIntentoSyncExcel": None,
            "UltimoErrorSyncExcel": "",
            "UltimoTrackingRun": "@utcNow()",
        }
    )


def build_definition() -> dict[str, Any]:
    file_filter = (
        "@and("
        "contains(toLower(triggerBody()?['{Path}']),"
        "'/proyectos/proyectos activos/'),"
        "contains(toLower(triggerBody()?['{Path}']),'/gantts/working/'),"
        "not(contains(toLower(triggerBody()?['{Path}']),"
        "'/proyectos terminados/')),"
        "not(startsWith(triggerBody()?['{FilenameWithExtension}'],'~$')),"
        "endsWith(toLower(triggerBody()?['{FilenameWithExtension}']),"
        "'_gantt_working.xlsx')"
        ")"
    )
    automation_change = (
        "@or("
        f"equals(toLower(coalesce({CONTROL_ITEM}?['EstadoSyncExcel'],'')),"
        "'procesando'),"
        "and("
        f"equals({CONTROL_ITEM}?['UltimoETagAutomatizacion'],"
        "'__POWER_AUTOMATE_PENDING_ETAG__'),"
        f"not(empty({CONTROL_ITEM}?['FechaUltimoSyncExcel'])),"
        "lessOrEquals("
        "ticks(triggerBody()?['Modified']),"
        f"ticks(addMinutes({CONTROL_ITEM}?['FechaUltimoSyncExcel'],2))"
        ")"
        ")"
        ")"
    )
    should_start_progress = (
        "@and("
        f"equals({READ_VALID},true),"
        f"or(equals({READ_STATUS},'Actual'),"
        f"equals({READ_STATUS},'En Progreso')),"
        f"equals({CONTROL_STATE},'Actual')"
        ")"
    )
    is_delivery = (
        "@and("
        f"equals({READ_VALID},true),"
        f"equals({READ_STATUS},'Entregar'),"
        "or("
        f"empty({CONTROL_ITEM}?['FechaUltimoVersionado']),"
        "less("
        f"ticks({CONTROL_ITEM}?['FechaUltimoVersionado']),"
        "ticks(triggerBody()?['Modified'])"
        ")"
        ")"
        ")"
    )
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
            "When_a_file_is_created_or_modified": {
                "recurrence": {"frequency": "Minute", "interval": 1},
                "evaluatedRecurrence": {
                    "frequency": "Minute",
                    "interval": 1,
                },
                "splitOn": "@triggerOutputs()?['body/value']",
                "type": "OpenApiConnection",
                "inputs": {
                    "parameters": {
                        "dataset": SITE_URL,
                        "table": DOCUMENT_LIBRARY_ID,
                        "folderPath": (
                            "/Documentos compartidos/Proyectos/"
                            "Proyectos Activos"
                        ),
                    },
                    "host": {
                        "apiId": (
                            "/providers/Microsoft.PowerApps/apis/"
                            "shared_sharepointonline"
                        ),
                        "connectionName": "shared_sharepointonline",
                        "operationId": "GetOnUpdatedFileItems",
                    },
                    "authentication": "@parameters('$authentication')",
                },
                "conditions": [{"expression": file_filter}],
                "runtimeConfiguration": {"concurrency": {"runs": 1}},
            }
        },
        "actions": {
            "Get_control_items": open_api_action(
                connection="shared_sharepointonline",
                operation="GetItems",
                parameters={
                    "dataset": SITE_URL,
                    "table": CONTROL_LIST_ID,
                    "$top": 500,
                },
            ),
            "Filter_control_items": {
                "runAfter": {"Get_control_items": ["Succeeded"]},
                "type": "Query",
                "inputs": {
                    "from": "@body('Get_control_items')?['value']",
                    "where": (
                        "@equals("
                        "item()?['ProyectoID'],"
                        "first(split("
                        "triggerBody()?['{FilenameWithExtension}'],"
                        "'_gantt_'"
                        "))"
                        ")"
                    ),
                },
            },
            "If_one_control_item": {
                "runAfter": {"Filter_control_items": ["Succeeded"]},
                "type": "If",
                "expression": (
                    "@equals(length(body('Filter_control_items')),1)"
                ),
                "actions": {
                    "Read_Gantt_status": open_api_action(
                        connection="shared_excelonlinebusiness",
                        operation="RunScriptProd",
                        parameters={
                            "source": SITE_URL,
                            "drive": WORKBOOK_DRIVE_ID,
                            "file": (
                                f"@{CONTROL_ITEM}?"
                                "['GanttWorkingIdentifier']"
                            ),
                            "scriptId": GET_STATUS_SCRIPT_ID,
                            "ScriptParameters": None,
                        },
                    ),
                    "If_automation_change": {
                        "runAfter": {"Read_Gantt_status": ["Succeeded"]},
                        "type": "If",
                        "expression": automation_change,
                        "actions": {
                            "Confirm_automation_change": update_control(
                                {
                                    "StatusExcel": f"@{READ_STATUS}",
                                    "EstadoSyncExcel": "Sincronizado",
                                    "UltimoTrackingRun": "@utcNow()",
                                }
                            )
                        },
                        "else": {
                            "actions": {
                                "If_start_progress": {
                                    "runAfter": {},
                                    "type": "If",
                                    "expression": should_start_progress,
                                    "actions": {
                                        "Mark_in_progress": (
                                            mark_in_progress_action()
                                        )
                                    },
                                    "else": {"actions": {}},
                                },
                                "If_delivery": {
                                    "runAfter": {
                                        "If_start_progress": ["Succeeded"]
                                    },
                                    "type": "If",
                                    "expression": is_delivery,
                                    "actions": {
                                        "Mark_delivery_requested": update_control(
                                            {
                                                "EstadoGannt/Value": "Entregar",
                                                "StatusExcel": "Entregar",
                                                "FechaEntregaSolicitada": (
                                                    "@if("
                                                    f"empty({CONTROL_ITEM}?"
                                                    "['FechaEntregaSolicitada']),"
                                                    "utcNow(),"
                                                    f"{CONTROL_ITEM}?"
                                                    "['FechaEntregaSolicitada']"
                                                    ")"
                                                ),
                                                "SolicitarVersionado": True,
                                                "UltimoTrackingRun": "@utcNow()",
                                            }
                                        ),
                                        "Get_pending_delivery_events": (
                                            open_api_action(
                                                connection=(
                                                    "shared_sharepointonline"
                                                ),
                                                operation="GetItems",
                                                parameters={
                                                    "dataset": SITE_URL,
                                                    "table": QUEUE_LIST_ID,
                                                    "$top": 500,
                                                },
                                                run_after={
                                                    "Mark_delivery_requested": [
                                                        "Succeeded"
                                                    ]
                                                },
                                            )
                                        ),
                                        "Filter_pending_delivery_events": {
                                            "runAfter": {
                                                "Get_pending_delivery_events": [
                                                    "Succeeded"
                                                ]
                                            },
                                            "type": "Query",
                                            "inputs": {
                                                "from": (
                                                    "@body("
                                                    "'Get_pending_delivery_events'"
                                                    ")?['value']"
                                                ),
                                                "where": (
                                                    "@and("
                                                    "equals(item()?['Title'],"
                                                    "triggerBody()?"
                                                    "['{FilenameWithExtension}']),"
                                                    "equals(item()?['EventType'],"
                                                    "'gantt_working_modificado'),"
                                                    "equals(item()?['Estado'],"
                                                    "'Pendiente')"
                                                    ")"
                                                ),
                                            },
                                        },
                                        "If_no_pending_delivery_event": {
                                            "runAfter": {
                                                "Filter_pending_delivery_events": [
                                                    "Succeeded"
                                                ]
                                            },
                                            "type": "If",
                                            "expression": (
                                                "@equals(length(body("
                                                "'Filter_pending_delivery_events'"
                                                ")),0)"
                                            ),
                                            "actions": {
                                                "Create_delivery_event": (
                                                    open_api_action(
                                                        connection=(
                                                            "shared_sharepointonline"
                                                        ),
                                                        operation="PostItem",
                                                        parameters={
                                                            "dataset": SITE_URL,
                                                            "table": QUEUE_LIST_ID,
                                                            "item/Title": FILE_NAME,
                                                            "item/EventType/Value": (
                                                                "gantt_working_modificado"
                                                            ),
                                                            "item/Estado/Value": (
                                                                "Pendiente"
                                                            ),
                                                            "item/Filename": FILE_NAME,
                                                            "item/FileID": (
                                                                "@triggerBody()?"
                                                                "['{Identifier}']"
                                                            ),
                                                            "item/FolderPath": (
                                                                "@triggerBody()?"
                                                                "['{Path}']"
                                                            ),
                                                            "item/FileLink": (
                                                                "@triggerBody()?"
                                                                "['{Link}']"
                                                            ),
                                                            "item/Intentos": "0",
                                                            "item/Notas": (
                                                                "Entrega detectada "
                                                                "por Power Automate."
                                                            ),
                                                        },
                                                    )
                                                )
                                            },
                                            "else": {"actions": {}},
                                        },
                                    },
                                    "else": {"actions": {}},
                                },
                            }
                        },
                    },
                },
                "else": {"actions": {}},
            },
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


def collection_url() -> str:
    return (
        f"{FLOW_API}/providers/Microsoft.ProcessSimple/environments/"
        f"{ENVIRONMENT}/flows?api-version={API_VERSION}"
    )


def item_url(flow_id: str) -> str:
    return (
        f"{FLOW_API}/providers/Microsoft.ProcessSimple/environments/"
        f"{ENVIRONMENT}/flows/{flow_id}?api-version={API_VERSION}"
    )


def deploy(token: str, *, start: bool) -> tuple[str, str]:
    response = requests.get(
        collection_url(),
        headers=flow_headers(token),
        timeout=60,
    )
    response.raise_for_status()
    existing = next(
        (
            flow
            for flow in response.json().get("value", [])
            if flow.get("properties", {}).get("displayName") == FLOW_NAME
        ),
        None,
    )
    flow_id = str(existing.get("name")) if existing else str(uuid.uuid4())
    payload: dict[str, Any] = {
        "properties": {
            "displayName": FLOW_NAME,
            "definition": build_definition(),
            "connectionReferences": connection_references(),
            "state": "Started" if start else "Stopped",
        }
    }
    if existing:
        deployed = requests.patch(
            item_url(flow_id),
            headers=flow_headers(token),
            json=payload,
            timeout=120,
        )
    else:
        payload["name"] = flow_id
        deployed = requests.post(
            collection_url(),
            headers=flow_headers(token),
            json=payload,
            timeout=120,
        )
    if not deployed.ok:
        raise RuntimeError(
            f"Power Automate {deployed.status_code}: {deployed.text[:5000]}"
        )
    body = deployed.json()
    return (
        str(body.get("name") or flow_id),
        str(body.get("properties", {}).get("state") or ""),
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--token-cache", default=".dataverse_token_cache.json")
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--start", action="store_true")
    parser.add_argument(
        "--output",
        default=(
            "power_automate/solution/generated/"
            "PA_S2_GanttWorkingModificado_A_Cola.definition.json"
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
