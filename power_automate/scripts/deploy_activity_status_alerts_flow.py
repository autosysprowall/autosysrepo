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


FLOW_NAME = "PA_S3_MonitorearEstadosActividades"
SITE_URL = "https://sciprowall.sharepoint.com/sites/PROYECTOSPROWALL"
CONTROL_LIST_ID = "afe5544b-3f60-40e4-81a0-01e86920f5f2"
WORKBOOK_DRIVE_ID = (
    "b!gT--atFwrEiNN1Dm9ygu7gAK2MbJi0RKuVaPM52nyYJBRIiNpNq1Rb608Jq90UM6"
)
ACTIVITY_SCRIPT_ID = (
    "ms-officescript%3A%2F%2Fonedrive_business_itemlink%2F"
    "01DOEBQSO45Q4J7LT44FA2VM7CZW75R4EW"
)
SHAREPOINT_CONNECTION = (
    "shared-sharepointonl-1473fe25-8e08-4ffa-86c4-606465f96830"
)
EXCEL_CONNECTION = "a607013d36214a7288b7cec8cf4d955f"
OUTLOOK_CONNECTION = (
    "shared-office365-d769f58f-56b3-4b9a-8e38-437eafa72c89"
)
CONTROL = "items('For_each_control')"
RESULT = "outputs('Read_activity_status')?['body/result']"
LAST_ALERT_INTERNAL = "UltimaAlertaActividadesFingerpri"


def open_api_action(
    connection: str,
    operation: str,
    parameters: dict[str, Any],
    *,
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
        "id": f"@{CONTROL}?['ID']",
        "item/Title": f"@{CONTROL}?['Title']",
    }
    parameters.update(
        {f"item/{name}": value for name, value in fields.items()}
    )
    return open_api_action(
        "shared_sharepointonline",
        "PatchItem",
        parameters,
        run_after=run_after,
    )


def build_definition() -> dict[str, Any]:
    plan = f"{RESULT}?['planningFingerprint']"
    activity = f"{RESULT}?['activityStatusFingerprint']"
    alert = f"{RESULT}?['alertFingerprint']"
    overdue = f"{RESULT}?['overdueCount']"
    summary = f"{RESULT}?['summary']"
    stored_plan = f"{CONTROL}?['PlanificacionFingerprint']"
    stored_activity = f"{CONTROL}?['EstadoActividadesFingerprint']"
    stored_alert = f"{CONTROL}?['{LAST_ALERT_INTERNAL}']"
    status_only = (
        "@and("
        f"not(empty({stored_plan})),"
        f"equals({plan},{stored_plan}),"
        f"not(equals({activity},{stored_activity}))"
        ")"
    )
    new_alert = (
        "@and("
        f"not(empty({alert})),"
        f"not(equals({alert},{stored_alert}))"
        ")"
    )
    trackable = (
        "@and("
        f"not(empty({CONTROL}?['GanttWorkingIdentifier'])),"
        "not(equals(toLower(coalesce("
        f"{CONTROL}?['EstadoGantt']?['Value'],"
        f"{CONTROL}?['EstadoGannt']?['Value'],"
        f"{CONTROL}?['EstadoGantt'],"
        f"{CONTROL}?['EstadoGannt'],'')),"
        "'vencido'))"
        ")"
    )
    tracking_fields = {
        "PlanificacionFingerprint": f"@{plan}",
        "EstadoActividadesFingerprint": f"@{activity}",
        "AlertasActividadesFingerprint": f"@{alert}",
        "ActividadesAtrasadas": f"@{overdue}",
        "ResumenActividadesAtrasadas": f"@{summary}",
        "FechaLecturaActividades": "@utcNow()",
        LAST_ALERT_INTERNAL: (
            f"@if(empty({alert}),'',{stored_alert})"
        ),
        "UltimoErrorActividades": "",
    }
    email_subject = (
        "@concat('Alerta de actividades atrasadas - ',"
        f"coalesce({CONTROL}?['ProyectoID'],''),' ',"
        f"coalesce({CONTROL}?['NombreProyecto'],''))"
    )
    email_body = (
        "@concat("
        "'<p>Se detectaron actividades atrasadas en el Gantt del proyecto "
        "',coalesce("
        f"{CONTROL}?['NombreProyecto'],{CONTROL}?['ProyectoID'],'sin identificar'"
        "),'.</p>',"
        "'<p><strong>Total:</strong> ',string("
        f"{overdue}),'</p>',"
        "'<p>',replace(coalesce("
        f"{summary},''),decodeUriComponent('%0A'),'<br>'),'</p>',"
        "'<p>Actualice la columna Estatus en el mismo archivo WORKING.</p>'"
        ")"
    )
    return {
        "$schema": (
            "https://schema.management.azure.com/providers/"
            "Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#"
        ),
        "contentVersion": "1.0.0.0",
        "parameters": {
            "$connections": {"defaultValue": {}, "type": "Object"},
            "$authentication": {
                "defaultValue": {},
                "type": "SecureObject",
            },
        },
        "triggers": {
            "Every_15_minutes": {
                "type": "Recurrence",
                "recurrence": {
                    "frequency": "Minute",
                    "interval": 15,
                    "timeZone": "SA Pacific Standard Time",
                },
            }
        },
        "actions": {
            "Get_control_items": open_api_action(
                "shared_sharepointonline",
                "GetItems",
                {
                    "dataset": SITE_URL,
                    "table": CONTROL_LIST_ID,
                    "$top": 500,
                },
            ),
            "For_each_control": {
                "runAfter": {"Get_control_items": ["Succeeded"]},
                "type": "Foreach",
                "foreach": "@body('Get_control_items')?['value']",
                "runtimeConfiguration": {
                    "concurrency": {"repetitions": 1}
                },
                "actions": {
                    "If_trackable": {
                        "runAfter": {},
                        "type": "If",
                        "expression": trackable,
                        "actions": {
                            "Read_activity_status": open_api_action(
                                "shared_excelonlinebusiness",
                                "RunScriptProd",
                                {
                                    "source": SITE_URL,
                                    "drive": WORKBOOK_DRIVE_ID,
                                    "file": (
                                        f"@{CONTROL}?"
                                        "['GanttWorkingIdentifier']"
                                    ),
                                    "scriptId": ACTIVITY_SCRIPT_ID,
                                },
                            ),
                            "If_valid_result": {
                                "runAfter": {
                                    "Read_activity_status": ["Succeeded"]
                                },
                                "type": "If",
                                "expression": (
                                    f"@equals({RESULT}?['valid'],true)"
                                ),
                                "actions": {
                                    "If_status_only_change": {
                                        "runAfter": {},
                                        "type": "If",
                                        "expression": status_only,
                                        "actions": {
                                            "Keep_general_status_actual": (
                                                update_control(
                                                    {
                                                        "EstadoGannt/Value": (
                                                            "Actual"
                                                        ),
                                                        "StatusExcelDeseado": (
                                                            "Actual"
                                                        ),
                                                        "EstadoSyncExcel": (
                                                            "Pendiente"
                                                        ),
                                                        "IntentosSyncExcel": 0,
                                                        "UltimoErrorSyncExcel": "",
                                                    }
                                                )
                                            )
                                        },
                                        "else": {"actions": {}},
                                    },
                                    "Update_activity_tracking": update_control(
                                        tracking_fields,
                                        run_after={
                                            "If_status_only_change": [
                                                "Succeeded"
                                            ]
                                        },
                                    ),
                                    "If_new_activity_alert": {
                                        "runAfter": {
                                            "Update_activity_tracking": [
                                                "Succeeded"
                                            ]
                                        },
                                        "type": "If",
                                        "expression": new_alert,
                                        "actions": {
                                            "If_engineer_email": {
                                                "runAfter": {},
                                                "type": "If",
                                                "expression": (
                                                    "@not(empty("
                                                    f"{CONTROL}?"
                                                    "['IngenieroEmail']))"
                                                ),
                                                "actions": {
                                                    "Send_activity_alert": (
                                                        open_api_action(
                                                            "shared_office365",
                                                            "SendEmailV2",
                                                            {
                                                                "emailMessage/To": (
                                                                    f"@{CONTROL}?"
                                                                    "['IngenieroEmail']"
                                                                ),
                                                                "emailMessage/Cc": (
                                                                    f"@{CONTROL}?"
                                                                    "['SupervisoresEmail']"
                                                                ),
                                                                "emailMessage/Subject": (
                                                                    email_subject
                                                                ),
                                                                "emailMessage/Body": (
                                                                    email_body
                                                                ),
                                                                "emailMessage/Importance": (
                                                                    "Normal"
                                                                ),
                                                            },
                                                        )
                                                    ),
                                                    "Confirm_activity_alert": (
                                                        update_control(
                                                            {
                                                                LAST_ALERT_INTERNAL: (
                                                                    f"@{alert}"
                                                                ),
                                                                "FechaUltimoCorreoActividades": (
                                                                    "@utcNow()"
                                                                ),
                                                                "UltimoErrorActividades": (
                                                                    ""
                                                                ),
                                                            },
                                                            run_after={
                                                                "Send_activity_alert": [
                                                                    "Succeeded"
                                                                ]
                                                            },
                                                        )
                                                    ),
                                                },
                                                "else": {
                                                    "actions": {
                                                        "Record_missing_email": (
                                                            update_control(
                                                                {
                                                                    "UltimoErrorActividades": (
                                                                        "IngenieroEmail vacío; "
                                                                        "no se envió alerta."
                                                                    )
                                                                }
                                                            )
                                                        )
                                                    }
                                                },
                                            }
                                        },
                                        "else": {"actions": {}},
                                    },
                                },
                                "else": {
                                    "actions": {
                                        "Record_invalid_result": (
                                            update_control(
                                                {
                                                    "UltimoErrorActividades": (
                                                        f"@{RESULT}?"
                                                        "['message']"
                                                    ),
                                                    "FechaLecturaActividades": (
                                                        "@utcNow()"
                                                    ),
                                                }
                                            )
                                        )
                                    }
                                },
                            },
                            "Record_script_failure": update_control(
                                {
                                    "UltimoErrorActividades": (
                                        "Falló GetGanttActivityStatus."
                                    ),
                                    "FechaLecturaActividades": "@utcNow()",
                                },
                                run_after={
                                    "Read_activity_status": [
                                        "Failed",
                                        "TimedOut",
                                    ]
                                },
                            ),
                        },
                        "else": {"actions": {}},
                    }
                },
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
        "shared_office365": {
            "connectionName": OUTLOOK_CONNECTION,
            "source": "Embedded",
            "id": "/providers/Microsoft.PowerApps/apis/shared_office365",
            "displayName": "Office 365 Outlook",
            "tier": "Standard",
            "apiName": "office365",
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
        response = requests.patch(
            item_url(flow_id),
            headers=flow_headers(token),
            json=payload,
            timeout=120,
        )
    else:
        payload["name"] = flow_id
        response = requests.post(
            collection_url(),
            headers=flow_headers(token),
            json=payload,
            timeout=120,
        )
    if not response.ok:
        raise RuntimeError(
            f"Power Automate {response.status_code}: {response.text[:5000]}"
        )
    body = response.json()
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
            "PA_S3_MonitorearEstadosActividades.definition.json"
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
