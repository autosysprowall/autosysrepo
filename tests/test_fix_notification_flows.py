from power_automate.scripts.fix_notification_flows import (
    corrected_assignment_definition,
    corrected_return_definition,
)


def email_action() -> dict:
    return {
        "actions": {
            "Send_an_email_(V2)": {
                "inputs": {
                    "parameters": {
                        "emailMessage/To": "fixed@example.com",
                        "emailMessage/Subject": "literal",
                        "emailMessage/Body": "literal",
                    }
                }
            }
        }
    }


def test_return_flow_uses_expression_for_subject_and_error() -> None:
    corrected = corrected_return_definition(email_action(), "budget-guide")
    parameters = corrected["actions"]["Send_an_email_(V2)"]["inputs"]["parameters"]

    assert parameters["emailMessage/Subject"].startswith("@concat(")
    assert parameters["emailMessage/To"] == "auto.sys@prowallpanama.com"
    assert parameters["emailMessage/Cc"] == ""
    assert "Presupuesto No Válido Proyecto" in parameters["emailMessage/Subject"]
    assert "UltimoError" in parameters["emailMessage/Body"]
    assert "CreatedByEmail" in parameters["emailMessage/Body"]
    assert parameters["emailMessage/Attachments"][-1]["ContentBytes"] == "budget-guide"


def test_assignment_flow_uses_notification_queue_fields() -> None:
    corrected = corrected_assignment_definition(email_action(), "engineer-guide")
    parameters = corrected["actions"]["Send_an_email_(V2)"]["inputs"]["parameters"]

    assert parameters["emailMessage/To"] == "@triggerBody()?['To']"
    assert parameters["emailMessage/Cc"] == "@triggerBody()?['Cc']"
    assert parameters["emailMessage/Subject"] == "@triggerBody()?['Subject']"
    assert parameters["emailMessage/Body"] == "@triggerBody()?['Body']"
    switch = corrected["actions"]["Confirm_delivery_in_control"]
    assert switch["expression"] == "@triggerBody()?['TipoNotificacion']"
    assert {
        case["case"] for case in switch["cases"].values()
    } == {
        "AsignacionGantt",
        "Advertencia1",
        "Advertencia2",
        "Vencimiento",
    }
    assignment = switch["cases"]["Assignment"]["actions"]["Confirm_assignment"]
    assert assignment["inputs"]["parameters"]["item/CorreoAsignacionEnviado"] is True
    assert parameters["emailMessage/Attachments"] == [
        {
            "Name": "Guia_AutoSys_Ingenieros_Residentes_Planta.pdf",
            "ContentBytes": "engineer-guide",
        }
    ]
