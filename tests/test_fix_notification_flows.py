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
    corrected = corrected_return_definition(email_action())
    parameters = corrected["actions"]["Send_an_email_(V2)"]["inputs"]["parameters"]

    assert parameters["emailMessage/Subject"].startswith("@concat(")
    assert "Filename" in parameters["emailMessage/Subject"]
    assert "UltimoError" in parameters["emailMessage/Body"]


def test_assignment_flow_uses_notification_queue_fields() -> None:
    corrected = corrected_assignment_definition(email_action())
    parameters = corrected["actions"]["Send_an_email_(V2)"]["inputs"]["parameters"]

    assert parameters["emailMessage/To"] == "@triggerBody()?['To']"
    assert parameters["emailMessage/Cc"] == "@triggerBody()?['Cc']"
    assert parameters["emailMessage/Subject"] == "@triggerBody()?['Subject']"
    assert parameters["emailMessage/Body"] == "@triggerBody()?['Body']"
