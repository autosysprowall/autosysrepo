type SyncResult = {
  code: string;
  previousStatus: string;
  finalStatus: string;
};

function main(
  workbook: ExcelScript.Workbook,
  desiredStatus: string,
  expectedCurrentStatus: string = "",
  allowReplaceEntregar: boolean = false
): SyncResult {
  const allowed = ["Actual", "En Progreso", "Entregar"];
  const desired = desiredStatus.trim();
  const expected = expectedCurrentStatus.trim();

  if (!allowed.includes(desired)) {
    throw new Error(`Estado no permitido: ${desired}`);
  }

  const sheet = workbook.getWorksheet("Gantt");
  if (!sheet) {
    throw new Error("No existe la hoja Gantt.");
  }

  const label = String(sheet.getRange("A6").getValue() ?? "").trim();
  if (label !== "Estado general del Gantt") {
    throw new Error(
      "Gantt!A6 no contiene la etiqueta Estado general del Gantt."
    );
  }

  const statusCell = sheet.getRange("B6");
  statusCell.getDataValidation().setRule({
    list: {
      inCellDropDown: true,
      source: "Actual,En Progreso,Entregar",
    },
  });
  const current = String(statusCell.getValue() ?? "").trim();

  if (current === desired) {
    return {
      code: "YA_SINCRONIZADO",
      previousStatus: current,
      finalStatus: current,
    };
  }

  if (current === "Entregar" && desired !== "Actual") {
    return {
      code: "PROTEGIDO_ENTREGAR",
      previousStatus: current,
      finalStatus: current,
    };
  }

  if (
    current === "Entregar"
    && desired === "Actual"
    && !allowReplaceEntregar
  ) {
    return {
      code: "REQUIERE_VERSION_EXITOSA",
      previousStatus: current,
      finalStatus: current,
    };
  }

  if (
    expected
    && current !== expected
    && !(current === "Entregar" && desired === "Actual")
  ) {
    return {
      code: "CONFLICTO_ESTADO",
      previousStatus: current,
      finalStatus: current,
    };
  }

  statusCell.setValue(desired);
  return {
    code: "ACTUALIZADO",
    previousStatus: current,
    finalStatus: desired,
  };
}
