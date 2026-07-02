type GanttStatusResult = {
  status: string;
  valid: boolean;
  message: string;
};

function main(workbook: ExcelScript.Workbook): GanttStatusResult {
  const sheet = workbook.getWorksheet("Gantt");
  if (!sheet) {
    return {
      status: "",
      valid: false,
      message: "No existe la hoja Gantt.",
    };
  }

  const label = String(sheet.getRange("A6").getValue() ?? "").trim();
  if (label !== "Estado general del Gantt") {
    return {
      status: "",
      valid: false,
      message: "Gantt!A6 no contiene la etiqueta esperada.",
    };
  }

  const status = String(sheet.getRange("B6").getValue() ?? "").trim();
  const allowed = ["Actual", "En Progreso", "Entregar"];
  return {
    status,
    valid: allowed.includes(status),
    message: allowed.includes(status)
      ? "Estado leído correctamente."
      : `Estado no reconocido: ${status}`,
  };
}
