type ActivityAlert = {
  row: number;
  activity: string;
  status: string;
  startDate: string;
  endDate: string;
  reason: string;
};

type ActivityStatusResult = {
  valid: boolean;
  message: string;
  generalStatus: string;
  planningFingerprint: string;
  activityStatusFingerprint: string;
  alertFingerprint: string;
  overdueCount: number;
  overdueActivities: ActivityAlert[];
  summary: string;
};

function normalize(value: unknown): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function excelDateToIso(value: string | number | boolean): string {
  if (typeof value === "number") {
    const milliseconds =
      Date.UTC(1899, 11, 30) + Math.round(value * 86400000);
    return new Date(milliseconds).toISOString().slice(0, 10);
  }
  const text = String(value ?? "").trim();
  if (!text) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const dayFirst = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (!dayFirst) return "";
  return `${dayFirst[3]}-${dayFirst[2].padStart(2, "0")}-${dayFirst[1].padStart(2, "0")}`;
}

function fingerprint(value: string): string {
  let hash = 2166136261;
  for (let index = 0; index < value.length; index++) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function main(
  workbook: ExcelScript.Workbook,
  todayIso: string = ""
): ActivityStatusResult {
  const emptyResult = (
    message: string,
    generalStatus: string = ""
  ): ActivityStatusResult => ({
    valid: false,
    message,
    generalStatus,
    planningFingerprint: "",
    activityStatusFingerprint: "",
    alertFingerprint: "",
    overdueCount: 0,
    overdueActivities: [],
    summary: "",
  });
  const sheet = workbook.getWorksheet("Gantt");
  if (!sheet) return emptyResult("No existe la hoja Gantt.");
  const usedRange = sheet.getUsedRange();
  if (!usedRange) return emptyResult("La hoja Gantt está vacía.");
  const values = usedRange.getValues();
  const rowCount = values.length;
  const columnCount = values[0]?.length ?? 0;
  let headerRow = -1;
  let headers: { [key: string]: number } = {};

  for (let row = 0; row < Math.min(rowCount, 30); row++) {
    const candidate: { [key: string]: number } = {};
    for (let column = 0; column < columnCount; column++) {
      const header = normalize(values[row][column]);
      if (header) candidate[header] = column;
    }
    const hasActivity =
      candidate["actividad"] !== undefined ||
      candidate["actividades"] !== undefined;
    const hasStart =
      candidate["fecha de inicio"] !== undefined ||
      candidate["fecha inicio"] !== undefined;
    const hasEnd =
      candidate["fecha de fin"] !== undefined ||
      candidate["fecha fin"] !== undefined ||
      candidate["fecha final"] !== undefined;
    const hasStatus =
      candidate["estatus"] !== undefined ||
      candidate["status"] !== undefined ||
      candidate["estado"] !== undefined;
    if (hasActivity && hasStart && hasEnd && hasStatus) {
      headerRow = row;
      headers = candidate;
      break;
    }
  }

  const generalStatus = String(
    sheet.getRange("B6").getValue() ?? ""
  ).trim();
  if (headerRow < 0) {
    return emptyResult(
      "No se encontraron Actividad, Fecha de Inicio, Fecha de Fin y Estatus.",
      generalStatus
    );
  }

  const findColumn = (names: string[]): number => {
    for (const name of names) {
      if (headers[name] !== undefined) return headers[name];
    }
    return -1;
  };
  const activityColumn = findColumn(["actividad", "actividades"]);
  const startColumn = findColumn(["fecha de inicio", "fecha inicio"]);
  const endColumn = findColumn([
    "fecha de fin",
    "fecha fin",
    "fecha final",
  ]);
  const statusColumn = findColumn(["estatus", "status", "estado"]);
  const planningColumns = [
    activityColumn,
    findColumn(["item", "ítem"]),
    findColumn(["cc"]),
    startColumn,
    endColumn,
    findColumn(["unidad"]),
    findColumn(["cantidad"]),
    findColumn(["costo unitario"]),
    findColumn(["costo total"]),
  ].filter((column) => column >= 0);
  const today =
    todayIso && /^\d{4}-\d{2}-\d{2}$/.test(todayIso)
      ? todayIso
      : new Date().toISOString().slice(0, 10);
  const planningParts: string[] = [];
  const statusParts: string[] = [];
  const alerts: ActivityAlert[] = [];

  for (let row = headerRow + 1; row < rowCount; row++) {
    const activity = String(values[row][activityColumn] ?? "").trim();
    if (!activity) continue;
    const rawStatus = normalize(values[row][statusColumn]);
    let status = "";
    if (rawStatus === "pendiente") status = "Pendiente";
    if (rawStatus === "en progreso") status = "En Progreso";
    if (rawStatus === "completada") status = "Completada";
    const startDate = excelDateToIso(values[row][startColumn]);
    const endDate = excelDateToIso(values[row][endColumn]);
    planningParts.push(
      `${row + 1}|${planningColumns
        .map((column) => String(values[row][column] ?? "").trim())
        .join("|")}`
    );
    statusParts.push(`${row + 1}|${activity}|${status}`);
    let reason = "";
    if (status === "Pendiente" && startDate && today > startDate) {
      reason = "Excedió Fecha de Inicio";
    }
    if (status === "En Progreso" && endDate && today > endDate) {
      reason = "Excedió Fecha de Fin";
    }
    if (reason) {
      alerts.push({
        row: row + 1,
        activity,
        status,
        startDate,
        endDate,
        reason,
      });
    }
  }

  const summary = alerts
    .map(
      (alert) =>
        `Fila ${alert.row}: ${alert.activity} - ${alert.status} - ${alert.reason}`
    )
    .join("\n");
  const alertCanonical = alerts
    .map(
      (alert) =>
        `${alert.row}|${alert.activity}|${alert.status}|${alert.startDate}|${alert.endDate}|${alert.reason}`
    )
    .join("\n");
  return {
    valid: true,
    message: "Estados de actividades leídos correctamente.",
    generalStatus,
    planningFingerprint: fingerprint(planningParts.join("\n")),
    activityStatusFingerprint: fingerprint(statusParts.join("\n")),
    alertFingerprint: alerts.length
      ? fingerprint(alertCanonical)
      : "",
    overdueCount: alerts.length,
    overdueActivities: alerts,
    summary,
  };
}
