import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/lupebedoya/Documents/New project/outputs/qa_onboarding";
const outputPath = `${outputDir}/tablero_onboarding_qa.xlsx`;

const topics = [
  "Introducción a PPM y equipos",
  "Validar permisos en computador y solicitar accesos",
  "Historia de la historia - Jira",
  "Procesos y Estándares de levantamiento de defectos",
  "Zephyr",
  "Playwright",
  "Apidog para pruebas de servicios",
  "TestSigma",
  "Logs AWS",
  "Base de datos y Algolia",
];

const statusOptions = ["Pendiente", "Agendada", "En curso", "Completada", "No aplica"];
const priorityOptions = ["Alta", "Media", "Baja"];
const modalityOptions = ["Presencial", "Virtual", "Hibrida"];

function setFill(range, color) {
  range.format.fill.color = color;
}

function setFont(range, options) {
  range.format.font = options;
}

function setBorder(range, preset = "outside") {
  range.format.borders = { preset, style: "thin", color: "#D8DEE9" };
}

function styleTitle(range) {
  setFill(range, "#1F4E79");
  setFont(range, { bold: true, color: "#FFFFFF", size: 16 });
  range.format.horizontalAlignment = "left";
  range.format.verticalAlignment = "middle";
}

function styleSection(range) {
  setFill(range, "#EAF2F8");
  setFont(range, { bold: true, color: "#1F4E79" });
  setBorder(range, "outside");
}

function styleHeader(range) {
  setFill(range, "#305496");
  setFont(range, { bold: true, color: "#FFFFFF" });
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "middle";
  range.format.wrapText = true;
  setBorder(range, "all");
}

function styleInput(range) {
  setFill(range, "#FFF2CC");
  setBorder(range, "all");
}

function styleTableBody(range) {
  range.format.verticalAlignment = "top";
  range.format.wrapText = true;
  setBorder(range, "all");
}

function applyStatusValidation(sheet, range) {
  sheet.getRange(range).dataValidation = { rule: { type: "list", values: statusOptions } };
}

function applyPriorityValidation(sheet, range) {
  sheet.getRange(range).dataValidation = { rule: { type: "list", values: priorityOptions } };
}

function applyModalityValidation(sheet, range) {
  sheet.getRange(range).dataValidation = { rule: { type: "list", values: modalityOptions } };
}

function applyStatusFormatting(range) {
  setFill(range, "#FFFFFF");
  setFont(range, { color: "#1F1F1F" });
}

const workbook = Workbook.create();
const dashboard = workbook.worksheets.add("Resumen");
const form = workbook.worksheets.add("Formulario Nuevo QA");
const board = workbook.worksheets.add("Tablero Onboarding");
const meetings = workbook.worksheets.add("Reuniones");
const lists = workbook.worksheets.add("Listas");

for (const sheet of [dashboard, form, board, meetings, lists]) {
  sheet.showGridLines = false;
}

dashboard.getRange("A1:J1").merge();
dashboard.getRange("A1").values = [["Tablero de onboarding QA"]];
styleTitle(dashboard.getRange("A1:J1"));
dashboard.getRange("A3:B5").values = [
  ["QAs registrados", "=COUNTA('Tablero Onboarding'!A4:A205)"],
  ["Onboarding completados", "=COUNTIF('Tablero Onboarding'!F4:F205,\"Completada\")"],
  ["Reuniones pendientes", "=COUNTIF(Reuniones!I4:I205,\"Pendiente\")"],
];
styleSection(dashboard.getRange("A3:A5"));
dashboard.getRange("B3:B5").format.font = { bold: true, color: "#1F4E79", size: 14 };
setBorder(dashboard.getRange("A3:B5"), "all");

dashboard.getRange("D3:J3").values = [["Nombre", "Fecha ingreso", "Proyecto", "Lider QA", "% avance", "Estado", "Pendientes"]];
styleHeader(dashboard.getRange("D3:J3"));
dashboard.getRange("D4:J23").formulas = Array.from({ length: 20 }, (_, index) => {
  const sourceRow = index + 4;
  return [
    `=IF('Tablero Onboarding'!A${sourceRow}=\"\",\"\",'Tablero Onboarding'!A${sourceRow})`,
    `=IF('Tablero Onboarding'!A${sourceRow}=\"\",\"\",'Tablero Onboarding'!B${sourceRow})`,
    `=IF('Tablero Onboarding'!A${sourceRow}=\"\",\"\",'Tablero Onboarding'!C${sourceRow})`,
    `=IF('Tablero Onboarding'!A${sourceRow}=\"\",\"\",'Tablero Onboarding'!E${sourceRow})`,
    `=IF('Tablero Onboarding'!A${sourceRow}=\"\",\"\",'Tablero Onboarding'!G${sourceRow})`,
    `=IF('Tablero Onboarding'!A${sourceRow}=\"\",\"\",'Tablero Onboarding'!F${sourceRow})`,
    `=IF('Tablero Onboarding'!A${sourceRow}=\"\",\"\",'Tablero Onboarding'!R${sourceRow})`,
  ];
});
styleTableBody(dashboard.getRange("D4:J23"));
dashboard.getRange("E4:E23").setNumberFormat("yyyy-mm-dd");
dashboard.getRange("H4:H23").setNumberFormat("0%");
applyStatusFormatting(dashboard.getRange("I4:I23"));

form.getRange("A1:H1").merge();
form.getRange("A1").values = [["Formulario de ingreso QA"]];
styleTitle(form.getRange("A1:H1"));
form.getRange("A3:B13").values = [
  ["Nombre completo", ""],
  ["Correo corporativo", ""],
  ["Fecha de ingreso", ""],
  ["Proyecto", ""],
  ["Equipo / celula", ""],
  ["Lider QA", ""],
  ["Mentor / buddy", ""],
  ["Cargo", ""],
  ["Modalidad", ""],
  ["Prioridad de onboarding", ""],
  ["Observaciones iniciales", ""],
];
styleSection(form.getRange("A3:A13"));
styleInput(form.getRange("B3:B13"));
form.getRange("B5").setNumberFormat("yyyy-mm-dd");
applyModalityValidation(form, "B11:B11");
applyPriorityValidation(form, "B12:B12");

form.getRange("D3:H3").values = [["Capacitacion", "Responsable", "Fecha propuesta", "Hora", "Estado"]];
styleHeader(form.getRange("D3:H3"));
form.getRange("D4:H13").values = topics.map((topic) => [topic, "", "", "", "Pendiente"]);
styleTableBody(form.getRange("D4:H13"));
form.getRange("F4:F13").setNumberFormat("yyyy-mm-dd");
applyStatusValidation(form, "H4:H13");
applyStatusFormatting(form.getRange("H4:H13"));

board.getRange("A1:R1").merge();
board.getRange("A1").values = [["Tablero general de onboarding QA"]];
styleTitle(board.getRange("A1:R1"));
board.getRange("A3:R3").values = [[
  "Nombre",
  "Fecha ingreso",
  "Proyecto",
  "Equipo",
  "Lider QA",
  "Estado general",
  "% avance",
  ...topics,
  "Pendientes",
]];
styleHeader(board.getRange("A3:R3"));
board.getRange("A4:R4").values = [
  ["Ejemplo QA", new Date("2026-07-01"), "Proyecto ejemplo", "Equipo ejemplo", "Líder QA", "En curso", "", ...Array(10).fill("Pendiente"), ""],
];
board.getRange("G4:G205").formulas = Array.from({ length: 202 }, (_, index) => {
  const row = index + 4;
  return [`=IF(COUNTA(H${row}:Q${row})=0,\"\",COUNTIF(H${row}:Q${row},\"Completada\")/COUNTA(H${row}:Q${row}))`];
});
board.getRange("R4:R205").formulas = Array.from({ length: 202 }, (_, index) => {
  const row = index + 4;
  return [`=IF(COUNTA(H${row}:Q${row})=0,\"\",COUNTIF(H${row}:Q${row},\"Pendiente\"))`];
});
styleTableBody(board.getRange("A4:R205"));
board.getRange("B4:B205").setNumberFormat("yyyy-mm-dd");
board.getRange("G4:G205").setNumberFormat("0%");
applyStatusValidation(board, "F4:F205");
applyStatusValidation(board, "H4:Q205");
applyStatusFormatting(board.getRange("F4:F205"));
applyStatusFormatting(board.getRange("H4:Q205"));
board.getRange("A3:R205").format.autofitColumns();
board.getRange("A3:R205").format.autofitRows();
board.freezePanes.freezeRows(3);

meetings.getRange("A1:J1").merge();
meetings.getRange("A1").values = [["Agenda de reuniones y capacitaciones"]];
styleTitle(meetings.getRange("A1:J1"));
meetings.getRange("A3:J3").values = [[
  "Nombre QA",
  "Proyecto",
  "Capacitacion",
  "Objetivo",
  "Responsable",
  "Fecha",
  "Hora",
  "Modalidad",
  "Estado",
  "Link / sala",
]];
styleHeader(meetings.getRange("A3:J3"));
meetings.getRange("A4:J13").values = topics.map((topic) => [
  "",
  "",
  topic,
  topic === "Validar permisos en computador y solicitar accesos"
    ? "Confirmar ambiente, herramientas y accesos necesarios"
    : `Capacitacion de ${topic}`,
  "",
  "",
  "",
  "Virtual",
  "Pendiente",
  "",
]);
styleTableBody(meetings.getRange("A4:J205"));
meetings.getRange("F4:F205").setNumberFormat("yyyy-mm-dd");
applyModalityValidation(meetings, "H4:H205");
applyStatusValidation(meetings, "I4:I205");
applyStatusFormatting(meetings.getRange("I4:I205"));
meetings.freezePanes.freezeRows(3);

lists.getRange("A1:D1").values = [["Estados", "Prioridades", "Modalidades", "Capacitaciones"]];
styleHeader(lists.getRange("A1:D1"));
lists.getRange("A2:A6").values = statusOptions.map((value) => [value]);
lists.getRange("B2:B4").values = priorityOptions.map((value) => [value]);
lists.getRange("C2:C4").values = modalityOptions.map((value) => [value]);
lists.getRange("D2:D11").values = topics.map((value) => [value]);
styleTableBody(lists.getRange("A2:D11"));

for (const sheet of [dashboard, form, board, meetings, lists]) {
  sheet.getUsedRange(true).format.font = { name: "Aptos", size: 10 };
}

dashboard.getRange("A1:J1").format.font = { name: "Aptos Display", bold: true, color: "#FFFFFF", size: 16 };
form.getRange("A1:H1").format.font = { name: "Aptos Display", bold: true, color: "#FFFFFF", size: 16 };
board.getRange("A1:R1").format.font = { name: "Aptos Display", bold: true, color: "#FFFFFF", size: 16 };
meetings.getRange("A1:J1").format.font = { name: "Aptos Display", bold: true, color: "#FFFFFF", size: 16 };

dashboard.getRange("A:J").format.autofitColumns();
form.getRange("A:H").format.autofitColumns();
meetings.getRange("A:J").format.autofitColumns();
lists.getRange("A:D").format.autofitColumns();

dashboard.getRange("A1:J1").format.rowHeight = 30;
form.getRange("A1:H1").format.rowHeight = 30;
board.getRange("A1:R1").format.rowHeight = 30;
meetings.getRange("A1:J1").format.rowHeight = 30;

const dashboardCheck = await workbook.inspect({
  kind: "table",
  range: "Resumen!A1:J23",
  include: "values,formulas",
  tableMaxRows: 24,
  tableMaxCols: 10,
  maxChars: 6000,
});
console.log(dashboardCheck.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 4000,
});
console.log(errors.ndjson);

for (const [sheetName, range] of [
  ["Resumen", "A1:J23"],
  ["Formulario Nuevo QA", "A1:H13"],
  ["Tablero Onboarding", "A1:R18"],
  ["Reuniones", "A1:J18"],
  ["Listas", "A1:D11"],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  const previewBytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(`${outputDir}/${sheetName.replaceAll(" ", "_").toLowerCase()}_preview.png`, previewBytes);
}

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${outputPath}`);
