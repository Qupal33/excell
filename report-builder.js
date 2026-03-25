const path = require("path");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");

const INVALID_SHEET_NAME_PATTERN = /[*?:\\/[\]]/g;

const REPORTS = [
  {
    id: "money-vr-march",
    code: "Отчёт №1",
    name: "Деньги ВР _март",
    description:
      "Сценарий подготовки итогового отчёта на основе плана и свода по поступлению денежных средств.",
    stageLabel: "Черновой контур",
    sources: [
      {
        id: "planFile",
        label: "Файл плана продаж",
        hint: "Excel-файл плана продаж в формате .xlsx, .xlsm или .xls. Например: План_март 2026 предварительный.",
        dialogTitle: "Выберите файл плана продаж"
      },
      {
        id: "cashflowFile",
        label: "Файл свода по поступлениям",
        hint: "Excel-файл свода по поступлению денежных средств в формате .xlsx, .xlsm или .xls. Например: Свод по поступлению денежных средств_март 2026.",
        dialogTitle: "Выберите файл свода по поступлениям"
      }
    ],
    parameters: [
      {
        id: "reportPeriod",
        label: "Период отчёта",
        placeholder: "Например: март 2026",
        defaultValue: "март 2026"
      },
      {
        id: "reportDate",
        label: "Дата свода",
        placeholder: "Например: 18.03.2026",
        defaultValue: "18.03.2026"
      }
    ]
  }
];

function getReports() {
  return REPORTS;
}

function findReport(reportId) {
  const report = REPORTS.find((item) => item.id === reportId);

  if (!report) {
    throw new Error("Неизвестный сценарий отчёта.");
  }

  return report;
}

function normalizeFileName(value) {
  return value
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, " ")
    .trim();
}

function sanitizeWorksheetName(value, fallbackName) {
  const cleaned = String(value || fallbackName || "Лист")
    .replace(INVALID_SHEET_NAME_PATTERN, "_")
    .trim();

  return (cleaned || fallbackName || "Лист").slice(0, 31);
}

function normalizeSourceValue(cell) {
  if (!cell) {
    return null;
  }

  if (cell.f) {
    return {
      formula: cell.f,
      result: cell.v ?? null
    };
  }

  if (cell.t === "d" && cell.v instanceof Date) {
    return cell.v;
  }

  if (cell.t === "e") {
    return cell.w || cell.v || null;
  }

  return cell.v ?? null;
}

function getWorksheetStats(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");

  return {
    rowCount: range.e.r + 1,
    columnCount: range.e.c + 1
  };
}

function getColumnWidths(sheet) {
  return Array.isArray(sheet["!cols"]) ? sheet["!cols"] : [];
}

function getRowHeights(sheet) {
  return Array.isArray(sheet["!rows"]) ? sheet["!rows"] : [];
}

function getMerges(sheet) {
  return Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
}

function loadWorkbook(filePath) {
  const extension = path.extname(filePath).toLowerCase();

  if (![".xlsx", ".xlsm", ".xls"].includes(extension)) {
    throw new Error(
      `Формат ${extension || "файла"} не поддерживается. Используйте Excel-файл .xlsx, .xlsm или .xls.`
    );
  }

  try {
    const workbook = XLSX.readFile(filePath, {
      cellDates: true,
      cellFormula: true,
      cellNF: true,
      dense: false
    });

    return {
      filePath,
      sheetNames: workbook.SheetNames,
      sheets: workbook.Sheets
    };
  } catch (error) {
    const message = String(error?.message || "");

    if (message.includes("Can't find end of central directory")) {
      throw new Error(
        "Файл не удалось прочитать как Excel. Проверьте, что это настоящий файл Excel (.xlsx, .xlsm или .xls), а не переименованный архив или повреждённый файл."
      );
    }

    throw new Error(`Не удалось открыть Excel-файл "${path.basename(filePath)}": ${message}`);
  }
}

function copySheetValues(sourceSheet, targetSheet, sourceFileLabel) {
  const rangeRef = sourceSheet["!ref"] || "A1:A1";
  const range = XLSX.utils.decode_range(rangeRef);

  getColumnWidths(sourceSheet).forEach((column, index) => {
    const targetColumn = targetSheet.getColumn(index + 1);

    if (typeof column?.wch === "number" && Number.isFinite(column.wch)) {
      targetColumn.width = Math.max(column.wch, 10);
    }

    if (column?.hidden) {
      targetColumn.hidden = true;
    }
  });

  getRowHeights(sourceSheet).forEach((row, index) => {
    if (typeof row?.hpt === "number" && Number.isFinite(row.hpt)) {
      targetSheet.getRow(index + 1).height = row.hpt;
    }
  });

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      const sourceCell = sourceSheet[address];
      const targetCell = targetSheet.getCell(row + 1, column + 1);

      targetCell.value = normalizeSourceValue(sourceCell);

      if (sourceCell?.z) {
        targetCell.numFmt = sourceCell.z;
      }
    }
  }

  getMerges(sourceSheet).forEach((merge) => {
    targetSheet.mergeCells(
      merge.s.r + 1,
      merge.s.c + 1,
      merge.e.r + 1,
      merge.e.c + 1
    );
  });

  targetSheet.getCell("A1").note = `Источник: ${sourceFileLabel}`;
}

function applyHeaderStyle(cell) {
  cell.font = {
    bold: true,
    color: { argb: "FFF7F2E6" }
  };
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF204E36" }
  };
  cell.alignment = {
    vertical: "middle",
    horizontal: "left"
  };
  cell.border = {
    top: { style: "thin", color: { argb: "FF173C2A" } },
    left: { style: "thin", color: { argb: "FF173C2A" } },
    bottom: { style: "thin", color: { argb: "FF173C2A" } },
    right: { style: "thin", color: { argb: "FF173C2A" } }
  };
}

function applyBodyStyle(cell) {
  cell.border = {
    top: { style: "thin", color: { argb: "FFD7D1C2" } },
    left: { style: "thin", color: { argb: "FFD7D1C2" } },
    bottom: { style: "thin", color: { argb: "FFD7D1C2" } },
    right: { style: "thin", color: { argb: "FFD7D1C2" } }
  };
  cell.alignment = {
    vertical: "middle",
    horizontal: "left",
    wrapText: true
  };
}

function buildSummarySheet(workbook, report, payload, planWorkbook, cashflowWorkbook) {
  const sheet = workbook.addWorksheet("Свод");

  sheet.columns = [
    { header: "Показатель", key: "label", width: 34 },
    { header: "Значение", key: "value", width: 72 }
  ];

  sheet.getRow(1).height = 22;
  sheet.getRow(1).eachCell(applyHeaderStyle);

  const rows = [
    ["Отчёт", report.name],
    ["Период", payload.parameters.reportPeriod || ""],
    ["Дата свода", payload.parameters.reportDate || ""],
    ["Файл плана", path.basename(payload.files.planFile)],
    ["Файл свода", path.basename(payload.files.cashflowFile)],
    ["Листов в плане", planWorkbook.sheetNames.length],
    ["Листов в своде", cashflowWorkbook.sheetNames.length],
    [
      "Комментарий",
      "Точное заполнение итогового файла по ячейкам будет добавлено после получения карты переноса данных."
    ]
  ];

  rows.forEach((row) => {
    const inserted = sheet.addRow(row);
    inserted.eachCell(applyBodyStyle);
  });

  sheet.addRow([]);
  const detailHeader = sheet.addRow(["Источник", "Лист / строк"]);
  detailHeader.eachCell(applyHeaderStyle);

  planWorkbook.sheetNames.forEach((sheetName) => {
    const stats = getWorksheetStats(planWorkbook.sheets[sheetName]);
    const row = sheet.addRow([
      "План",
      `${sheetName} / строк: ${stats.rowCount}, столбцов: ${stats.columnCount}`
    ]);
    row.eachCell(applyBodyStyle);
  });

  cashflowWorkbook.sheetNames.forEach((sheetName) => {
    const stats = getWorksheetStats(cashflowWorkbook.sheets[sheetName]);
    const row = sheet.addRow([
      "Свод поступлений",
      `${sheetName} / строк: ${stats.rowCount}, столбцов: ${stats.columnCount}`
    ]);
    row.eachCell(applyBodyStyle);
  });

  sheet.views = [{ state: "frozen", ySplit: 1 }];
}

async function buildMoneyVrMarchReport(report, payload) {
  const requiredFiles = ["planFile", "cashflowFile"];

  for (const fileId of requiredFiles) {
    if (!payload.files?.[fileId]) {
      const fileLabel = report.sources.find((item) => item.id === fileId)?.label || fileId;
      throw new Error(`Не выбран обязательный источник: ${fileLabel}.`);
    }
  }

  const planWorkbook = loadWorkbook(payload.files.planFile);
  const cashflowWorkbook = loadWorkbook(payload.files.cashflowFile);

  const outputWorkbook = new ExcelJS.Workbook();
  outputWorkbook.creator = "Codex";
  outputWorkbook.created = new Date();
  outputWorkbook.modified = new Date();

  buildSummarySheet(outputWorkbook, report, payload, planWorkbook, cashflowWorkbook);

  planWorkbook.sheetNames.forEach((sheetName, index) => {
    const targetName = sanitizeWorksheetName(`План_${index + 1}_${sheetName}`, `План_${index + 1}`);
    const target = outputWorkbook.addWorksheet(targetName);
    copySheetValues(planWorkbook.sheets[sheetName], target, path.basename(payload.files.planFile));
  });

  cashflowWorkbook.sheetNames.forEach((sheetName, index) => {
    const targetName = sanitizeWorksheetName(
      `Поступления_${index + 1}_${sheetName}`,
      `Поступления_${index + 1}`
    );
    const target = outputWorkbook.addWorksheet(targetName);
    copySheetValues(
      cashflowWorkbook.sheets[sheetName],
      target,
      path.basename(payload.files.cashflowFile)
    );
  });

  const period = payload.parameters.reportPeriod || "отчет";
  const defaultFileName = `${normalizeFileName(`Деньги ВР_${period}`)}.xlsx`;

  return {
    workbook: outputWorkbook,
    defaultFileName
  };
}

async function generateReport(payload) {
  if (!payload?.reportId) {
    throw new Error("Не передан идентификатор отчёта.");
  }

  const report = findReport(payload.reportId);

  switch (report.id) {
    case "money-vr-march":
      return buildMoneyVrMarchReport(report, payload);
    default:
      throw new Error("Для выбранного отчёта ещё не настроена логика формирования.");
  }
}

module.exports = {
  getReports,
  generateReport
};
