const fs = require("fs/promises");
const path = require("path");
const XLSX = require("xlsx");

const INVALID_SHEET_NAME_PATTERN = /[*?:\\/[\]]/g;
const SOURCE_SHEET_NAME = "За неделю";
const EXCEL_EXTENSIONS = new Set([".xlsx", ".xlsm", ".xls"]);
const MONTH_NAMES = [
  "январь",
  "февраль",
  "март",
  "апрель",
  "май",
  "июнь",
  "июль",
  "август",
  "сентябрь",
  "октябрь",
  "ноябрь",
  "декабрь"
];

const REPORTS = [
  {
    id: "monthly-plan-report",
    code: "Отчёт №1",
    name: "План продаж по месяцу",
    description:
      "Берёт лист \"За неделю\" из выбранного плана продаж, находит итоговый шаблон по месяцу выбранной даты и записывает данные в лист месяца итогового документа.",
    stageLabel: "Простой сценарий",
    sources: [
      {
        id: "planFile",
        label: "Файл плана продаж",
        hint:
          "Excel-файл плана продаж в формате .xlsx, .xlsm или .xls. Из него будет скопирован лист \"За неделю\".",
        dialogTitle: "Выберите файл плана продаж"
      },
      {
        id: "templateDirectory",
        label: "Папка с итоговыми шаблонами",
        hint:
          "Выберите папку, где лежат месячные итоговые файлы. Приложение само найдёт шаблон, в имени которого есть нужный месяц из выбранной даты.",
        dialogTitle: "Выберите папку с итоговыми шаблонами",
        selectionType: "directory"
      }
    ],
    parameters: [
      {
        id: "reportDate",
        label: "Дата отчёта",
        placeholder: "Например: 18.03.2026",
        defaultValue: "2026-03-18",
        inputType: "date"
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
  return String(value || "")
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

function capitalizeFirstLetter(value) {
  if (!value) {
    return "";
  }

  return value.charAt(0).toUpperCase() + value.slice(1);
}

function parseReportDate(value) {
  const normalizedValue = String(value || "").trim();

  if (!normalizedValue) {
    throw new Error("Укажите дату отчёта.");
  }

  let year;
  let month;
  let day;

  const dotMatch = normalizedValue.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);

  if (dotMatch) {
    day = Number(dotMatch[1]);
    month = Number(dotMatch[2]);
    year = Number(dotMatch[3]);
  } else {
    const isoMatch = normalizedValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);

    if (!isoMatch) {
      throw new Error(
        "Неверный формат даты. Используйте dd.mm.yyyy или yyyy-mm-dd."
      );
    }

    year = Number(isoMatch[1]);
    month = Number(isoMatch[2]);
    day = Number(isoMatch[3]);
  }

  const reportDate = new Date(year, month - 1, day);

  if (
    Number.isNaN(reportDate.getTime()) ||
    reportDate.getFullYear() !== year ||
    reportDate.getMonth() !== month - 1 ||
    reportDate.getDate() !== day
  ) {
    throw new Error("Некорректная дата отчёта.");
  }

  return reportDate;
}

function getReportPeriodInfo(reportDateValue) {
  const reportDate = parseReportDate(reportDateValue);
  const monthName = MONTH_NAMES[reportDate.getMonth()];
  const year = reportDate.getFullYear();

  return {
    reportDate,
    monthName,
    titleMonthName: capitalizeFirstLetter(monthName),
    year,
    sheetName: `${monthName} ${year}`,
    displayDate: `${String(reportDate.getDate()).padStart(2, "0")}.${String(
      reportDate.getMonth() + 1
    ).padStart(2, "0")}.${reportDate.getFullYear()}`
  };
}

function loadWorkbook(filePath) {
  const extension = path.extname(filePath).toLowerCase();

  if (!EXCEL_EXTENSIONS.has(extension)) {
    throw new Error(
      `Формат ${extension || "файла"} не поддерживается. Используйте Excel-файл .xlsx, .xlsm или .xls.`
    );
  }

  try {
    const workbook = XLSX.readFile(filePath, {
      cellDates: true,
      cellFormula: true,
      cellNF: true,
      cellStyles: true,
      bookVBA: true,
      dense: false
    });

    return workbook;
  } catch (error) {
    const message = String(error?.message || "");

    if (message.includes("Can't find end of central directory")) {
      throw new Error(
        "Файл не удалось прочитать как Excel. Проверьте, что это настоящий файл Excel (.xlsx, .xlsm или .xls), а не переименованный архив или повреждённый файл."
      );
    }

    throw new Error(
      `Не удалось открыть Excel-файл "${path.basename(filePath)}": ${message}`
    );
  }
}

function normalizeSearchValue(value) {
  return String(value || "")
    .toLowerCase()
    .replaceAll("ё", "е");
}

async function resolveTemplateFile(templateDirectory, periodInfo) {
  const directoryPath = String(templateDirectory || "").trim();

  if (!directoryPath) {
    throw new Error("Не выбрана папка с итоговыми шаблонами.");
  }

  let entries;

  try {
    entries = await fs.readdir(directoryPath, { withFileTypes: true });
  } catch (error) {
    throw new Error(
      `Не удалось прочитать папку с шаблонами "${path.basename(directoryPath) || directoryPath}": ${error.message}`
    );
  }

  const monthToken = normalizeSearchValue(periodInfo.monthName);
  const yearToken = String(periodInfo.year);

  const candidates = entries
    .filter((entry) => entry.isFile())
    .map((entry) => {
      const extension = path.extname(entry.name).toLowerCase();
      const normalizedName = normalizeSearchValue(entry.name);

      if (!EXCEL_EXTENSIONS.has(extension) || !normalizedName.includes(monthToken)) {
        return null;
      }

      let score = 10;

      if (normalizedName.includes(yearToken)) {
        score += 5;
      }

      if (normalizedName.startsWith(monthToken)) {
        score += 2;
      }

      return {
        filePath: path.join(directoryPath, entry.name),
        name: entry.name,
        score
      };
    })
    .filter(Boolean)
    .sort((left, right) => {
      if (right.score !== left.score) {
        return right.score - left.score;
      }

      return left.name.localeCompare(right.name, "ru");
    });

  if (candidates.length === 0) {
    throw new Error(
      `В папке "${path.basename(directoryPath) || directoryPath}" не найден шаблон Excel с месяцем "${periodInfo.titleMonthName}".`
    );
  }

  return candidates[0].filePath;
}

function cloneValue(value) {
  if (value instanceof Date) {
    return new Date(value.getTime());
  }

  if (Array.isArray(value)) {
    return value.map((item) => cloneValue(item));
  }

  if (value && typeof value === "object") {
    const cloned = {};

    for (const [key, nestedValue] of Object.entries(value)) {
      cloned[key] = cloneValue(nestedValue);
    }

    return cloned;
  }

  return value;
}

function cloneSheet(sourceSheet) {
  const targetSheet = {};

  for (const [key, value] of Object.entries(sourceSheet || {})) {
    targetSheet[key] = cloneValue(value);
  }

  return targetSheet;
}

function replaceWorksheet(workbook, sheetName, sheet) {
  const hasSheet = workbook.SheetNames.includes(sheetName);
  workbook.Sheets[sheetName] = sheet;

  if (!hasSheet) {
    workbook.SheetNames.push(sheetName);
  }
}

function getBookType(filePath) {
  const extension = path.extname(filePath).toLowerCase();

  if (!EXCEL_EXTENSIONS.has(extension)) {
    throw new Error(
      "Неверное расширение итогового файла. Используйте .xlsx, .xlsm или .xls."
    );
  }

  return extension.slice(1);
}

function getDefaultOutputFileName(periodInfo, templateFilePath) {
  const templateExtension = path.extname(templateFilePath).toLowerCase() || ".xlsx";
  const templateBaseName = path.basename(templateFilePath, templateExtension).trim();
  const fallbackBaseName = normalizeFileName(
    `Итог_${periodInfo.titleMonthName}_${periodInfo.year}`
  );

  return `${normalizeFileName(templateBaseName || fallbackBaseName)}${templateExtension}`;
}

async function buildMonthlyPlanReport(_report, payload) {
  if (!payload.files?.planFile) {
    throw new Error("Не выбран обязательный источник: файл плана продаж.");
  }

  if (!payload.files?.templateDirectory) {
    throw new Error("Не выбрана папка с итоговыми шаблонами.");
  }

  const periodInfo = getReportPeriodInfo(payload.parameters?.reportDate);
  const targetSheetName = sanitizeWorksheetName(periodInfo.sheetName, "Отчёт");
  const planWorkbook = loadWorkbook(payload.files.planFile);
  const sourceSheet = planWorkbook.Sheets?.[SOURCE_SHEET_NAME];

  if (!sourceSheet) {
    throw new Error(
      `В файле "${path.basename(payload.files.planFile)}" не найден лист "${SOURCE_SHEET_NAME}".`
    );
  }

  const templateFilePath = await resolveTemplateFile(
    payload.files.templateDirectory,
    periodInfo
  );
  const outputWorkbook = loadWorkbook(templateFilePath);
  const targetSheet = cloneSheet(sourceSheet);

  replaceWorksheet(outputWorkbook, targetSheetName, targetSheet);

  return {
    defaultFileName: getDefaultOutputFileName(periodInfo, templateFilePath),
    defaultExtension: getBookType(templateFilePath),
    save: async (filePath) => {
      XLSX.writeFile(outputWorkbook, filePath, {
        bookType: getBookType(filePath),
        bookVBA: true,
        cellStyles: true
      });
    }
  };
}

async function generateReport(payload) {
  if (!payload?.reportId) {
    throw new Error("Не передан идентификатор отчёта.");
  }

  const report = findReport(payload.reportId);

  switch (report.id) {
    case "monthly-plan-report":
      return buildMonthlyPlanReport(report, payload);
    default:
      throw new Error(
        "Для выбранного отчёта ещё не настроена логика формирования."
      );
  }
}

module.exports = {
  getReports,
  generateReport,
  parseReportDate,
  getReportPeriodInfo,
  resolveTemplateFile
};
