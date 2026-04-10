const state = {
  reports: [],
  activeReportId: null,
  totalReportsGenerated: 0,
  files: {},
  parameters: {},
  status: {
    type: "",
    text: ""
  },
  busy: false
};

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

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function basename(filePath) {
  if (!filePath) {
    return "";
  }

  return filePath.split(/[\\/]/).pop();
}

function getActiveReport() {
  return state.reports.find((report) => report.id === state.activeReportId) || null;
}

function padToTwoDigits(value) {
  return String(value).padStart(2, "0");
}

function getTodayIsoDate() {
  const today = new Date();

  return `${today.getFullYear()}-${padToTwoDigits(today.getMonth() + 1)}-${padToTwoDigits(
    today.getDate()
  )}`;
}

function getDatePreview(value) {
  const normalizedValue = String(value || "").trim();
  const match = normalizedValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const monthIndex = Number(match[2]) - 1;
  const day = Number(match[3]);
  const date = new Date(year, monthIndex, day);

  if (
    Number.isNaN(date.getTime()) ||
    date.getFullYear() !== year ||
    date.getMonth() !== monthIndex ||
    date.getDate() !== day
  ) {
    return null;
  }

  const monthName = MONTH_NAMES[monthIndex];

  return {
    displayDate: `${padToTwoDigits(day)}.${padToTwoDigits(monthIndex + 1)}.${year}`,
    monthName,
    titleMonthName: monthName.charAt(0).toUpperCase() + monthName.slice(1),
    sheetName: `${monthName} ${year}`
  };
}

function getSourceSelectionText(source, hasValue) {
  const isDirectory = source?.selectionType === "directory";

  if (hasValue) {
    return isDirectory ? "Выбрать другую папку" : "Заменить файл";
  }

  return isDirectory ? "Выбрать папку" : "Выбрать файл";
}

function getSourcePlaceholderText(source) {
  return source?.selectionType === "directory" ? "Папка не выбрана" : "Файл не выбран";
}

function getPickedSourceText(source, filePath) {
  const itemType = source?.selectionType === "directory" ? "Папка" : "Файл";
  return `${itemType} выбрана: ${basename(filePath)}`;
}

function setStatus(type, text) {
  state.status = { type, text };
  render();
}

function resetReportState(reportId) {
  state.activeReportId = reportId;
  state.files = {};
  state.parameters = {};
  state.status = { type: "", text: "" };

  const report = getActiveReport();
  if (!report) {
    return;
  }

  for (const parameter of report.parameters) {
    state.parameters[parameter.id] = parameter.defaultValue || "";
  }
}

function renderStatus() {
  if (!state.status.text) {
    return "";
  }

  return `
    <div class="status status-${escapeHtml(state.status.type || "info")}">
      ${escapeHtml(state.status.text)}
    </div>
  `;
}

function renderDatePreview(report) {
  const dateParameter = report.parameters.find((parameter) => parameter.id === "reportDate");

  if (!dateParameter) {
    return "";
  }

  const preview = getDatePreview(state.parameters[dateParameter.id]);

  if (!preview) {
    return `
      <div class="date-preview date-preview-muted">
        <strong>Месяц отчёта</strong>
        <span>Выберите дату через календарь, и приложение сразу подставит нужный месяц и имя листа.</span>
      </div>
    `;
  }

  return `
    <div class="date-preview">
      <strong>${escapeHtml(preview.displayDate)}</strong>
      <span>Месяц отчёта: ${escapeHtml(preview.titleMonthName)}</span>
      <span>Лист итогового файла: ${escapeHtml(preview.sheetName)}</span>
    </div>
  `;
}

function renderHomeWithStats() {
  return `
    <section class="hero">
      <div class="hero-copy">
        <span class="hero-kicker">АО "Архангельский ЦБК"</span>
        <h1>Формирование финансовых отчётов отдела продаж</h1>
        <p>
          Локальное приложение работает на компьютере пользователя, принимает Excel-файлы,
          находит нужный месячный шаблон и сохраняет готовый отчёт в выбранную папку.
        </p>
      </div>
      <div class="hero-panel">
        <div class="hero-metric">${escapeHtml(state.totalReportsGenerated)}</div>
        <div>
          <strong>Сформировано отчётов</strong>
          <p>
            Сейчас доступен сценарий, в котором месяц отчёта определяется по выбранной дате,
            а итоговый шаблон подбирается автоматически по имени месяца.
          </p>
        </div>
      </div>
    </section>

    <section class="section">
      <div class="section-head">
        <div>
          <span class="section-label">Выбор сценария</span>
          <h2>Какой отчёт необходимо сформировать</h2>
        </div>
      </div>
      <div class="report-grid">
        ${state.reports
          .map(
            (report) => `
              <article class="report-card">
                <div class="report-card-top">
                  <span class="report-tag">${escapeHtml(report.code)}</span>
                  <span class="report-stage">${escapeHtml(report.stageLabel)}</span>
                </div>
                <h3>${escapeHtml(report.name)}</h3>
                <p>${escapeHtml(report.description)}</p>
                <ul class="report-meta">
                  ${report.sources
                    .map((source) => `<li>${escapeHtml(source.label)}</li>`)
                    .join("")}
                </ul>
                <button class="primary-button" data-action="open-report" data-report-id="${escapeHtml(
                  report.id
                )}">
                  Открыть форму
                </button>
              </article>
            `
          )
          .join("")}
      </div>
    </section>
  `;
}

function renderReportForm(report) {
  return `
    <section class="section">
      <div class="section-head">
        <div>
          <button class="link-button" data-action="back-home">Назад к выбору отчётов</button>
          <span class="section-label">${escapeHtml(report.code)}</span>
          <h2>${escapeHtml(report.name)}</h2>
          <p class="section-copy">${escapeHtml(report.description)}</p>
        </div>
        <div class="step-pill">Шаг 2 из 2</div>
      </div>

      <div class="form-layout">
        <section class="panel">
          <div class="panel-head">
            <span class="panel-index">01</span>
            <div>
              <h3>Исходные данные</h3>
              <p>Для каждого источника выберите Excel-файл или папку с месячными шаблонами.</p>
            </div>
          </div>

          <div class="source-list">
            ${report.sources
              .map((source) => {
                const filePath = state.files[source.id] || "";
                const hasFile = Boolean(filePath);

                return `
                  <div class="source-item">
                    <div class="source-copy">
                      <strong>${escapeHtml(source.label)}</strong>
                      <span>${escapeHtml(source.hint)}</span>
                    </div>
                    <div class="source-actions">
                      <button
                        class="secondary-button"
                        data-action="pick-file"
                        data-source-id="${escapeHtml(source.id)}"
                      >
                        ${escapeHtml(getSourceSelectionText(source, hasFile))}
                      </button>
                      <div class="file-badge ${hasFile ? "file-badge-ready" : ""}">
                        ${
                          hasFile
                            ? escapeHtml(basename(filePath))
                            : escapeHtml(getSourcePlaceholderText(source))
                        }
                      </div>
                    </div>
                  </div>
                `;
              })
              .join("")}
          </div>
        </section>

        <section class="panel">
          <div class="panel-head">
            <span class="panel-index">02</span>
            <div>
              <h3>Параметры формирования</h3>
              <p>Пока используется один критерий: дата отчёта, по которой определяется месяц и лист итогового документа.</p>
            </div>
          </div>

          <div class="field-list">
            ${report.parameters
              .map(
                (parameter) => `
                  <label class="field">
                    <span>${escapeHtml(parameter.label)}</span>
                    ${
                      parameter.inputType === "date"
                        ? `
                          <div class="date-input-row">
                            <input
                              class="date-input"
                              type="date"
                              data-action="set-parameter"
                              data-parameter-id="${escapeHtml(parameter.id)}"
                              value="${escapeHtml(state.parameters[parameter.id] || "")}"
                              placeholder="${escapeHtml(parameter.placeholder || "")}"
                            />
                            <button
                              type="button"
                              class="secondary-button date-picker-button"
                              data-action="open-date-picker"
                              data-parameter-id="${escapeHtml(parameter.id)}"
                            >
                              Календарь
                            </button>
                            <button
                              type="button"
                              class="ghost-button date-today-button"
                              data-action="set-today"
                              data-parameter-id="${escapeHtml(parameter.id)}"
                            >
                              Сегодня
                            </button>
                          </div>
                        `
                        : `
                          <input
                            type="${escapeHtml(parameter.inputType || "text")}"
                            data-action="set-parameter"
                            data-parameter-id="${escapeHtml(parameter.id)}"
                            value="${escapeHtml(state.parameters[parameter.id] || "")}"
                            placeholder="${escapeHtml(parameter.placeholder || "")}"
                          />
                        `
                    }
                  </label>
                `
              )
              .join("")}
          </div>

          ${renderDatePreview(report)}

          ${renderStatus()}

          <div class="action-row">
            <button class="ghost-button" data-action="back-home">К списку отчётов</button>
            <button class="primary-button" data-action="generate-report" ${
              state.busy ? "disabled" : ""
            }>
              ${state.busy ? "Формирование..." : "Сформировать"}
            </button>
          </div>
        </section>
      </div>
    </section>
  `;
}

function render() {
  const app = document.getElementById("app");
  const report = getActiveReport();

  app.innerHTML = `
    <main class="page-shell">
      <header class="topbar">
        <div class="brand-lockup">
          <div class="brand-mark">АЦБК</div>
          <div>
            <span class="brand-caption">Локальное desktop-приложение</span>
            <strong>Финансы отдела продаж</strong>
          </div>
        </div>
      </header>
      ${report ? renderReportForm(report) : renderHomeWithStats()}
    </main>
  `;

  bindEvents();
}

function bindEvents() {
  document.querySelectorAll("[data-action='open-report']").forEach((button) => {
    button.addEventListener("click", () => {
      resetReportState(button.dataset.reportId);
      render();
    });
  });

  document.querySelectorAll("[data-action='back-home']").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeReportId = null;
      state.status = { type: "", text: "" };
      render();
    });
  });

  document.querySelectorAll("[data-action='pick-file']").forEach((button) => {
    button.addEventListener("click", async () => {
      const report = getActiveReport();
      const source = report?.sources.find((item) => item.id === button.dataset.sourceId);

      if (!source) {
        return;
      }

      const response = await window.reportApp.pickFile(source);

      if (!response.canceled && response.filePath) {
        state.files[source.id] = response.filePath;
        setStatus("info", getPickedSourceText(source, response.filePath));
      }
    });
  });

  document.querySelectorAll("[data-action='set-parameter']").forEach((input) => {
    input.addEventListener("input", () => {
      state.parameters[input.dataset.parameterId] = input.value;
      render();
    });
  });

  document.querySelectorAll("[data-action='open-date-picker']").forEach((button) => {
    button.addEventListener("click", () => {
      const input = document.querySelector(
        `[data-action='set-parameter'][data-parameter-id='${button.dataset.parameterId}']`
      );

      if (!input) {
        return;
      }

      if (typeof input.showPicker === "function") {
        input.showPicker();
        return;
      }

      input.focus();
      input.click();
    });
  });

  document.querySelectorAll("[data-action='set-today']").forEach((button) => {
    button.addEventListener("click", () => {
      state.parameters[button.dataset.parameterId] = getTodayIsoDate();
      render();
    });
  });

  const generateButton = document.querySelector("[data-action='generate-report']");

  if (generateButton) {
    generateButton.addEventListener("click", handleGenerateReportWithStats);
  }
}

async function handleGenerateReportWithStats() {
  const report = getActiveReport();

  if (!report || state.busy) {
    return;
  }

  state.busy = true;
  render();

  const response = await window.reportApp.generateReport({
    reportId: report.id,
    files: state.files,
    parameters: state.parameters
  });

  state.busy = false;

  if (response.success) {
    state.totalReportsGenerated = Number.isFinite(response.totalReportsGenerated)
      ? response.totalReportsGenerated
      : state.totalReportsGenerated;
    setStatus(
      "success",
      `Отчёт сохранён: ${response.savedTo}. Всего сформировано: ${state.totalReportsGenerated}.`
    );
    return;
  }

  if (response.canceled) {
    setStatus("warning", "Сохранение отменено пользователем.");
    return;
  }

  setStatus("error", response.error || "Не удалось сформировать отчёт.");
}

async function init() {
  state.reports = await window.reportApp.listReports();
  const stats = await window.reportApp.getStats();
  state.totalReportsGenerated = Number.isFinite(stats?.totalReportsGenerated)
    ? stats.totalReportsGenerated
    : 0;
  render();
}

window.addEventListener("DOMContentLoaded", init);
