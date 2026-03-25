const state = {
  reports: [],
  activeReportId: null,
  files: {},
  parameters: {},
  status: {
    type: "",
    text: ""
  },
  busy: false
};

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

function renderHome() {
  return `
    <section class="hero">
      <div class="hero-copy">
        <span class="hero-kicker">АО "Архангельский ЦБК"</span>
        <h1>Формирование финансовых отчётов отдела продаж</h1>
        <p>
          Локальное приложение работает на компьютере пользователя, принимает Excel-файлы,
          собирает итоговый документ и сохраняет его в выбранную папку.
        </p>
      </div>
      <div class="hero-panel">
        <div class="hero-metric">1</div>
        <div>
          <strong>Доступный отчёт</strong>
          <p>Стартовая версия уже содержит сценарий для отчёта "Деньги ВР _март".</p>
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
              <h3>Исходные файлы</h3>
              <p>Для каждого источника выберите Excel-файл с компьютера.</p>
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
                        ${hasFile ? "Заменить файл" : "Выбрать файл"}
                      </button>
                      <div class="file-badge ${hasFile ? "file-badge-ready" : ""}">
                        ${
                          hasFile
                            ? escapeHtml(basename(filePath))
                            : "Файл не выбран"
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
              <p>Эти поля можно уточнять и расширять по мере детализации бизнес-логики.</p>
            </div>
          </div>

          <div class="field-list">
            ${report.parameters
              .map(
                (parameter) => `
                  <label class="field">
                    <span>${escapeHtml(parameter.label)}</span>
                    <input
                      type="text"
                      data-action="set-parameter"
                      data-parameter-id="${escapeHtml(parameter.id)}"
                      value="${escapeHtml(state.parameters[parameter.id] || "")}"
                      placeholder="${escapeHtml(parameter.placeholder || "")}"
                    />
                  </label>
                `
              )
              .join("")}
          </div>

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
      ${report ? renderReportForm(report) : renderHome()}
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
        setStatus("info", `Выбран файл: ${basename(response.filePath)}`);
      }
    });
  });

  document.querySelectorAll("[data-action='set-parameter']").forEach((input) => {
    input.addEventListener("input", () => {
      state.parameters[input.dataset.parameterId] = input.value;
    });
  });

  const generateButton = document.querySelector("[data-action='generate-report']");

  if (generateButton) {
    generateButton.addEventListener("click", handleGenerateReport);
  }
}

async function handleGenerateReport() {
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
    setStatus("success", `Отчёт сохранён: ${response.savedTo}`);
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
  render();
}

window.addEventListener("DOMContentLoaded", init);
