const fs = require("fs/promises");
const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const path = require("path");
const { getReports, generateReport } = require("./report-builder");

const DEFAULT_STATS = {
  totalReportsGenerated: 0
};

function appendDefaultExtension(filePath, defaultExtension) {
  if (!filePath || path.extname(filePath)) {
    return filePath;
  }

  return `${filePath}.${defaultExtension || "xlsx"}`;
}

function getStatsFilePath() {
  return path.join(app.getPath("userData"), "report-stats.json");
}

async function readStats() {
  try {
    const raw = await fs.readFile(getStatsFilePath(), "utf8");
    const parsed = JSON.parse(raw);

    return {
      totalReportsGenerated: Number.isFinite(parsed?.totalReportsGenerated)
        ? Math.max(0, Math.trunc(parsed.totalReportsGenerated))
        : DEFAULT_STATS.totalReportsGenerated
    };
  } catch (error) {
    if (error.code === "ENOENT") {
      return { ...DEFAULT_STATS };
    }

    console.error("Не удалось прочитать статистику отчётов:", error);
    return { ...DEFAULT_STATS };
  }
}

async function writeStats(stats) {
  await fs.mkdir(path.dirname(getStatsFilePath()), { recursive: true });
  await fs.writeFile(getStatsFilePath(), JSON.stringify(stats, null, 2), "utf8");
}

async function incrementReportsGeneratedCount() {
  const currentStats = await readStats();
  const nextStats = {
    totalReportsGenerated: currentStats.totalReportsGenerated + 1
  };

  await writeStats(nextStats);

  return nextStats;
}

function createWindow() {
  const window = new BrowserWindow({
    width: 1360,
    height: 900,
    minWidth: 1160,
    minHeight: 760,
    title: "Финансы продаж АЦБК",
    backgroundColor: "#f4efe5",
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  window.loadFile("index.html");
}

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

ipcMain.handle("reports:list", async () => {
  return getReports();
});

ipcMain.handle("stats:get", async () => {
  return readStats();
});

ipcMain.handle("dialog:pick-file", async (_event, source) => {
  const isDirectory = source?.selectionType === "directory";
  const dialogResult = await dialog.showOpenDialog({
    title: source?.dialogTitle || "Выберите Excel-файл",
    properties: [isDirectory ? "openDirectory" : "openFile"],
    filters: isDirectory
      ? undefined
      : [
          {
            name: "Excel",
            extensions: ["xlsx", "xlsm", "xls"]
          }
        ]
  });

  if (dialogResult.canceled || dialogResult.filePaths.length === 0) {
    return { canceled: true };
  }

  return {
    canceled: false,
    filePath: dialogResult.filePaths[0]
  };
});

ipcMain.handle("report:generate", async (_event, payload) => {
  try {
    const result = await generateReport(payload);
    const saveResult = await dialog.showSaveDialog({
      title: "Сохранение сформированного отчёта",
      defaultPath: result.defaultFileName,
      filters: [
        {
          name: "Excel",
          extensions: ["xlsx", "xlsm", "xls"]
        }
      ]
    });

    if (saveResult.canceled || !saveResult.filePath) {
      return {
        success: false,
        canceled: true
      };
    }

    const outputPath = appendDefaultExtension(
      saveResult.filePath,
      result.defaultExtension
    );

    await result.save(outputPath);
    const stats = await incrementReportsGeneratedCount();

    return {
      success: true,
      savedTo: outputPath,
      totalReportsGenerated: stats.totalReportsGenerated
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
});
