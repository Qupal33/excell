const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const path = require("path");
const { getReports, generateReport } = require("./report-builder");

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

ipcMain.handle("dialog:pick-file", async (_event, source) => {
  const dialogResult = await dialog.showOpenDialog({
    title: source?.dialogTitle || "Выберите Excel-файл",
    properties: ["openFile"],
    filters: [
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
          extensions: ["xlsx"]
        }
      ]
    });

    if (saveResult.canceled || !saveResult.filePath) {
      return {
        success: false,
        canceled: true
      };
    }

    await result.workbook.xlsx.writeFile(saveResult.filePath);

    return {
      success: true,
      savedTo: saveResult.filePath
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
});
