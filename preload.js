const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("reportApp", {
  listReports: () => ipcRenderer.invoke("reports:list"),
  pickFile: (source) => ipcRenderer.invoke("dialog:pick-file", source),
  generateReport: (payload) => ipcRenderer.invoke("report:generate", payload)
});
