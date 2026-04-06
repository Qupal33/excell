const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("reportApp", {
  listReports: () => ipcRenderer.invoke("reports:list"),
  getStats: () => ipcRenderer.invoke("stats:get"),
  pickFile: (source) => ipcRenderer.invoke("dialog:pick-file", source),
  generateReport: (payload) => ipcRenderer.invoke("report:generate", payload)
});
