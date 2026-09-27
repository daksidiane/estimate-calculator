const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  loadDb: () => ipcRenderer.invoke("db:load"),
  saveDb: (db) => ipcRenderer.invoke("db:save", db),
  downloadDb: () => ipcRenderer.invoke("db:download"),
  loadHistory: () => ipcRenderer.invoke("history:load"),
  saveHistory: (entries) => ipcRenderer.invoke("history:save", entries),
  appInfo: () => ipcRenderer.invoke("app:info")
});