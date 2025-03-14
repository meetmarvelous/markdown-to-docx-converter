const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  convertMdToDocx: (markdown) => ipcRenderer.invoke("convert-md-to-docx", markdown),
});