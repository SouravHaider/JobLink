const { contextBridge, ipcRenderer } = require("electron");

// Expose a safe, narrow API to the renderer (App.jsx)
contextBridge.exposeInMainWorld("electron", {
  // Dark mode — read initial state and listen for changes
  isDarkMode: () => ipcRenderer.sendSync("is-dark-mode"),
  onThemeChange: (cb) => {
    ipcRenderer.on("theme-changed", (_event, isDark) => cb(isDark));
  },
});

// Outlook bridge — desktop OAuth runs in the main process
contextBridge.exposeInMainWorld("outlook", {
  status: () => ipcRenderer.invoke("outlook:status"),
  connect: (args) => ipcRenderer.invoke("outlook:connect", args),
  disconnect: () => ipcRenderer.invoke("outlook:disconnect"),
  scan: (args) => ipcRenderer.invoke("outlook:scan", args),
});

// Job-page fetcher — reads a posting URL and extracts details
contextBridge.exposeInMainWorld("jobfetch", {
  fetch: (url) => ipcRenderer.invoke("jobfetch:fetch", { url }),
});

// Generic IMAP mailbox bridge (Gmail / Yahoo / iCloud / AOL / custom)
contextBridge.exposeInMainWorld("mailbox", {
  list: () => ipcRenderer.invoke("mailbox:list"),
  connect: (creds) => ipcRenderer.invoke("mailbox:connect", creds),
  disconnect: (id) => ipcRenderer.invoke("mailbox:disconnect", { id }),
  scan: (opts) => ipcRenderer.invoke("mailbox:scan", opts),
});
