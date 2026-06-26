const { app, BrowserWindow, shell, Menu, nativeTheme, ipcMain, safeStorage } = require("electron");
const path = require("path");
const outlook = require("./outlook");
const fetchjob = require("./fetchjob");
const imap = require("./imap");

const isDev = !app.isPackaged;

/* ── App menu (macOS app menu only on Mac) ─────────────────────── */
function buildMenu(win) {
  const template = [
    // macOS app menu — Windows/Linux don't have this concept
    ...(process.platform === "darwin"
      ? [{
          label: "JobLink",
          submenu: [
            { role: "about" },
            { type: "separator" },
            { role: "services" },
            { type: "separator" },
            { role: "hide" },
            { role: "hideOthers" },
            { role: "unhide" },
            { type: "separator" },
            { role: "quit" },
          ],
        }]
      : []),
    {
      label: "File",
      submenu: [
        {
          label: "New Application",
          accelerator: "CmdOrCtrl+N",
          click: () => {
            if (win) win.webContents.executeJavaScript("window.__jlNewApp && window.__jlNewApp()");
          },
        },
        { type: "separator" },
        process.platform === "darwin" ? { role: "close" } : { role: "quit" },
      ],
    },
    {
      label: "Edit",
      submenu: [
        { role: "undo" },
        { role: "redo" },
        { type: "separator" },
        { role: "cut" },
        { role: "copy" },
        { role: "paste" },
        { role: "selectAll" },
      ],
    },
    {
      label: "View",
      submenu: [
        ...(isDev ? [{ role: "toggleDevTools" }, { type: "separator" }] : []),
        { role: "resetZoom" },
        { role: "zoomIn" },
        { role: "zoomOut" },
        { type: "separator" },
        { role: "togglefullscreen" },
      ],
    },
    {
      label: "Window",
      submenu: [
        { role: "minimize" },
        { role: "zoom" },
        { type: "separator" },
        { role: "front" },
      ],
    },
  ];

  return Menu.buildFromTemplate(template);
}

/* ── Window factory ────────────────────────────────────────────── */
const isMac = process.platform === "darwin";

function createWindow() {
  const win = new BrowserWindow({
    width: 1180,
    height: 840,
    minWidth: 780,
    minHeight: 600,
    // macOS gets the frameless inset title bar + vibrancy; Windows/Linux
    // keep their standard window chrome so the controls aren't missing.
    ...(isMac
      ? { titleBarStyle: "hiddenInset", trafficLightPosition: { x: 16, y: 16 }, vibrancy: "under-window", visualEffectState: "active" }
      : {}),
    backgroundColor: nativeTheme.shouldUseDarkColors ? "#0F1724" : "#EEF1F4",
    title: "JobLink",
    show: false,                         // wait until ready-to-show to avoid white flash
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  // Avoid white flash on launch
  win.once("ready-to-show", () => win.show());

  if (isDev) {
    win.loadURL("http://localhost:5173");
  } else {
    win.loadFile(path.join(__dirname, "../dist/index.html"));
  }

  // Set menu after window is created so Cmd+N can reference it
  Menu.setApplicationMenu(buildMenu(win));

  // Push dark/light mode changes to the renderer
  nativeTheme.on("updated", () => {
    if (!win.isDestroyed()) {
      win.webContents.send("theme-changed", nativeTheme.shouldUseDarkColors);
      win.setBackgroundColor(nativeTheme.shouldUseDarkColors ? "#0F1724" : "#EEF1F4");
    }
  });

  // Open external links (job boards, LinkedIn, etc.) in Safari / default browser
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith("http") && !url.startsWith("http://localhost")) {
      shell.openExternal(url);
      return { action: "deny" };
    }
    return { action: "allow" };
  });

  win.webContents.on("will-navigate", (event, url) => {
    if (!url.startsWith("file://") && !url.startsWith("http://localhost")) {
      event.preventDefault();
      shell.openExternal(url);
    }
  });
}

/* ── IPC handlers ──────────────────────────────────────────────── */
ipcMain.on("is-dark-mode", (event) => {
  event.returnValue = nativeTheme.shouldUseDarkColors;
});

// Outlook / Microsoft Graph — desktop OAuth (loopback + PKCE)
outlook.register(ipcMain, app, safeStorage);

// Job-posting URL fetcher — pull details from a pasted link
fetchjob.register(ipcMain);

// Generic IMAP mailbox (Gmail / Yahoo / iCloud / AOL / custom)
imap.register(ipcMain, app, safeStorage);

/* ── App lifecycle ─────────────────────────────────────────────── */
app.whenReady().then(() => {
  // On Windows/Linux, force the dark theme so it always matches the Mac
  // look regardless of the OS appearance setting. macOS keeps following
  // the system setting.
  if (process.platform !== "darwin") nativeTheme.themeSource = "dark";
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
