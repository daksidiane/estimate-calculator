const { app, BrowserWindow, ipcMain, shell, nativeTheme, net } = require("electron");
const path = require("path");
const fs = require("fs");

let mainWindow = null;

const MAX_DB_BYTES = 20 * 1024 * 1024;
const MAX_HISTORY_ENTRIES = 500;

function userDataFile(name) {
  return path.join(app.getPath("userData"), name);
}

function readBundledDefault() {
  const file = path.join(__dirname, "data", "default_database.json");
  return JSON.parse(fs.readFileSync(file, "utf-8"));
}

function readOrNull(file) {
  try {
    if (fs.existsSync(file)) {
      return JSON.parse(fs.readFileSync(file, "utf-8"));
    }
  } catch (e) {
    // ignore corrupt files
  }
  return null;
}

function createWindow() {
  nativeTheme.themeSource = "dark";
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 980,
    minHeight: 700,
    backgroundColor: "#0b0f16",
    autoHideMenuBar: true,
    icon: path.join(__dirname, "..", "build", "icon.png"),
    title: "Калькулятор конструкций",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  });

  mainWindow.setMenuBarVisibility(false);
  const indexFile = path.join(__dirname, "renderer", "index.html");
  mainWindow.loadFile(indexFile);

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith("https://") || url.startsWith("http://")) {
      shell.openExternal(url);
    }
    return { action: "deny" };
  });

  mainWindow.webContents.on("will-navigate", (event, url) => {
    const fileUrl = require("url").pathToFileURL(indexFile).href;
    if (url !== fileUrl) event.preventDefault();
  });
}

ipcMain.handle("db:load", async () => {
  const saved = readOrNull(userDataFile("database.json"));
  if (saved) return { db: saved, from: "remote" };
  return { db: readBundledDefault(), from: "bundled" };
});

ipcMain.handle("db:save", async (_e, db) => {
  if (!db || typeof db !== "object" || Array.isArray(db)) throw new Error("Invalid database payload");
  const json = JSON.stringify(db, null, 1);
  if (Buffer.byteLength(json, "utf-8") > MAX_DB_BYTES) throw new Error("Database payload too large");
  fs.writeFileSync(userDataFile("database.json"), json, "utf-8");
  return { ok: true };
});

ipcMain.handle("history:load", async () => {
  const h = readOrNull(userDataFile("history.json"));
  return Array.isArray(h) ? h : [];
});

ipcMain.handle("history:save", async (_e, entries) => {
  if (!Array.isArray(entries)) throw new Error("Invalid history payload");
  if (entries.length > MAX_HISTORY_ENTRIES) {
    throw new Error("History payload too large");
  }
  const json = JSON.stringify(entries, null, 1);
  if (Buffer.byteLength(json, "utf-8") > MAX_DB_BYTES) throw new Error("History payload too large");
  fs.writeFileSync(userDataFile("history.json"), json, "utf-8");
  return { ok: true };
});

ipcMain.handle("app:info", async () => {
  return {
    version: app.getVersion(),
    userData: app.getPath("userData"),
    platform: process.platform
  };
});

const YANDEX_PUBLIC_KEY = "";

async function fetchYandexMeta() {
  if (!YANDEX_PUBLIC_KEY) {
    throw new Error("Укажите публичную ссылку на файл прайса в настройках. В этом образце цены берутся из локальной базы.");
  }
  const api =
    "https://cloud-api.yandex.net/v1/disk/public/resources?public_key=" +
    encodeURIComponent(YANDEX_PUBLIC_KEY);
  const res = await net.fetch(api);
  if (!res.ok) throw new Error("Ошибка доступа к Яндекс.Диску (" + res.status + ")");
  const data = await res.json();
  if (!data || !data.file) throw new Error("Ссылка на файл базы не найдена");
  assertYandexDownloadUrl(data.file);
  return data;
}

function assertYandexDownloadUrl(url) {
  let parsed;
  try {
    parsed = new URL(url);
  } catch (e) {
    throw new Error("Некорректная ссылка на файл прайса");
  }
  if (parsed.protocol !== "https:") {
    throw new Error("Скачивание прайса разрешено только по HTTPS");
  }
  const host = parsed.hostname.toLowerCase();
  const allowed =
    host === "downloader.disk.yandex.ru" ||
    host.endsWith(".disk.yandex.ru") ||
    host.endsWith(".disk.yandex.net") ||
    host === "disk.yandex.ru" ||
    host === "disk.yandex.net";
  if (!allowed) throw new Error("Ссылка на файл вне Яндекс.Диска");
}

ipcMain.handle("db:download", async () => {
  const meta = await fetchYandexMeta();
  assertYandexDownloadUrl(meta.file);
  const fileRes = await net.fetch(meta.file);
  if (!fileRes.ok) throw new Error("Ошибка скачивания файла (" + fileRes.status + ")");
  const buf = Buffer.from(await fileRes.arrayBuffer());
  if (buf.length > MAX_DB_BYTES) throw new Error("Файл прайса слишком большой");
  return {
    bytes: buf.length,
    buffer: new Uint8Array(buf),
    name: meta.name,
    drive: {
      modified: meta.modified || null,
      md5: meta.md5 || null,
      size: meta.size || buf.length
    }
  };
});

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});