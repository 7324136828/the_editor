const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const isDevelopment = process.argv.includes('--dev');
const isSmokeTest = process.argv.includes('--smoke-test');
let mainWindow = null;

function windowForEvent(event) {
  const window = BrowserWindow.fromWebContents(event.sender);
  if (!window || event.senderFrame !== window.webContents.mainFrame) return null;
  return window;
}

function registerWindowControls() {
  ipcMain.on('window:minimize', (event) => {
    windowForEvent(event)?.minimize();
  });

  ipcMain.on('window:toggle-maximize', (event) => {
    const window = windowForEvent(event);
    if (!window) return;
    if (window.isMaximized()) window.unmaximize();
    else window.maximize();
  });

  ipcMain.on('window:close', (event) => {
    windowForEvent(event)?.close();
  });
}

function createWindow() {
  const window = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 900,
    minHeight: 600,
    show: false,
    frame: false,
    backgroundColor: '#217346',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  mainWindow = window;

  window.on('closed', () => {
    if (mainWindow === window) mainWindow = null;
  });

  window.once('ready-to-show', () => window.show());

  if (isSmokeTest) {
    window.webContents.once('did-finish-load', () => {
      console.log('Electron smoke test passed: renderer loaded.');
      app.quit();
    });

    window.webContents.once('did-fail-load', (_event, code, description) => {
      console.error(`Electron smoke test failed (${code}): ${description}`);
      app.exit(1);
    });
  }

  window.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('https://') || url.startsWith('mailto:')) {
      void shell.openExternal(url);
    }
    return { action: 'deny' };
  });

  window.webContents.on('will-navigate', (event, url) => {
    const allowedUrl = isDevelopment
      ? 'http://127.0.0.1:5173/'
      : pathToFileURL(path.join(__dirname, '..', 'dist', 'index.html')).href;

    if (url !== allowedUrl) event.preventDefault();
  });

  if (isDevelopment) {
    void window.loadURL('http://127.0.0.1:5173');
  } else {
    void window.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }
}

app.whenReady().then(() => {
  registerWindowControls();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
