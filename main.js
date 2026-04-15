const { app, BrowserWindow } = require('electron');
const path = require('path');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1100,
    height: 700,
    minWidth: 900,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
    title: 'Visa Form',
    backgroundColor: '#0a0a0a',
  });

  mainWindow.loadURL('http://localhost:3721');

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

function startServer() {
  const server = require('./server');
  return new Promise((resolve) => {
    server.listen(3721, '127.0.0.1', () => {
      console.log('Express server running on port 3721');
      resolve();
    });
  });
}

app.whenReady().then(async () => {
  await startServer();
  createWindow();
});

app.on('window-all-closed', () => {
  app.quit();
});

app.on('activate', () => {
  if (mainWindow === null) createWindow();
});
