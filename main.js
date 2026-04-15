const { app, BrowserWindow } = require('electron');
const { execSync } = require('child_process');

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

function killPort3721() {
  try {
    const result = execSync('netstat -ano | findstr ":3721 "', { encoding: 'utf8' });
    const pids = new Set();
    for (const line of result.trim().split('\n')) {
      const parts = line.trim().split(/\s+/);
      const pid = parts[parts.length - 1];
      if (pid && pid !== '0') pids.add(pid);
    }
    for (const pid of pids) {
      try { execSync(`taskkill /PID ${pid} /F`); } catch (_) {}
    }
  } catch (_) {}
}

function startServer() {
  const server = require('./server');
  return new Promise((resolve, reject) => {
    const s = server.listen(3721, '127.0.0.1', () => {
      console.log('Express server running on port 3721');
      resolve();
    });
    s.on('error', (err) => {
      if (err.code === 'EADDRINUSE') {
        console.log('Port 3721 in use — killing existing process and retrying...');
        s.close();
        killPort3721();
        setTimeout(() => {
          const s2 = server.listen(3721, '127.0.0.1', () => {
            console.log('Express server running on port 3721 (retry)');
            resolve();
          });
          s2.on('error', reject);
        }, 600);
      } else {
        reject(err);
      }
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
