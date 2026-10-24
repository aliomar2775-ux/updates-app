'use strict';
const { app, BrowserWindow } = require('electron');
const path = require('path');
const http = require('http');

let mainWindow = null;
let serverProcess = null;

const PORT = process.env.PORT || 3000;
const AUTH_USER = process.env.APP_USER || 'admin@12';
const AUTH_PASS = process.env.APP_PASS || 'admin123';

function startServer() {
  try {
    serverProcess = require('./server.js');
  } catch (err) {
    console.error('Failed to start server:', err);
  }
}

function waitForServer(callback, maxAttempts = 30) {
  let attempts = 0;
  const interval = setInterval(() => {
    attempts++;
    const req = http.get(`http://127.0.0.1:${PORT}/healthz`, res => {
      if (res.statusCode === 200) {
        clearInterval(interval);
        callback();
      }
    });
    req.on('error', () => {
      if (attempts >= maxAttempts) {
        clearInterval(interval);
        console.error('Server did not respond in time, loading anyway...');
        callback();
      }
    });
    req.end();
  }, 250);
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 850,
    minWidth: 960,
    minHeight: 650,
    title: 'كتيبة الاتصالات — سجل تحديث الأجهزة',
    icon: path.join(__dirname, 'public', 'icon.png'),
    autoHideMenuBar: true,
    show: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    mainWindow.focus();
  });

  mainWindow.loadURL(`http://localhost:${PORT}`);

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// المصادقة التلقائية لنافذة التطبيق المكتبي حتى لا يطلب كلمة مرور محلياً
app.on('login', (event, webContents, request, authInfo, callback) => {
  event.preventDefault();
  callback(AUTH_USER, AUTH_PASS);
});

// تشغيل نسخة واحدة فقط من التطبيق
const gotSingleInstanceLock = app.requestSingleInstanceLock();
if (!gotSingleInstanceLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.show();
      mainWindow.focus();
    }
  });

  app.whenReady().then(() => {
    startServer();
    waitForServer(() => {
      createWindow();
    });

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
      app.quit();
    }
  });

  app.on('will-quit', () => {
    if (serverProcess && typeof serverProcess.close === 'function') {
      serverProcess.close();
    }
  });
}
