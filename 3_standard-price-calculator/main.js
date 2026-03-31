const { app, BrowserWindow, nativeTheme } = require('electron');
const path = require('path');

nativeTheme.themeSource = 'dark';

// 서버 내장 - 별도 프로세스 불필요
require('./server.js');

let mainWindow = null;

app.whenReady().then(() => {
  mainWindow = new BrowserWindow({
    width: 1129,
    height: 750,
    x: 150,
    y: 30,
    backgroundColor: '#191919',
    title: 'TAX AI',
    icon: path.join(__dirname, 'taxai.ico'),
    autoHideMenuBar: true,
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  mainWindow.loadURL('http://localhost:3100');
  mainWindow.setTitle('TAX AI');

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
});

app.on('window-all-closed', () => {
  app.quit();
});
