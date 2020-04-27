/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* from the electron quick-start */
let electron = require('electron');
let XLSX = require('xlsx');
let app = electron.app;

let win = null;

function createWindow() {
	if(win) return;
	win = new electron.BrowserWindow({width:1200, height:700});
	win.loadURL("file://" + __dirname + "/index.html");
	win.webContents.openDevTools();
	win.on('closed', function() { win = null; });
}
if(app.setAboutPanelOptions) app.setAboutPanelOptions({ applicationName: 'sheetjs-electron', applicationVersion: "XLSX " + XLSX.version, copyright: "(C) 2017-present SheetJS LLC" });
app.on('open-file', function() { console.log(arguments); });
app.on('ready', createWindow);
app.on('activate', createWindow);
app.on('window-all-closed', function() { if(process.platform !== 'darwin') app.quit(); });
