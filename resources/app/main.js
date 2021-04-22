const { app, BrowserWindow } = require('electron')

function createWindow () {
  const win = new BrowserWindow({
    width: 980,
    height: 900,
    webPreferences: {
      //contextIsolation: true // prevents access to version numbers of chrome, node, and electron, but is recommended for all applications--> keep an eye on
      nodeIntegration: true,
      enableRemoteModule: true
    
    }
  })

  win.loadFile('index.html')
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})