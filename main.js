const { app, BrowserWindow } = require('electron');
const path = require('path');
const { spawn } = require('child_process');

let mainWindow;
let pyProcess = null;

const PY_PORT = 9112; // Choosing a specific port

function startPythonBackend() {
  const fs = require('fs');
  let pyPath;
  let args = [];
  let cwd;
  let shell = false;
  
  if (app.isPackaged) {
    const backendDir = path.join(process.resourcesPath, 'backend');
    const exePath = path.join(backendDir, 'app.exe');
    const appPyPath = path.join(backendDir, 'app.py');
    
    console.log(`[Backend] Starting Python backend from packaged app`);
    console.log(`[Backend] Backend directory: ${backendDir}`);
    console.log(`[Backend] app.exe exists: ${fs.existsSync(exePath)}`);
    console.log(`[Backend] app.py exists: ${fs.existsSync(appPyPath)}`);

    if (fs.existsSync(exePath)) {
      // If PyInstaller built executable exists, run it directly
      pyPath = exePath;
      args = [];
      cwd = backendDir;
      shell = false;
    } else {
      // Fallback to python.exe or system python if app.exe is not found
      const pythonExe = path.join(backendDir, 'python.exe');
      pyPath = 'cmd.exe';
      cwd = backendDir;
      shell = true;
      if (fs.existsSync(pythonExe)) {
        args = ['/c', `"${pythonExe}" app.py`];
      } else {
        console.log(`[Backend] No app.exe or python.exe in backend, using system Python`);
        args = ['/c', 'python app.py'];
      }
    }
  } else {
    // Development mode
    console.log(`[Backend] Starting Python backend in development mode`);
    pyPath = 'python';
    args = ['app.py'];
    cwd = process.cwd();
  }

  console.log(`[Backend] Spawning: ${pyPath}`);
  console.log(`[Backend] Args: ${JSON.stringify(args)}`);
  console.log(`[Backend] Working directory: ${cwd}`);

  pyProcess = spawn(pyPath, args, {
    cwd: cwd,
    env: { ...process.env, PORT: PY_PORT },
    shell: shell,
    detached: false
  });

  pyProcess.stdout.on('data', (data) => {
    console.log(`Backend: ${data}`);
  });

  pyProcess.stderr.on('data', (data) => {
    console.error(`Backend Error: ${data}`);
  });
}

const http = require('http');

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    title: "TCUA BOM Comparison",
    show: false, // Don't show until ready
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    },
    icon: path.join(__dirname, 'static', 'logo.png')
  });

  const url = `http://127.0.0.1:${PY_PORT}`;
  
  const checkServer = () => {
    http.get(url, (res) => {
      console.log('Server is ready, loading UI...');
      mainWindow.loadURL(url);
      mainWindow.once('ready-to-show', () => {
        mainWindow.show();
        // mainWindow.webContents.openDevTools(); // Useful for debugging if still white
      });
    }).on('error', () => {
      console.log('Waiting for server...');
      setTimeout(checkServer, 1000);
    });
  };

  checkServer();


  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.on('ready', () => {
  startPythonBackend();
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('quit', () => {
  if (pyProcess) {
    pyProcess.kill();
  }
});
