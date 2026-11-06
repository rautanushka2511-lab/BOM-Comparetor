const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

const backendDir = 'D:\\akshar\\apps\\TCUA Perfect\\Deployment\\win-unpacked\\resources\\backend';
const appBatPath = path.join(backendDir, 'app.exe.bat');

console.log(`Testing backend spawn...`);
console.log(`Backend dir: ${backendDir}`);
console.log(`app.exe.bat exists: ${fs.existsSync(appBatPath)}`);

if (fs.existsSync(appBatPath)) {
  console.log(`\nAttempting to spawn cmd.exe with app.exe.bat...`);

  const pyProcess = spawn('cmd.exe', ['/c', appBatPath], {
    cwd: backendDir,
    env: { ...process.env, PORT: 9112 },
    shell: true
  });

  pyProcess.stdout.on('data', (data) => {
    console.log(`stdout: ${data}`);
  });

  pyProcess.stderr.on('data', (data) => {
    console.log(`stderr: ${data}`);
  });

  pyProcess.on('error', (err) => {
    console.error(`Error spawning process: ${err.message}`);
  });

  pyProcess.on('exit', (code) => {
    console.log(`Process exited with code ${code}`);
  });

  // Keep the script running for 10 seconds
  setTimeout(() => {
    console.log('\nTest complete');
    process.exit(0);
  }, 10000);
} else {
  console.error(`app.exe.bat not found!`);
}

