# bootstrap.ps1 — One-shot VM setup for DingTalk RPA
# Downloads and installs everything, deploys the project.
# Run as Administrator in PowerShell.

$ErrorActionPreference = "Stop"
$HostIP = "192.168.122.1"
$ProjectUrl = "http://${HostIP}:8080"
$ProjectDir = "C:\dd_group_collection"

Write-Host "=== Testing connectivity to host ===" -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri "http://${HostIP}:8080/" -UseBasicParsing -TimeoutSec 5 | Out-Null
    Write-Host "Host reachable."
} catch {
    Write-Host "ERROR: Cannot reach host at $HostIP:8080" -ForegroundColor Red
    Write-Host "Make sure the HTTP server is running on the host." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "=== Step 1: Install Python 3.11 ===" -ForegroundColor Cyan
$PythonInstaller = "$env:TEMP\python-3.11.9-amd64.exe"
$PythonUrl = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"

$pythonPath = Get-Command python.exe -ErrorAction SilentlyContinue | Where-Object { $_.Source -notlike "*WindowsApps*" }
if ($pythonPath) {
    $ver = & $pythonPath.Source --version 2>&1
    Write-Host "Python already installed: $ver"
} else {
    Write-Host "Downloading Python 3.11.9..."
    Invoke-WebRequest -Uri $PythonUrl -OutFile $PythonInstaller
    Write-Host "Installing Python (silent, this takes a minute)..."
    Start-Process -Wait -FilePath $PythonInstaller -ArgumentList `
        "/quiet", "InstallAllUsers=1", "PrependPath=1", "Include_pip=1"
    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    Write-Host "Python installed: $(python --version)"
}

Write-Host ""
Write-Host "=== Step 2: Deploy project files ===" -ForegroundColor Cyan
if (Test-Path $ProjectDir) {
    Remove-Item -Recurse -Force $ProjectDir
}
New-Item -ItemType Directory -Path $ProjectDir -Force | Out-Null

# Download each project file from host HTTP server
$files = @(
    "run.py",
    "start.bat",
    "config.yaml",
    "requirements.txt",
    "dd_collector/__init__.py",
    "dd_collector/config.py",
    "dd_collector/dedup.py",
    "dd_collector/dingtalk_ui.py",
    "dd_collector/file_mover.py",
    "dd_collector/logger.py",
    "dd_collector/main.py",
    "dd_collector/ui_helpers.py",
    "tools/inspect_dingtalk.py"
)

foreach ($file in $files) {
    $dir = Split-Path -Parent "$ProjectDir\$file"
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $url = "$ProjectUrl/$file"
    $dest = "$ProjectDir\$file"
    Write-Host "  Downloading $file..."
    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
    } catch {
        Write-Host "  WARNING: Failed to download $file - $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Create data directory
New-Item -ItemType Directory -Path "$ProjectDir\data" -Force | Out-Null
New-Item -ItemType Directory -Path "$ProjectDir\logs" -Force | Out-Null

Write-Host "Project deployed to $ProjectDir"

Write-Host ""
Write-Host "=== Step 3: Install pip dependencies ===" -ForegroundColor Cyan
python -m pip install --upgrade pip 2>&1 | Select-Object -Last 1
python -m pip install -r "$ProjectDir\requirements.txt"

Write-Host ""
Write-Host "=== Step 4: Update config.yaml for this VM ===" -ForegroundColor Cyan
$configPath = "$ProjectDir\config.yaml"
$config = Get-Content $configPath -Raw
$config = $config -replace 'C:\\Users\\rick\\', "C:\Users\$env:USERNAME\"
Set-Content -Path $configPath -Value $config
Write-Host "Updated paths for user: $env:USERNAME"

Write-Host ""
Write-Host "=== Step 5: Disable screen lock and standby ===" -ForegroundColor Cyan
powercfg -change -monitor-timeout-ac 0
powercfg -change -standby-timeout-ac 0
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization" /v NoLockScreen /t REG_DWORD /d 1 /f | Out-Null
Write-Host "Screen lock and standby disabled."

Write-Host ""
Write-Host "=== Step 6: Create startup shortcut ===" -ForegroundColor Cyan
$startupDir = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
$shortcutPath = "$startupDir\dd_group_collection.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "$ProjectDir\start.bat"
$shortcut.WorkingDirectory = $ProjectDir
$shortcut.Save()
Write-Host "Startup shortcut created."

Write-Host ""
Write-Host "=== Step 7: Configure auto-login ===" -ForegroundColor Cyan
$username = $env:USERNAME
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoAdminLogon /t REG_SZ /d 1 /f | Out-Null
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultUserName /t REG_SZ /d $username /f | Out-Null
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultPassword /t REG_SZ /d "" /f | Out-Null
Write-Host "Auto-login configured for user: $username (no password)"

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "  Setup complete!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Remaining manual steps:"
Write-Host "  1. Install DingTalk: https://www.dingtalk.com/download"
Write-Host "     -> Log in, enable auto-start in Settings > General"
Write-Host "  2. Install Google Drive for Desktop"
Write-Host "     -> Log in, verify G:\ drive appears"
Write-Host "  3. Test: cd $ProjectDir && python tools\inspect_dingtalk.py"
Write-Host "  4. Run:  cd $ProjectDir && python run.py"
Write-Host ""
