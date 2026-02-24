# setup-vm-guest.ps1 — Set up the Windows VM guest for DingTalk RPA
# Run inside the Windows VM as Administrator.
# Requires internet access.

$ErrorActionPreference = "Stop"

$ProjectDir = "C:\dd_group_collection"
$PythonVersion = "3.11.9"
$PythonInstaller = "$env:TEMP\python-$PythonVersion-amd64.exe"
$PythonUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
$DingTalkUrl = "https://g.alicdn.com/dingding/win-download/6.5.50/DingTalk-win-x64.exe"
$DingTalkInstaller = "$env:TEMP\DingTalk-installer.exe"

Write-Host "=== Step 1: Install Python $PythonVersion ===" -ForegroundColor Cyan
if (Get-Command python -ErrorAction SilentlyContinue) {
    Write-Host "Python already installed: $(python --version)"
} else {
    Write-Host "Downloading Python..."
    Invoke-WebRequest -Uri $PythonUrl -OutFile $PythonInstaller
    Write-Host "Installing Python (silent)..."
    Start-Process -Wait -FilePath $PythonInstaller -ArgumentList `
        "/quiet", "InstallAllUsers=1", "PrependPath=1", "Include_pip=1"
    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    Write-Host "Python installed: $(python --version)"
}

Write-Host ""
Write-Host "=== Step 2: Install pip packages ===" -ForegroundColor Cyan
python -m pip install --upgrade pip
python -m pip install uiautomation PyYAML

Write-Host ""
Write-Host "=== Step 3: Download DingTalk installer ===" -ForegroundColor Cyan
if (Test-Path $DingTalkInstaller) {
    Write-Host "DingTalk installer already downloaded."
} else {
    Write-Host "Downloading DingTalk..."
    Invoke-WebRequest -Uri $DingTalkUrl -OutFile $DingTalkInstaller
}
Write-Host "DingTalk installer saved to: $DingTalkInstaller"
Write-Host "  -> Run it manually and log in to your account."

Write-Host ""
Write-Host "=== Step 4: Configure Windows for unattended operation ===" -ForegroundColor Cyan

# Disable screen lock / monitor timeout
powercfg -change -monitor-timeout-ac 0
powercfg -change -standby-timeout-ac 0
Write-Host "Disabled monitor timeout and standby."

# Disable lock screen
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization" /v NoLockScreen /t REG_DWORD /d 1 /f | Out-Null
Write-Host "Disabled lock screen."

Write-Host ""
Write-Host "=== Step 5: Configure auto-login ===" -ForegroundColor Cyan
$username = $env:USERNAME
$password = Read-Host "Enter password for auto-login (or press Enter to skip)"
if ($password) {
    reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoAdminLogon /t REG_SZ /d 1 /f | Out-Null
    reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultUserName /t REG_SZ /d $username /f | Out-Null
    reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultPassword /t REG_SZ /d $password /f | Out-Null
    Write-Host "Auto-login configured for user: $username"
} else {
    Write-Host "Skipped auto-login setup."
}

Write-Host ""
Write-Host "=== Step 6: Create startup shortcut for RPA ===" -ForegroundColor Cyan
$startupDir = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
$shortcutPath = "$startupDir\dd_group_collection.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "$ProjectDir\start.bat"
$shortcut.WorkingDirectory = $ProjectDir
$shortcut.Save()
Write-Host "Startup shortcut created: $shortcutPath"

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "  Guest setup complete!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Manual steps remaining:"
Write-Host "  1. Run DingTalk installer: $DingTalkInstaller"
Write-Host "  2. Log in to DingTalk and enable auto-start in Settings > General"
Write-Host "  3. Install Google Drive for Desktop and log in"
Write-Host "  4. Copy project files to $ProjectDir"
Write-Host "  5. Edit $ProjectDir\config.yaml with your group names"
Write-Host "  6. Test: python $ProjectDir\tools\inspect_dingtalk.py"
Write-Host ""
