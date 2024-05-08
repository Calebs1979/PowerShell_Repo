# Uninstall Firefox
$firefoxPath = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -eq "Mozilla Firefox" } | Select-Object -ExpandProperty UninstallString

if ($firefoxPath) {
    Write-Host "Uninstalling Firefox..."
    Start-Process -FilePath "$firefoxPath" -ArgumentList "/S" -Wait
    Write-Host "Firefox has been uninstalled."
} else {
    Write-Host "Firefox is not installed on this system."
}

# Remove leftover files and registry entries
$firefoxInstallDir = "C:\Program Files\Mozilla Firefox"

if (Test-Path $firefoxInstallDir) {
    Write-Host "Removing leftover files and registry entries..."
    
    # Stop Firefox processes
    Stop-Process -Name "firefox" -Force -ErrorAction SilentlyContinue
    
    # Remove installation directory
    Remove-Item -Path $firefoxInstallDir -Recurse -Force -ErrorAction SilentlyContinue
     
    # Remove Firefox registry entries
    Remove-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Mozilla\Mozilla Firefox" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "HKCU:\Software\Mozilla\Mozilla Firefox" -Recurse -Force -ErrorAction SilentlyContinue
    
    Write-Host "Leftover files and registry entries have been removed."
} else {
    Write-Host "No leftover files or registry entries found."
}

# Remove leftover files and registry entries from %APPDATA%
$firefoxAppDataDir = "$env:APPDATA\Mozilla\Firefox"

if (Test-Path $firefoxAppDataDir) {
    Write-Host "Removing leftover files and registry entries..."
    
    # Stop Firefox processes
    Stop-Process -Name "firefox" -Force -ErrorAction SilentlyContinue
      
    # Remove user data directory
    Remove-Item -Path $firefoxAppDataDir -Recurse -Force -ErrorAction SilentlyContinue
    
    # Remove Firefox registry entries
    Remove-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Mozilla\Mozilla Firefox" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "HKCU:\Software\Mozilla\Mozilla Firefox" -Recurse -Force -ErrorAction SilentlyContinue
    
    Write-Host "Leftover files and registry entries have been removed."
} else {
    Write-Host "No leftover files or registry entries found."
}

# Remove Firefox shortcuts
$firefoxShortcuts = Get-ChildItem -Path "$env:USERPROFILE\OneDrive\Desktop" -Filter "Mozilla Firefox*.lnk" -File

foreach ($shortcut in $firefoxShortcuts) {
    Remove-Item -Path $shortcut.FullName -Force
    Write-Host "Removed shortcut: $($shortcut.FullName)"
}

foreach ($shortcut in $firefoxShortcuts) {
    Remove-Item -Path $shortcut.FullName -Force
    Write-Host "Removed shortcut: $($shortcut.FullName)"
}