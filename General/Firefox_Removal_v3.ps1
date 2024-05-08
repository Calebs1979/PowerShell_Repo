# Uninstall Mozilla Firefox
$firefoxVersions = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE 'Firefox%'" | Select-Object Name, Version

foreach ($version in $firefoxVersions) {
    Write-Host "Uninstalling $($version.Name) $($version.Version)"
    Start-Process "msiexec.exe" -ArgumentList "/x $($version.Name)" -Wait
}

# Remove registry entries
$registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$firefoxKeys = Get-Item -LiteralPath $registryPath | Get-ItemProperty | Where-Object { $_.DisplayName -like 'Firefox*' }

foreach ($key in $firefoxKeys) {
    Write-Host "Removing registry entry for $($key.DisplayName)"
    Remove-Item -LiteralPath "$registryPath\$($key.PSChildName)" -Force
}

# Remove user profiles
$profilesPath = "C:\Users"
$firefoxProfiles = Get-ChildItem -Path $profilesPath -Filter "Firefox*" -Recurse -Directory

foreach ($profile in $firefoxProfiles) {
    Write-Host "Removing Firefox files from $($profile.FullName)"
    Remove-Item -LiteralPath $profile.FullName -Recurse -Force
}

# Remove AppData directories
$appDataPath = "C:\Users\*\AppData\"
$firefoxAppData = Get-ChildItem -Path $appDataPath -Filter "Mozilla Firefox" -Recurse -Directory

foreach ($appDataDir in $firefoxAppData) {
    Write-Host "Removing Firefox AppData from $($appDataDir.FullName)"
    Remove-Item -LiteralPath $appDataDir.FullName -Recurse -Force
}

Write-Host "Mozilla Firefox uninstallation and cleanup completed."
