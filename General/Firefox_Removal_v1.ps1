# Uninstall Firefox
$firefoxUninstallString = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Mozilla Firefox" } | Select-Object -ExpandProperty UninstallString
if ($firefoxUninstallString) {
    Write-Host "Uninstalling Mozilla Firefox..."
    Start-Process "$firefoxUninstallString" -Wait
    Write-Host "Mozilla Firefox has been uninstalled."
} else {
    Write-Host "Mozilla Firefox not found. No action taken."
}

# Remove Firefox shortcuts
$firefoxShortcuts = @(
    [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Firefox.lnk'),
    [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\Firefox.lnk'),
    [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\Firefox Private Browsing.lnk'),
    [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Windows\Start Menu\Programs\Firefox.lnk'),
    [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Windows\Start Menu\Programs\Firefox Private Browsing.lnk'),
    [System.IO.Path]::Combine($env:PROGRAMDATA, 'Microsoft\Windows\Start Menu\Programs\Firefox.lnk'),
    [System.IO.Path]::Combine($env:PROGRAMDATA, 'Microsoft\Windows\Start Menu\Programs\Firefox Private Browsing.lnk')
)

foreach ($shortcut in $firefoxShortcuts) {
    if (Test-Path $shortcut) {
        Write-Host "Removing shortcut: $shortcut"
        Remove-Item $shortcut -Force
    }
}

# Remove Firefox from recently added applications
$recentlyAddedApps = Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs
foreach ($key in $recentlyAddedApps.PSObject.Properties) {
    if ($key.Name -match 'Firefox') {
        Write-Host "Removing recently added application: $key"
        Remove-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs -Name $key.Name
    }
}

Write-Host "Cleanup completed."
