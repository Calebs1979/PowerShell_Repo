$autoPlaySetting = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" -Name "DisableAutoplay"

if ($autoPlaySetting.DisableAutoplay -eq 0) {
    Write-Host "Autoplay is enabled."
} else {
    Write-Host "Autoplay is disabled."
}
