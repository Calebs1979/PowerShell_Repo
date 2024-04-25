# Start transcript
Start-Transcript -Path "C:\temp\winget_install_transcript.log" -Append

# Download and install winget
Invoke-WebRequest -Uri https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle -OutFile C:\temp\winget.msixbundle
Add-AppxPackage -Path C:\Temp\winget.msixbundle

# Sleep to allow download and install to complete
Start-Sleep -Seconds 60

# Set PSRepository PSGallery trusted
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

# Install winget module
Install-Module winget

# Import winget module
Import-Module winget

# Check for all updates
$updates = winget upgrade --all | Where-Object { $_.Title -notlike '*KB5034441*' }

# Filter out unknown third-party updates
$unknownUpdates = $updates | Where-Object { $_.Publisher -eq "Unknown" }

# Install updates silently without user action and accepting any prompts
foreach ($update in $unknownUpdates) {
    winget upgrade --id $update.Id --accept-source agpl --accept-package-agreements -y
}

# End transcript
Stop-Transcript

# Delete winget from c:\temp directory
Remove-Item "C:\temp\winget.msixbundle" -Force
