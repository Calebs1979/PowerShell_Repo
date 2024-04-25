# Check if directory c:\temp\dotnet folder exists and create it if not
$dotnetFolderPath = "C:\temp\dotnet-core-uninstall-tool"
if (-not (Test-Path $dotnetFolderPath)) {
    New-Item -Path $dotnetFolderPath -ItemType Directory | Out-Null
    Write-Host "Created $dotnetFolderPath"
}

# Download dotnet-core-uninstall-1.6.0.msi from the URL and save it to c:\temp\dotnet folder
$downloadUrl = "https://github.com/dotnet/cli-lab/releases/download/1.6.0/dotnet-core-uninstall-1.6.0.msi"
$downloadPath = Join-Path -Path $dotnetFolderPath -ChildPath "dotnet-core-uninstall-1.6.0.msi"
Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath
Write-Host "Downloaded dotnet-core-uninstall-1.6.0.msi to $downloadPath"

# Silently install dotnet-core-uninstall-1.6.0.msi
Write-Host "Installing dotnet-core-uninstall-1.6.0.msi..."
Start-Process msiexec -ArgumentList "/i `"$downloadPath`" /quiet /norestart" -Wait

# Wait for 30 seconds
Start-Sleep -Seconds 30

# Run dotnet-core-uninstall remove 3.1.32 --aspnet-runtime --yes --verbosity q
dotnet-core-uninstall.exe remove 3.1.32 --runtime --yes --verbosity q
dotnet-core-uninstall.exe remove 3.1.32 --aspnet-runtime --yes --verbosity q