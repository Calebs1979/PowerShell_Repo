$registryPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
$registryName = "HideNewOutlookToggle"  # Replace with your desired registry entry name
$registryValue = 1

# Check if the registry path exists, and if not, create it
if (-not (Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force
}

# Create or update the DWORD registry entry
Set-ItemProperty -Path $registryPath -Name $registryName -Value $registryValue
