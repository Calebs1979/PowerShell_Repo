# Uninstall all versions of Java silently
# Get a list of installed Java versions
$javaVersions = Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" |
                Where-Object { $_.DisplayName -like 'Java *' }

# Uninstall each Java version silently
foreach ($javaVersion in $javaVersions) {
    $uninstallString = $javaVersion.UninstallString
    if ($uninstallString -ne $null) {
        # Extract the executable path from the uninstall string
        $executablePath = $uninstallString -replace '.*?([C-Z]:\\.*?\.exe).*', '$1'

        # Uninstall Java silently
        Start-Process -FilePath $executablePath -ArgumentList "/s", "/x" -Wait
    }
}
Write-Host "Java uninstallation completed."
