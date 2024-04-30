Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

#region Create Transcript Reports log folder if not exist
$folderPath = "C:\Transcripts\IntuneDeviceSync Logs\"
if (-not (Test-Path -Path $folderPath)) {      
    New-Item -Path $folderPath -ItemType Directory
}

# Start transcript log  file and write to path C:\ccadmin\
Start-Transcript -Path "C:\Transcripts\IntuneDeviceSync Logs\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Transcript Log.txt"

#endregion Create Reports log folder if not exist

#region Create Header Picture Box
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Location = New-Object System.Drawing.Point(10, 30)
    $imageURL = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Device%20Audit%20Assessment%20GUI/Security-assessment-image.png"
    $image = [System.Net.WebRequest]::Create($imageUrl).GetResponse().GetResponseStream()
    $pictureBox.Image = [System.Drawing.Image]::FromStream($image)
    $pictureBox.SizeMode = "Stretch"  # Fit the image within the PictureBox
    $pictureBox.Size = New-Object System.Drawing.Size(370, 70)
    $form.Controls.Add($pictureBox)
#endregion Create Picture Box

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Audit Device"
$form.Size = New-Object System.Drawing.Size(600,600)
$form.StartPosition = "CenterScreen"

# Create Label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(200,20)
$label.Text = "Audit Device"
$form.Controls.Add($label)

# Create RichTextBox
$richtextbox = New-Object System.Windows.Forms.RichTextBox
$richtextbox.Location = New-Object System.Drawing.Point(10,110)
$richtextbox.Size = New-Object System.Drawing.Size(560,400)
$form.Controls.Add($richtextbox)

# Create "Run Audit" Button
$buttonRunAudit = New-Object System.Windows.Forms.Button
$buttonRunAudit.Location = New-Object System.Drawing.Point(10,520)
$buttonRunAudit.Size = New-Object System.Drawing.Size(100,30)
$buttonRunAudit.Text = "Run Audit"
$buttonRunAudit.Add_Click({
    $richtextbox.Clear()

    # Get logged on user
    $loggedOnUser = whoami
    $richtextbox.AppendText("Logged On User: $loggedOnUser`n")

    # Get device name
    $deviceName = $env:COMPUTERNAME
    $richtextbox.AppendText("Device Name: $deviceName`n")

    # Get OS info
    $osInfo = Get-CimInstance Win32_OperatingSystem
    $richtextbox.AppendText("OS Info:`n")
    $richtextbox.AppendText("  Caption: $($osInfo.Caption)`n")
    $richtextbox.AppendText("  Version: $($osInfo.Version)`n")
    $richtextbox.AppendText("  BuildNumber: $($osInfo.BuildNumber)`n")

    # Get Disk info
    $diskInfo = Get-CimInstance Win32_LogicalDisk
    $richtextbox.AppendText("Disk Info:`n")
    foreach ($disk in $diskInfo) {
        $richtextbox.AppendText("  Drive: $($disk.DeviceID)`n")
        $richtextbox.AppendText("  Size: $($disk.Size/1GB) GB`n")
        $richtextbox.AppendText("  FreeSpace: $($disk.FreeSpace/1GB) GB`n")
    }

    # Get NIC info
    $nicInfo = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null }
    $richtextbox.AppendText("NIC Info:`n")
    foreach ($nic in $nicInfo) {
        $richtextbox.AppendText("  Description: $($nic.Description)`n")
        $richtextbox.AppendText("  IP Address: $($nic.IPAddress)`n")
    }

    # Get last reboot time
    $lastReboot = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty LastBootUpTime
    $richtextbox.AppendText("Last Reboot Time: $lastReboot`n")

    # Get device installed windows updates
    $installedUpdates = Get-WmiObject -Class Win32_QuickFixEngineering | Select-Object -Property Description,HotFixID,InstalledOn | Format-Table -AutoSize
    $richtextbox.AppendText("Installed Windows Updates:`n")
    $richtextbox.AppendText($installedUpdates)
    $richtextbox.AppendText("`n")

    # Get device missing updates
    $missingUpdates = Get-WindowsUpdate -NotCategory "Driver" -IsInstalled $false | Select-Object -Property Title,KB,Description,UpdateType | Format-Table -AutoSize
    $richtextbox.AppendText("Missing Windows Updates:`n")
    $richtextbox.AppendText($missingUpdates)
    $richtextbox.AppendText("`n")

    # Get device firewall status
    $firewallStatus = Get-NetFirewallProfile | Select-Object -Property Name,Enabled | Format-Table -AutoSize
    $richtextbox.AppendText("Firewall Status:`n")
    $richtextbox.AppendText($firewallStatus)
    $richtextbox.AppendText("`n")

    # Query the registry for installed antivirus software
    $antivirusInfo = @()

    # Check if Windows Defender is enabled
    $windowsDefender = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows Defender" -Name DisableAntiSpyware
if ($windowsDefender -eq $null -or $windowsDefender.DisableAntiSpyware -ne 1) {
    $antivirusInfo += [PSCustomObject]@{
        "Name" = "Windows Defender"
        "Status" = "Enabled"
    }
}

    # Query installed antivirus software from registry
    $antivirusRegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $antivirusKeys = Get-ChildItem -Path $antivirusRegistryPath | Where-Object { $_.GetValue("DisplayName") -match "antivirus|security" }

foreach ($key in $antivirusKeys) {
    $name = $key.GetValue("DisplayName")
    $status = "Installed"
    $antivirusInfo += [PSCustomObject]@{
        "Name" = $name
        "Status" = $status
    }
}

    # Output antivirus information
    $antivirusInfo | Format-Table -AutoSize
    $richtextbox.AppendText("Antivirus Information:`n")
    $richtextbox.AppendText($antivirusInfo)
    $richtextbox.AppendText("`n")

    # Check if SMBv1 is enabled
    $smbv1Enabled = (Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name "SMB1") -eq 1

    # Check if SMBv2 is enabled
    $smbv2Enabled = (Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name "SMB2") -eq 1

    # Check if SMBv3 is enabled
    $smbv3Enabled = (Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name "SMB3") -eq 1

    # Output SMB protocol status
    $richtextbox.AppendText("SMB Protocol Status:`n")
    $richtextbox.AppendText("SMBv1 Enabled: $smbv1Enabled`n")
    $richtextbox.AppendText("SMBv2 Enabled: $smbv2Enabled`n")
    $richtextbox.AppendText("SMBv3 Enabled: $smbv3Enabled`n")
    $richtextbox.AppendText("`n")

    # Define registry path for 32-bit and 64-bit applications
    $uninstallKeys32 = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $uninstallKeys64 = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

    # Get uninstall strings for 32-bit applications
    $uninstallStrings32 = Get-ItemProperty -Path $uninstallKeys32 | Where-Object {$_.DisplayName -ne $null} | Select-Object DisplayName, UninstallString

    # Get uninstall strings for 64-bit applications
    $uninstallStrings64 = Get-ItemProperty -Path $uninstallKeys64 | Where-Object {$_.DisplayName -ne $null} | Select-Object DisplayName, UninstallString

    # Combine 32-bit and 64-bit uninstall strings
    $uninstallStrings = $uninstallStrings32 + $uninstallStrings64

    # Output uninstall strings
    $richtextbox.AppendText("Uninstall Strings:`n")
    $uninstallStrings | Format-Table -AutoSize | Out-String | ForEach-Object { $richtextbox.AppendText($_) }
    $richtextbox.AppendText("`n")

    # Get BitLocker status
    $bitlockerStatus = Get-BitLockerVolume | Select-Object -Property VolumeStatus, MountPoint

    # Output BitLocker status
    $richtextbox.AppendText("BitLocker Status:`n")
    $bitlockerStatus | Format-Table -AutoSize | Out-String | ForEach-Object { $richtextbox.AppendText($_) }
    $richtextbox.AppendText("`n")

})
$form.Controls.Add($buttonRunAudit)

    # Create "Export to CSV" Button
    $buttonExportCSV = New-Object System.Windows.Forms.Button
    $buttonExportCSV.Location = New-Object System.Drawing.Point(120,520)
    $buttonExportCSV.Size = New-Object System.Drawing.Size(120,30)
    $buttonExportCSV.Text = "Export to CSV"
    $buttonExportCSV.Add_Click({
    $richtextbox.Text | Out-File -FilePath "AuditOutput.csv"
    Write-Host "Exported to CSV successfully."
})
    $form.Controls.Add($buttonExportCSV)

    # Create "Exit Form" Button
    $buttonExit = New-Object System.Windows.Forms.Button
    $buttonExit.Location = New-Object System.Drawing.Point(250,520)
    $buttonExit.Size = New-Object System.Drawing.Size(100,30)
    $buttonExit.Text = "Exit Form"
    $buttonExit.Add_Click({
    $form.Close()
})
    $form.Controls.Add($buttonExit)

    # Show the Form
    $form.ShowDialog() | Out-Null
