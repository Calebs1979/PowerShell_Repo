Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Enable Visual Styles
[Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::EnableVisualStyles()

# Get the current logged-on user
$currentUserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name 
#
# Extract only the username from the full user name
$currentUserName = $currentUserName -replace ".*\\"  
#
# Capitalize the first letter of the username
$capitalizedUserName = $currentUserName.Substring(0,1).ToUpper() + $currentUserName.Substring(1)
#
# Create the message
$message = "Hello, $capitalizedUserName!`n`nPlease ensure that you close the Tool when you have completed.`n`nThank you $capitalizedUserName" 
#                                                                                                                                                               
# Display the popup window                                                                                                                                       
Add-Type -AssemblyName PresentationFramework                                                                                                                     
[System.Windows.MessageBox]::Show($message, "Informational Message", "OK", "Information")
#
# Create transcript log folder if not exist
$folderPath = "C:\Temp\Users Logon History Reports\Transcript logs"
if (-not (Test-Path -Path $folderPath)) {  
    New-Item -Path $folderPath -ItemType Directory
}
#
# Create Reports log folder if not exist
$folderPath = "C:\Temp\Users Logon History Reports"
if (-not (Test-Path -Path $folderPath)) {
    New-Item -Path $folderPath -ItemType Directory
}
#
# Start transcript log  file and write to path C:\RTAdmin\
Start-Transcript -Path "C:\Temp\Users Logon History Reports\Transcript logs\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Transcript Log.txt"

# Function to search user logon history across all domain controllers or a specific one
function Search-UserLogonHistory {
    param (
        [string]$username,
        [int]$days,
        [string]$domainController
    )

    # Get current date
    $endDate = Get-Date

    # Calculate start date based on selected days
    $startDate = $endDate.AddDays(-$days)

    # Initialize array to store logon events
    $logonEvents = @()

    # Query a specific domain controller or all domain controllers in the domain
    if ($domainController) {
        $dcList = @($domainController)
    } else {
        $dcList = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName
    }

    # Query each domain controller for logon events
    foreach ($dc in $dcList) {
        $events = Get-WinEvent -ComputerName $dc -FilterHashtable @{
            LogName = 'Security'
            ID = 4624
            StartTime = $startDate
            EndTime = $endDate
        } | Where-Object { $_.Properties[5].Value -eq $username } | ForEach-Object {
            [PSCustomObject]@{
                'User' = $_.Properties[5].Value
                'LogonDate' = $_.TimeCreated
                'DomainController' = $dc
            }
        }

        # Add events from current domain controller to the logon events array
        $logonEvents += $events
    }

    # Output logon events
    return $logonEvents
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "User Logon History"
$form.Size = New-Object System.Drawing.Size(600,600)
$form.BackColor = [System.Drawing.Color]::AliceBlue
$form.Forecolor="#000000"
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

# Create PictureBox
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.ImageLocation = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Active%20Directory%20Users%20Logon%20History%20Report%20GU/Header.BlueText.png"  # Replace with the path to your image
$pictureBox.Size = New-Object System.Drawing.Size(430, 165)
$pictureBox.Location = New-Object System.Drawing.Point(70,10)
$pictureBox.SizeMode = "Zoom"
$form.Controls.Add($pictureBox)

# Add controls to the form
$labelUsername = New-Object System.Windows.Forms.Label
$labelUsername.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
$labelUsername.Text = "Username:"
$labelUsername.Location = New-Object System.Drawing.Point(10,180)
$labelUsername.AutoSize = $true
$form.Controls.Add($labelUsername)

$textboxUsername = New-Object System.Windows.Forms.TextBox
$textboxUsername.Location = New-Object System.Drawing.Point(90,179)
$textboxUsername.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textboxUsername)

$labelDays = New-Object System.Windows.Forms.Label
$labelDays.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
$labelDays.Text = "Select days:"
$labelDays.Location = New-Object System.Drawing.Point(300,180)
$labelDays.AutoSize = $true
$form.Controls.Add($labelDays)

$dropdownDays = New-Object System.Windows.Forms.ComboBox
$dropdownDays.Location = New-Object System.Drawing.Point(385,179)
$dropdownDays.Size = New-Object System.Drawing.Size(150,20)
$dropdownDays.Items.AddRange(@("Past 30 days", "Past 60 days", "Past 90 days"))
$form.Controls.Add($dropdownDays)

$labelDomainController = New-Object System.Windows.Forms.Label
$labelDomainController.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
$labelDomainController.Text = "Select Domain Controller (Optional):"
$labelDomainController.Location = New-Object System.Drawing.Point(10,220)
$labelDomainController.AutoSize = $true
$form.Controls.Add($labelDomainController)

$dropdownDomainController = New-Object System.Windows.Forms.ComboBox
$dropdownDomainController.Location = New-Object System.Drawing.Point(253,219)
$dropdownDomainController.Size = New-Object System.Drawing.Size(150,20)
# Populate domain controller dropdown with domain controllers
$dropdownDomainController.Items.AddRange((Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName))
$form.Controls.Add($dropdownDomainController)

$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.BackColor = "#00b4f2"
$buttonRun.ForeColor = "#ffffff"
$buttonRun.Text = "Run"
$buttonRun.Location = New-Object System.Drawing.Point(10,247)
$buttonRun.Size = New-Object System.Drawing.Size(80, 30)
$buttonRun.Add_Click({
    $username = $textboxUsername.Text
    $selectedDays = $dropdownDays.SelectedItem
    $selectedDomainController = $dropdownDomainController.SelectedItem

    if (-not [string]::IsNullOrEmpty($username) -and $selectedDays) {
        switch ($selectedDays) {
            "Past 30 days" {
                $logonEvents = Search-UserLogonHistory -username $username -days 30 -domainController $selectedDomainController
            }
            "Past 60 days" {
                $logonEvents = Search-UserLogonHistory -username $username -days 60 -domainController $selectedDomainController
            }
            "Past 90 days" {
                $logonEvents = Search-UserLogonHistory -username $username -days 90 -domainController $selectedDomainController
            }
        }

        # Display logon events in the output textbox
        if ($logonEvents) {
            $textboxOutput.Text = $logonEvents | Format-Table -AutoSize | Out-String
        } else {
            $textboxOutput.Text = "No logon events found for user '$username' in the past $selectedDays."
        }
    } else {
        $textboxOutput.Text = "Please provide a username and select days."
    }
})
$form.Controls.Add($buttonRun)

$buttonalluserslastlogondate = New-Object System.Windows.Forms.Button
$buttonalluserslastlogondate.BackColor = "#91c300"
$buttonalluserslastlogondate.ForeColor = "#ffffff"
$buttonalluserslastlogondate.Text = "List All users Last Logon Date"
$buttonalluserslastlogondate.Location = New-Object System.Drawing.Point(100,247)
$buttonalluserslastlogondate.Size = New-Object System.Drawing.Size(200, 30)
$buttonalluserslogonhistory.Add_Click({
$Output = Get-ADUser -Filter * -Properties LastLogon | Select-Object Name, DistinguishedName, SID, @{Name="LastLogon"; Expression={[DateTime]::FromFileTime($_.LastLogon)}} | Out-GridView -Title "Search results for all users last logon date." -PassThru | Export-Csv -Path "C:\Temp\Users Logon History Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All users Last Logon Date Report.csv" -NoTypeInformation -Delimiter ";"})
Write-host "List All users Last Logon Date Report, has been exported to 'C:\Temp\Users Logon History Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All users Last Logon Date Report.csv'"
$form.Controls.Add($buttonalluserslastlogondate)

$textboxOutput = New-Object System.Windows.Forms.TextBox
$textboxOutput.Multiline = $true
$textboxOutput.Location = New-Object System.Drawing.Point(10,280)
$textboxOutput.Size = New-Object System.Drawing.Size(570,270)
$textboxOutput.ScrollBars = "Vertical"
$textboxOutput.Text = ""
$textboxOutput.Font = New-Object System.Drawing.Font("Consolas", 10)
function Add-textboxOutput{
		[CmdletBinding()]
		param ($text)
		$textboxOutput.Text += "$text;"
		$textboxOutput.Text += "`n# # # # # # # # # #`n"
	}
$form.Controls.Add($textboxOutput)

# Create Export to CSV Button
$exportToCsvButton = New-Object System.Windows.Forms.Button
$exportToCsvButton.BackColor = "#00b4f2"
$exportToCsvButton.ForeColor = "#ffffff"
$exportToCsvButton.Location = New-Object System.Drawing.Point(10, 560)
$exportToCsvButton.Size = New-Object System.Drawing.Size(80, 30)
$exportToCsvButton.Text = "Export to CSV"
$exportToCsvButton.AutoSize = $true
$exportToCsvButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.Title = "Export to CSV"
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $outputText = $textboxOutput.Text
        $outputText | Out-File -FilePath $saveFileDialog.FileName -Encoding utf8
        [System.Windows.Forms.MessageBox]::Show("File exported successfully!", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK)
    }
})
$form.Controls.Add($exportToCsvButton)

# Create Copy to Clipboard Button
$copyToClipboardButton = New-Object System.Windows.Forms.Button
$copyToClipboardButton.BackColor = "#00b4f2"
$copyToClipboardButton.ForeColor = "#ffffff"
$copyToClipboardButton.Location = New-Object System.Drawing.Point(205, 560)
$copyToClipboardButton.Size = New-Object System.Drawing.Size(100, 30)
$copyToClipboardButton.Text = "Copy to Clipboard"
$copyToClipboardButton.AutoSize = $true
$copyToClipboardButton.Add_Click({
    $textboxOutput.SelectAll()
    $textboxOutput.Copy()
})
$form.Controls.Add($copyToClipboardButton)

# Create Clear Output Button
$clearOutputButton = New-Object System.Windows.Forms.Button
$clearOutputButton.BackColor = "#00b4f2"
$clearOutputButton.ForeColor = "#ffffff"
$clearOutputButton.Location = New-Object System.Drawing.Point(110, 560)
$clearOutputButton.Size = New-Object System.Drawing.Size(80, 30)
$clearOutputButton.Text = "Clear Output"
#$clearOutputButton.AutoSize = $true
$clearOutputButton.Add_Click({
    $textboxOutput.Clear()
})
$form.Controls.Add($clearOutputButton)

# Create Exit Button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.BackColor = "#00b4f2"
$exitButton.ForeColor = "#ffffff"
$exitButton.Location = New-Object System.Drawing.Point(500, 560)
$exitButton.Size = New-Object System.Drawing.Size(80, 30)
$exitButton.Text = "Exit"
#$exitButton.AutoSize = $true
$exitButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($exitButton)

# Display the form
$form.ShowDialog() | Out-Null
