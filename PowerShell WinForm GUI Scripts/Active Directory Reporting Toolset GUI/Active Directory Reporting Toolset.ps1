Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check if Active Directory module is installed
            $adModule = Get-Module -Name ActiveDirectory -ListAvailable
# If not installed, install Active Directory module
        if (-not $adModule) {
            Write-Host "Installing Active Directory module..."
            Install-WindowsFeature RSAT-AD-PowerShell
# Import Active Directory module
            Import-Module ActiveDirectory
}
# Create transcript log folder if not exist
            $folderPath = "C:\RTAdmin\Reports Toolset Files\Transcript logs"
            if (-not (Test-Path -Path $folderPath)) {  
            New-Item -Path $folderPath -ItemType Directory
}
# Create Reports log folder if not exist
            $folderPath = "C:\RTAdmin\Reports Toolset Files\Reports"
            if (-not (Test-Path -Path $folderPath)) {
            New-Item -Path $folderPath -ItemType Directory
}
# Start transcript log  file and write to path C:\RTAdmin\
            Start-Transcript -Path "C:\RTAdmin\Reports Toolset Files\Transcript logs\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Log1.txt"
# Create the form
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Active Directory Report Toolset"
            $form.Size = New-Object System.Drawing.Size(800, 800)
            $Form.StartPosition = "CenterScreen"
            $form.BackColor = [System.Drawing.Color]::AliceBlue
            $form.Forecolor="#000000"

# monthCalendar1
#
            $monthCalendar1 = New-Object System.Windows.Forms.MonthCalendar
            $monthCalendar1.BackColor = [System.Drawing.Color]::AliceBlue
            $monthCalendar1.FirstDayOfWeek = [System.Windows.Forms.Day]::Monday
            $monthCalendar1.Location = New-Object System.Drawing.Point(520, 50)
            $monthCalendar1.MaximumSize = New-Object System.Drawing.Size(230, 200)
            $monthCalendar1.Name = "monthCalendar1"
            $monthCalendar1.TabIndex = 1
            $form.Controls.Add($monthCalendar1)

# Create a picture box
            $pictureBox = New-Object System.Windows.Forms.PictureBox
            $pictureBox.Size = New-Object System.Drawing.Size(450, 200)
            $pictureBox.Location = New-Object System.Drawing.Point(20, 20)
            $pictureBox.ImageLocation = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Active%20Directory%20Reporting%20Toolset%20GUI/AD_ReportingToolset_Header.png"
            $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
            $form.Controls.Add($pictureBox)
# Create a label
            $labelTitle = New-Object System.Windows.Forms.Label
            $labelTitle.Text = "Select Function"
            $labelTitle.AutoSize = $true
            $labelTitle.Forecolor = [System.Drawing.Color]::LightSteelBlue
            $labelTitle.Location = New-Object System.Drawing.Point(330, 257)
            $labelTitle.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
            $form.Controls.Add($labelTitle)
# Create a dropdown list for functions
            $comboBoxFunctions = New-Object System.Windows.Forms.ComboBox
            $comboBoxFunctions.Size = New-Object System.Drawing.Size(300, 30)
            $comboBoxFunctions.Location = New-Object System.Drawing.Point(10, 260)
            $comboBoxFunctions.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
            $form.Controls.Add($comboBoxFunctions)
# Add functions to the dropdown list
            $functions = @(
            "Backup All Group Policies",
            "Check AD Replication",
            "Default Domain Password Policy",
            "Export All DNS Zone Records",
            "Generate Passwords",
            "List All AD Users",
            "List All Domain Controllers",
            "List All Domain SSL Certificates",
            "List All Installed Hotfixes",
            "List All Objects Deleted in past month"
            "List All Security Groups",
            "List All Users Password Last Reset",
            "List All Users Logon History",
            "List User Password Last Reset",
            "List User Logon History",
            "List Disabled User Accounts",
            "List Inactive Users (30 Days)",
            "List Inactive Users (60 Days)",
            "List Inactive Users (90 Days)",
            "List Locked User Accounts",
            "List Users Created in the Past Month",
            "Search Group Policies",
            "Unlock AD User Accounts"
            
)
            $functions | ForEach-Object { $comboBoxFunctions.Items.Add($_) }
# Create a text box for status output
            $textBoxOutput = New-Object System.Windows.Forms.TextBox
            $textBoxOutput.Multiline = $true
            $textBoxOutput.ScrollBars = "Vertical"
            $textBoxOutput.Size = New-Object System.Drawing.Size(760, 400)
            $textBoxOutput.Location = New-Object System.Drawing.Point(10, 300)
            $textBoxOutput.Text = ""
            $textBoxOutput.Font = New-Object System.Drawing.Font("Consolas", 10)
            function Add-textBoxOutput{
		    [CmdletBinding()]
		    param ($text)
		    $textBoxOutput.Text += "$text;"
		    $textBoxOutput.Text += "`n# # # # # # # # # #`n"
	}
            $form.Controls.Add($textBoxOutput)
# Function to add output to the text box
            function Add-Output {
            param ($text)
            $textBoxOutput.AppendText("$text`r`n")
}
# Function to handle the selected function from the dropdown
            $comboBoxFunctions.Add_SelectedIndexChanged({
            $selectedFunction = $comboBoxFunctions.SelectedItem.ToString()
            switch ($selectedFunction) {
        "List All Users Logon History" {
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
        $buttonalluserslastlogondate.Add_Click({$Output = Get-ADUser -Filter * -Properties LastLogon | Select-Object Name, DistinguishedName, SID, @{Name="LastLogon"; Expression={[DateTime]::FromFileTime($_.LastLogon)}} | Out-GridView -Title "Search results for all users last logon date." -PassThru | Export-Csv -Path "C:\Temp\Users Logon History Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All users Last Logon Date Report.csv" -NoTypeInformation -Delimiter ";"})
        #Write-host "List All users Last Logon Date Report, has been exported to 'C:\Temp\Users Logon History Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All users Last Logon Date Report.csv'"
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
}
        "List User Logon History" {
# Create a form
            $form = New-Object System.Windows.Forms.Form
            $form.BackColor = [System.Drawing.Color]::AliceBlue
            $form.StartPosition = "CenterScreen"
            $form.Text = "User Logon History"
            $form.Width = 300
            $form.Height = 150  # Increased height to accommodate more space for text box

# Create a label
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10, 20)
            $label.Size = New-Object System.Drawing.Size(200, 20)
            $label.Text = "Enter AD Username:"
            $form.Controls.Add($label)

# Create a text box for input
            $textbox = New-Object System.Windows.Forms.TextBox
            $textbox.Location = New-Object System.Drawing.Point(10, 40)
            $textbox.Size = New-Object System.Drawing.Size(250, 20)
            $form.Controls.Add($textbox)

# Create a button
            $button = New-Object System.Windows.Forms.Button
            $button.Location = New-Object System.Drawing.Point(10, 70)
            $button.Size = New-Object System.Drawing.Size(75, 23)
            $button.Text = "OK"
            $button.Add_Click({$Output = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624; StartTime=(Get-Date).AddDays(-30)} | Where-Object { $_.Properties[5].Value -eq $textbox.text } | Format-Table -AutoSize | Out-String
            $textBoxOutput[0].AppendText($output)
})
            $form.Controls.Add($button)

            # Show the form
            $form.ShowDialog() | Out-Null
}
        "Backup All Group Policies" {
            $backupPath = "C:\Temp\GPO_Policy_Backups\"
            if (-not (Test-Path $backupPath)) {
                New-Item -ItemType Directory -Path $backupPath | Out-Null
            }
            Backup-GPO -All -Path $backupPath
            Add-Output "All Group Policies have been backed up to $backupPath."
        }
        "Check AD Replication" {$output = repadmin /replsum | Format-list | Out-string
            # Display results in text output box
            $textBoxOutput[0].AppendText($output)
}
        "List Disabled User Accounts" {
            $disabledUsers = Get-ADUser -Filter { Enabled -eq $false } -Property DisplayName, Enabled
            $disabledUsers | Out-GridView -Title "Disabled User Accounts"
            Add-Output "Disabled User Accounts displayed in GridView."
        }
        "List All Users Password Last Reset" {
            $Output = Get-ADUser -Filter * -Properties PasswordLastSet | Sort PasswordLastSet -Descending | Select Name,Enabled,PasswordLastSet | Out-String
            $textBoxOutput[0].AppendText($output)
        }
        "List User Password Last Reset" {
        # Function to create a GUI window for input
            function Show-InputWindow {
            param (
            [string]$Message
    )

            Add-Type -AssemblyName System.Windows.Forms

            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Enter Username"
            $form.Width = 300
            $form.Height = 150
            $form.StartPosition = "CenterScreen"
            $form.FormBorderStyle = 'FixedDialog'

            $label = New-Object System.Windows.Forms.Label
            $label.Text = $Message
            $label.AutoSize = $true
            $label.Location = New-Object System.Drawing.Point(10, 20)
            $form.Controls.Add($label)

            $textbox = New-Object System.Windows.Forms.TextBox
            $textbox.Location = New-Object System.Drawing.Point(10, 50)
            $textbox.Size = New-Object System.Drawing.Size(260, 20)
            $form.Controls.Add($textbox)

            $button = New-Object System.Windows.Forms.Button
            $button.Location = New-Object System.Drawing.Point(100, 80)
            $button.Size = New-Object System.Drawing.Size(100, 30)
            $button.Text = "Check"
            $button.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.AcceptButton = $button
            $form.Controls.Add($button)

            $form.Topmost = $true

            $form.Add_Shown({$textbox.Select()})
            $result = $form.ShowDialog()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $textbox.Text
    }
}

# Prompt user to enter username
            $username = Show-InputWindow -Message "Enter the username to check password last reset."

        if ($username) {
# Retrieve the user object from Active Directory
        $user = Get-ADUser -Identity $username -Properties PasswordLastSet

        if ($user -ne $null) {
        # Check if PasswordLastSet is not null
        if ($user.PasswordLastSet -ne $null) {
            # Convert PasswordLastSet to a readable format
            $lastResetDateTime = $user.PasswordLastSet 

            # Display result in a message box
            $textBoxOutput[0].AppendText("$username `nLast Password Reset: $lastResetDateTime")
        }
        else {
            $textBoxOutput[0].AppendText("Password last reset information not available for user '$username'.", "Information Not Available")
        }
    }
    else {
        $textBoxOutput[0].AppendText("User '$username' not found in Active Directory.", "User Not Found")
         }
    }
}
        "List All Installed Hotfixes" {
            $output =  Get-HotFix | Format-Table -AutoSize | Out-String
            $textBoxOutput[0].AppendText($output)
        }
        "List Locked User Accounts" {
            $lockedUsers = Search-ADAccount -LockedOut
            $lockedUsers | Out-GridView -Title "Locked User Accounts"
            Add-Output "Locked User Accounts displayed in GridView."
        }
        "List All Objects Deleted in past month" {
# Import the necessary assemblies
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

# Function to create the GUI and get the user selection
    function Show-SelectionForm {
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Select Time Frame"
            $form.Size = New-Object System.Drawing.Size(300, 220)
            $form.StartPosition = "CenterScreen"
            $form.BackColor = [System.Drawing.Color]::AliceBlue
            $form.Forecolor="#000000"

            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Select the time frame:"
            $label.Size = New-Object System.Drawing.Size(250, 20)
            $label.Location = New-Object System.Drawing.Point(25, 20)
            $form.Controls.Add($label)

            $radio30 = New-Object System.Windows.Forms.RadioButton
            $radio30.Text = "30 days"
            $radio30.Location = New-Object System.Drawing.Point(25, 50)
            $form.Controls.Add($radio30)

            $radio60 = New-Object System.Windows.Forms.RadioButton
            $radio60.Text = "60 days"
            $radio60.Location = New-Object System.Drawing.Point(25, 80)
            $form.Controls.Add($radio60)

            $radio90 = New-Object System.Windows.Forms.RadioButton
            $radio90.Text = "90 days"
            $radio90.Location = New-Object System.Drawing.Point(25, 110)
            $form.Controls.Add($radio90)

            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "OK"
            $okButton.Location = New-Object System.Drawing.Point(100, 140)
            $okButton.Add_Click({
             if ($radio30.Checked) { $global:timeFrame = "30 days" }
             elseif ($radio60.Checked) { $global:timeFrame = "60 days" }
             elseif ($radio90.Checked) { $global:timeFrame = "90 days" }
            $form.Close()
        })
            $form.Controls.Add($okButton)

            $form.ShowDialog() | Out-Null
}

# Get the user selection from the GUI
            Show-SelectionForm

          if (-not $global:timeFrame) {
            [System.Windows.Forms.MessageBox]::Show("No time frame selected. Exiting script.", "Error")
            return
}

# Convert the user input into a time span
            switch ($global:timeFrame) {
            "30 days" { $timeSpan = (Get-Date).AddDays(-30) }
            "60 days" { $timeSpan = (Get-Date).AddDays(-60) }
            "90 days" { $timeSpan = (Get-Date).AddDays(-90) }
}

# Get deleted AD objects within the specified time frame
            $deletedObjects = Get-ADObject -Filter {IsDeleted -eq $true -and WhenChanged -ge $timeSpan} -IncludeDeletedObjects -Property WhenChanged, LastKnownParent, DistinguishedName, Name

# Check if any deleted objects were found
          if ($deletedObjects.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No deleted objects found within the specified time frame.", "Results")
            return
}

# Output the deleted objects and the users who deleted them
            $output = @()
         foreach ($object in $deletedObjects) {
            $distinguishedName = $object.DistinguishedName
            $auditLogs = Get-WinEvent -FilterHashTable @{LogName='Security'; ID=4726; StartTime=$timeSpan} |
            Where-Object { $_.Properties[0].Value -like "*$($distinguishedName)*" }

         foreach ($log in $auditLogs) {
            $user = $log.Properties[5].Value
            $output += [PSCustomObject]@{
            DeletedObject = $object.Name
            DeletedBy     = $user
            WhenDeleted   = $object.WhenChanged
        }
    }
}

# Debugging Output: Display the captured deleted objects
            $debugOutput = $deletedObjects | Format-Table -Property Name, DistinguishedName, WhenChanged | Out-String
            $textBoxOutput[0].AppendText($debugOutput)

# Output the results
         if ($output.Count -gt 0) {
            $textBoxOutput[0].AppendText($output)
        }
}
            "Export All DNS Zone Records" {
# Define output directory
            $outputDirectory = "C:\Temp\DNSExports"
# Create the output directory if it doesn't exist
            if (-not (Test-Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
            $dnsRecords = @()
            $zones = Get-DnsServerZone
            foreach ($zone in $zones) {
            $zoneInfo = Get-DnsServerResourceRecord -ZoneName $zone.ZoneName
            foreach ($info in $zoneInfo) {
            $timestamp = if ($info.Timestamp) { $info.Timestamp } else { "static" }
            $timetolive = $info.TimeToLive.TotalSeconds
            $recordData = switch ($info.RecordType) {
                        'A' { $info.RecordData.IPv4Address }
                        'CNAME' { $info.RecordData.HostnameAlias }
                        'NS' { $info.RecordData.NameServer }
                        'SOA' { "[$($info.RecordData.SerialNumber)] $($info.RecordData.PrimaryServer), $($info.RecordData.ResponsiblePerson)" }
                        'SRV' { $info.RecordData.DomainName }
                        'PTR' { $info.RecordData.PtrDomainName }
                        'MX' { $info.RecordData.MailExchange }
                        'AAAA' { $info.RecordData.IPv6Address }
                        'TXT' { $info.RecordData.DescriptiveText }
            default { $null }
   }
            $dnsRecords += [pscustomobject]@{
                        Name       = $zone.ZoneName
                        Hostname   = $info.Hostname
                        Type       = $info.RecordType
                        Data       = $recordData
                        Timestamp  = $timestamp
                        TimeToLive = $timetolive
        }
    }
}
            $dnsRecords | Export-Csv "$outputDirectory\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_AllDNSZonesRecords.csv" -NoTypeInformation -Encoding utf8
            $dnsRecords = "DNS Zones have been exported to C:\Temp\DNSExports"
# Display a pop-up window
            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup($dnsRecords, 0, "Export Completed", 0x1)
    }
        "List All AD Users" {
            $adUsers = Get-ADUser -Filter * -Property DisplayName, SamAccountName | Out-GridView -Title "AD Users" -PassThru
            Add-Output "AD Users displayed in GridView."
        }
        "List All Domain Controllers" {$Output = Get-ADDomainController -Filter * | Select-Object Name, Domain, IPvAddress, Forest, ComputerObjectDN, OperationMasterRoles, OperatingSystem | Format-list | Out-String 
            $textBoxOutput[0].AppendText($output)            
        }
        "List All Domain SSL Certificates" {
        # Add your implementation for listing SSL certificates
            Add-Output "Listing of Domain SSL Certificates not implemented yet."
        }
        "List All Security Groups" {
            $securityGroups = Get-ADGroup -Filter *
            $securityGroups | Out-GridView -Title "Security Groups"
            Add-Output "Security Groups displayed in GridView."
        }
        "Search Group Policies" {
# Create a form
            $gposearchform = New-Object System.Windows.Forms.Form
            $gposearchform.Text = "Search AD Group Policies"
            $gposearchform.BackColor = [System.Drawing.Color]::AliceBlue
            $gposearchform.Forecolor="#000000"
            $gposearchform.Size = New-Object System.Drawing.Size(300, 370)
            $gposearchform.StartPosition = "CenterScreen"
            $gposearchform.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
# Create Form Header Label
            $formheaderlabel = New-Object System.Windows.Forms.Label
            $formheaderlabel.Font = New-Object System.Drawing.Font("Arial Black", 12,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
            $formheaderlabel.Text = "Search Group Policies"
            $formheaderlabel.Location = New-Object System.Drawing.Point(10, 180)
            $formheaderlabel.AutoSize = $true
            $gposearchform.Controls.Add($formheaderlabel)
# Create label for input
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10, 240)
            $label.Size = New-Object System.Drawing.Size(200, 20)
            $label.Text = "Enter search criteria:"
            $gposearchform.Controls.Add($label)
# Create text box for input
            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Location = New-Object System.Drawing.Point(10, 280)
            $textBox.Size = New-Object System.Drawing.Size(280, 20)
            $gposearchform.Controls.Add($textBox)
# Create submit button
            $gposearchbutton = New-Object System.Windows.Forms.Button
            $gposearchbutton.Location = New-Object System.Drawing.Point(10, 320)
            $gposearchbutton.Size = New-Object System.Drawing.Size(80, 30)
            $gposearchbutton.Text = "Submit"
            $gposearchbutton.BackColor = [System.Drawing.Color]::White
            $gposearchbutton.ForeColor = [System.Drawing.Color]::Black
            $gposearchbutton.Add_Click({
# Get input from the text box
            $searchCriteria = $textBox.Text
# Search Active Directory group policies
            $groupPolicies = Get-GPO -All | Where-Object { $_.DisplayName -like "*$searchCriteria*" } | Select-Object DisplayName, DomainName, GpoStatus, CreationTime, Modification, WmiFilter | Out-GridView -Title "Search results for $searchCriteria" -Passthru
    })
            $gposearchform.Controls.Add($gposearchbutton)
            $gposearchform.Controls.Add($searchCriteria)
#$gposearchform.Controls.Add($searchCriteria)

# Add Exit button control to Password Form
            $exitButton = New-Object System.Windows.Forms.Button
            $exitButton.Location = New-Object System.Drawing.Point(210, 320)
            $exitButton.Name = "exitButton"
            $exitButton.Size = New-Object System.Drawing.Size(80, 30)
            $exitButton.BackColor = [System.Drawing.Color]::White
            $exitButton.ForeColor = "#000000"
            $exitButton.TabIndex = 7
            $exitButton.Text = "Exit"
            $exitButton.Add_Click({
            $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Exit Confirmation", "YesNo", "Question")
        if ($confirm -eq "Yes") {
            $gposearchform.Close()
    }
    })
            $gposearchform.Controls.Add($exitButton)
# Create Picture Box
            $pictureBox = New-Object System.Windows.Forms.PictureBox
            $pictureBox.Location = New-Object System.Drawing.Point(10, 10)
            $imageURL = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/GPO%20Search%20GUI/GPO%20Search.png"
            $image = [System.Net.WebRequest]::Create($imageUrl).GetResponse().GetResponseStream()
            $pictureBox.Image = [System.Drawing.Image]::FromStream($image)
            $pictureBox.SizeMode = "Stretch"  # Fit the image within the PictureBox
            $pictureBox.Size = New-Object System.Drawing.Size(300, 150)
            $gposearchform.Controls.Add($pictureBox)
# Show the form
            $gposearchform.ShowDialog() | Out-Null
        }
            "Generate Passwords" {
            function Create-GUI {
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Password Generator"
            $form.Width = 500
            $form.Height = 400
            $form.BackColor = [System.Drawing.Color]::AliceBlue
            $form.ForeColor = [System.Drawing.Color]::Black
            $form.StartPosition = "CenterScreen"
            $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

            $formheaderlabel = New-Object System.Windows.Forms.Label
            $formheaderlabel.Font = New-Object System.Drawing.Font("Arial Black", 12,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
            $formheaderlabel.Text = "Password Generator v1.0"
            $formheaderlabel.Location = New-Object System.Drawing.Point(10, 20)
            $formheaderlabel.AutoSize = $true
            $form.Controls.Add($formheaderlabel)

            $dropdown = New-Object System.Windows.Forms.ComboBox
            $dropdown.Location = New-Object System.Drawing.Point(10, 80)
            $dropdown.Size = New-Object System.Drawing.Size(80, 30)
            1..100 | ForEach-Object { [void]$dropdown.Items.Add($_) }
            $form.Controls.Add($dropdown)
  
# Create Dropdown Label
            $labeldropdown = New-Object System.Windows.Forms.Label
            $labeldropdown.Text = "Password Length:"
            $labeldropdown.Location = New-Object System.Drawing.Point(95, 83)
            $labeldropdown.AutoSize = $true
            $form.Controls.Add($labeldropdown)

# Create Numbers checkbox
            $Numberscheckbox = New-Object System.Windows.Forms.CheckBox
            $Numberscheckbox.Text = "Include Numbers"
            $Numberscheckbox.Location = New-Object System.Drawing.Point(10, 130)
            $Numberscheckbox.Size = New-Object System.Drawing.Size(180, 30)
            $form.Controls.Add($Numberscheckbox)

# Create Uppercase Checkbox
            $UppercaseCheckBox = New-Object System.Windows.Forms.CheckBox
            $UppercaseCheckBox.Location = New-Object System.Drawing.Point(10, 170)
            $UppercaseCheckBox.AutoSize = $true
            $UppercaseCheckBox.Text = "Include Uppercase"
            $form.Controls.Add($UppercaseCheckBox)

# Create Lowercase Checkbox
            $LowercaseCheckBox = New-Object System.Windows.Forms.CheckBox
            $LowercaseCheckBox.Location = New-Object System.Drawing.Point(10, 210)
            $LowercaseCheckBox.AutoSize = $true
            $LowercaseCheckBox.Text = "Include Lowercase"
            $form.Controls.Add($LowercaseCheckBox)

# Create Special Characters
            $specialCharsCheckBox = New-Object System.Windows.Forms.CheckBox
            $specialCharsCheckBox.Location = New-Object System.Drawing.Point(10, 250)
            $specialCharsCheckBox.AutoSize = $true
            $specialCharsCheckBox.Text = "Include Special Characters"
            $form.Controls.Add($specialCharsCheckBox)

            $outputTextBox = New-Object System.Windows.Forms.TextBox
            $outputTextBox.Location = New-Object System.Drawing.Point(10, 300)
            $outputTextBox.Size = New-Object System.Drawing.Size(400, 30)
            $form.Controls.Add($outputTextBox)

# Create Generate Password button
            $generateButton = New-Object System.Windows.Forms.Button
            $generateButton.Text = "Generate Password"
            $generateButton.Location = New-Object System.Drawing.Point(10, 350)
            $generateButton.BackColor = [System.Drawing.Color]::White
            $generateButton.ForeColor = "#000000"
            $generateButton.Size = New-Object System.Drawing.Size(112, 30)

# Add click event handler for Generate button
            $generateButton.Add_Click({
# Function to generate password based on selected options
            function Generate-Password {
            $length = $dropdown.SelectedItem
            $characters = @()
            if ($LowercaseCheckBox.Checked) { $characters += [char[]]'abcdefghijklmnopqrstuvwxyz' }
            if ($UppercaseCheckBox.Checked) { $characters += [char[]]'ABCDEFGHIJKLMNOPQRSTUVWXYZ' }
            if ($Numberscheckbox.Checked) { $characters += [char[]]'0123456789' }
            if ($specialCharsCheckBox.Checked) { $characters += [char[]]'!@#$%^&*()-_=+[]{};:,.<>?/' }

            $password = ''
            for ($i = 0; $i -lt $length; $i++) {
            $password += $characters[(Get-Random -Minimum 0 -Maximum $characters.Length)]
            }
            $outputTextBox.Text = $password
        }

# Generate password
            Generate-Password
        })
            $form.Controls.Add($generateButton)
# Add Password copy button control to Password Form
            $copyButton = New-Object System.Windows.Forms.Button
            $copyButton.Location = New-Object System.Drawing.Point(150, 350)
            $copyButton.Name = "copyButton"
            $copyButton.Size = New-Object System.Drawing.Size(120, 30)
            $copyButton.BackColor = [System.Drawing.Color]::White
            $copyButton.ForeColor = "#000000"
            $copyButton.TabIndex = 7
            $copyButton.Text = "Copy to Clipboard"
            $copyButton.Add_Click({
            [System.Windows.Forms.Clipboard]::SetText($outputTextBox[0].Text)
        })
            $form.Controls.Add($copyButton)
# Add Exit button control to Password Form
            $exitButton = New-Object System.Windows.Forms.Button
            $exitButton.Location = New-Object System.Drawing.Point(400, 350)
            $exitButton.Name = "exitButton"
            $exitButton.Size = New-Object System.Drawing.Size(80, 30)
            $exitButton.BackColor = [System.Drawing.Color]::White
            $exitButton.ForeColor = "#000000"
            $exitButton.TabIndex = 7
            $exitButton.Text = "Exit"
            $exitButton.Add_Click({
            $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Exit Confirmation", "YesNo", "Question")
        if ($confirm -eq "Yes") {
            $form.Close()
    }
        })
            $form.Controls.Add($exitButton)
# Create Picture Box
            $pictureBox = New-Object System.Windows.Forms.PictureBox
            $pictureBox.Location = New-Object System.Drawing.Point(250, 70)
            $imageURL = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Password%20Generator%20GUI/PWD_Logo.png"
            $image = [System.Net.WebRequest]::Create($imageUrl).GetResponse().GetResponseStream()
            $pictureBox.Image = [System.Drawing.Image]::FromStream($image)
            $pictureBox.SizeMode = "Stretch"  # Fit the image within the PictureBox
            $pictureBox.Size = New-Object System.Drawing.Size(200, 200)
            $form.Controls.Add($pictureBox)

            $form.ShowDialog() | Out-Null
}
# Continue running the script in the background
Create-GUI
}
        "Unlock AD User Accounts" {
# Create a form
            $lockedUsersform = New-Object System.Windows.Forms.Form
            $lockedUsersform.Text = "Locked User Account Management"
            $lockedUsersform.Size = New-Object System.Drawing.Size(370,170)
            $lockedUsersform.StartPosition = "CenterScreen"
            $lockedUsersform.BackColor = [System.Drawing.Color]::AliceBlue

# Create Password Length label
            $lockedUserslabel = New-Object System.Windows.Forms.Label
            $lockedUserslabel.Text = "Locked User Accounts"
            $lockedUserslabel.Location = New-Object System.Drawing.Point(10, 20)
            $lockedUserslabel.AutoSize = $true

# Create a dropdown list for locked user accounts
            $lockedUsersComboBox = New-Object System.Windows.Forms.ComboBox
            $lockedUsersComboBox.Location = New-Object System.Drawing.Point(10,50)
            $lockedUsersComboBox.Size = New-Object System.Drawing.Size(200,30)

# Get locked user accounts
            $lockedUsers = Search-ADAccount -LockedOut | Select-Object Name, SamAccountName | Sort-Object Name
foreach ($user in $lockedUsers) {
            $lockedUsersComboBox.Items.Add($user.SamAccountName)
}
# Create a button to unlock the selected user account
            $unlockButton = New-Object System.Windows.Forms.Button
            $unlockButton.BackColor = [System.Drawing.Color]::White
            $unlockButton.Location = New-Object System.Drawing.Point(10,90)
            $unlockButton.Size = New-Object System.Drawing.Size(100,30)
            $unlockButton.Text = "Unlock Account"
            $unlockButton.Add_Click({
            $selectedUser = $lockedUsersComboBox.SelectedItem
      if ($selectedUser) {
            Unlock-ADAccount -Identity $selectedUser
            [System.Windows.Forms.MessageBox]::Show("User account '$selectedUser' has been unlocked.")
            Enable-ADAccount -Identity $selectedUser
            [System.Windows.Forms.MessageBox]::Show("User account '$selectedUser' has been Enabled.")
      } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a user account.")
    }
})
# Add controls to the form
            $lockedUsersform.Controls.Add($lockedUserslabel)
            $lockedUsersform.Controls.Add($lockedUsersComboBox)
            $lockedUsersform.Controls.Add($unlockButton)

# Display the form
            $lockedUsersform.ShowDialog() | Out-Null
}
        "List Inactive Users (30 Days)" {
            $inactiveUsers = Search-ADAccount -AccountInactive -TimeSpan 30.00:00:00
            $inactiveUsers | Out-GridView -Title "Inactive Users (30 Days)"
            Add-Output "Inactive Users for the past 30 days displayed in GridView."
        }
        "List Inactive Users (60 Days)" {
            $inactiveUsers = Search-ADAccount -AccountInactive -TimeSpan 60.00:00:00
            $inactiveUsers | Out-GridView -Title "Inactive Users (60 Days)"
            Add-Output "Inactive Users for the past 60 days displayed in GridView."
        }
        "List Inactive Users (90 Days)" {
            $inactiveUsers = Search-ADAccount -AccountInactive -TimeSpan 90.00:00:00
            $inactiveUsers | Out-GridView -Title "Inactive Users (90 Days)"
            Add-Output "Inactive Users for the past 90 days displayed in GridView."
        }
        "List Users Created in the Past Month" {
            $startDate = (Get-Date).AddMonths(-1)
            $newUsers = Get-ADUser -Filter { WhenCreated -ge $startDate } -Property Name, WhenCreated | Select-Object Name, WhenCreated | Out-String
            $textBoxOutput[0].AppendText($newUsers)
        }
    }
})
# Create buttons
            $buttonHeight = 30
            $buttonWidth = 100
            $buttonPadding = 10
# Export to CSV button
            $buttonExport = New-Object System.Windows.Forms.Button
            $buttonExport.Text = "Export to CSV"
            $buttonExport.BackColor = [System.Drawing.Color]::White
            $buttonExport.Forecolor="#000000"
            $buttonExport.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $buttonExport.Location = New-Object System.Drawing.Point(10, 720)
            $form.Controls.Add($buttonExport)
            $buttonExport.Add_Click({
            $outputDirectory = "C:\Temp\ADReports\"
        if (-not (Test-Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }
            $csvPath = "$outputDirectory\ADReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $textBoxOutput.Lines | ConvertFrom-Csv | Export-Csv -Path $csvPath -NoTypeInformation
            Add-Output "Output exported to $csvPath."
})
# Clear text button
            $buttonClear = New-Object System.Windows.Forms.Button
            $buttonClear.Text = "Clear Text"
            $buttonClear.BackColor = [System.Drawing.Color]::White
            $buttonClear.Forecolor="#000000"
            $buttonClear.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $buttonClear.Location = New-Object System.Drawing.Point(140, 720)
            $form.Controls.Add($buttonClear)
            $buttonClear.Add_Click({
            $textBoxOutput.Clear()
})
# Exit button
            $buttonExit = New-Object System.Windows.Forms.Button
            $buttonExit.Text = "Exit"
            $buttonExit.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
            $buttonExit.BackColor = [System.Drawing.Color]::White
            $buttonExit.Forecolor="#000000"
            $buttonExit.Location = New-Object System.Drawing.Point(670, 720)
            $form.Controls.Add($buttonExit)
            $buttonExit.Add_Click({
            $form.Close()
})
# Show the form
$form.ShowDialog() | Out-Null
