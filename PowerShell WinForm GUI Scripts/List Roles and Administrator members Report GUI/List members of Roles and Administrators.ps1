# Connect-AzureAD

# Create the Windows Form
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure EntraID Roles and Administrators"
$form.Size = New-Object System.Drawing.Size(810, 830)
$form.BackColor = [System.Drawing.Color]::LightBlue
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$form.ShowIcon = $false

# Helper function to set hover effects
function Set-HoverEffect {
    param (
        [System.Windows.Forms.Button]$button
    )

    $script:originalBackColor = $button.BackColor
    $script:originalForeColor = $button.ForeColor

    $button.Add_MouseEnter({
        param ($sender, $eventArgs)
        $sender.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#156996")
        $sender.ForeColor = [System.Drawing.Color]::White
    })

    $button.Add_MouseLeave({
        param ($sender, $eventArgs)
        $sender.BackColor = $script:originalBackColor
        $sender.ForeColor = $script:originalForeColor
    })
}

# Create Picture Box
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Point(110, 40)
$imageURL = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/List%20Roles%20and%20Administrator%20members%20Report%20GUI/Header.png"
$image = [System.Drawing.Image]::FromFile($imageURL)
$pictureBox.Image = $image
$pictureBox.SizeMode = "Stretch"  # Fit the image within the PictureBox
$pictureBox.Size = New-Object System.Drawing.Size(500, 120)
$form.Controls.Add($pictureBox)

$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(780, 550)
$listView.Location = New-Object System.Drawing.Point(10, 190)
$listView.View = [System.Windows.Forms.View]::Details
$listView.BackColor = [System.Drawing.Color]::White
$listView.ForeColor = "#156996"
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Text = ""
$listView.Font = New-Object System.Drawing.Font("Consolas", 10)

# Add columns to the ListView
$listView.Columns.Add("Role", 200)
$listView.Columns.Add("User", 580)

$form.Controls.Add($listView)

# Authenticate to Azure AD
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "Connect to Azure AD"
$connectButton.Location = New-Object System.Drawing.Point(10, 750)
$connectButton.AutoSize = $true
$connectButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$connectButton.BackColor = [System.Drawing.Color]::White
$connectButton.ForeColor = [System.Drawing.Color]::Black
$connectButton.Add_Click({
    try {
        Connect-AzureAD
        $connectButton.Text = "Connected"
        $connectButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
        $connectButton.BackColor = [System.Drawing.Color]::Green
        $connectButton.ForeColor = [System.Drawing.Color]::White
        $connectButton.Enabled = $false
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to Azure AD. Please try again.", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($connectButton)

# Fetch and display roles and members
$fetchButton = New-Object System.Windows.Forms.Button
$fetchButton.BackColor = [System.Drawing.Color]::White
$fetchButton.ForeColor = [System.Drawing.Color]::Black
$fetchButton.Text = "Fetch Roles and Administrators"
$fetchButton.Location = New-Object System.Drawing.Point(150, 750)
$fetchButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$fetchButton.AutoSize = $true
$fetchButton.Enabled = $true
$fetchButton.Add_Click({
    $roles = Get-AzureADDirectoryRole | Sort-Object DisplayName
    $listView.Items.Clear()
    
    foreach ($role in $roles) {
        $roleMembers = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId | Sort-Object UserPrincipalName
        foreach ($member in $roleMembers) {
            if ($null -ne $member -and $null -ne $member.UserPrincipalName) {
                $item = New-Object System.Windows.Forms.ListViewItem($role.DisplayName)
                $item.SubItems.Add($member.UserPrincipalName)
                $listView.Items.Add($item)
            }
        }
    }
})
$form.Controls.Add($fetchButton)

# Export to CSV Button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(343, 750)
$exportButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$exportButton.Name = "exportButton"
$exportButton.AutoSize = $true
$exportButton.BackColor = [System.Drawing.Color]::White
$exportButton.ForeColor = [System.Drawing.Color]::Black
$exportButton.Text = "Export to CSV"
$exportButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save CSV File"
    $saveFileDialog.ShowDialog()

    if ($saveFileDialog.FileName -ne "") {
        $csvData = @()
        foreach ($item in $listView.Items) {
            $csvData += [PSCustomObject]@{
                Role = $item.Text
                User = $item.SubItems[1].Text
            }
        }
        $csvData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Delimiter ";"
        [System.Windows.Forms.MessageBox]::Show("Data exported successfully.", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$form.Controls.Add($exportButton)

# Add Exit button control to Password Form
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(715, 750)
$exitButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$exitButton.Name = "exitButton"
$exitButton.AutoSize = $true
$exitButton.BackColor = [System.Drawing.Color]::White
$exitButton.ForeColor = [System.Drawing.Color]::Black
$exitButton.Text = "Exit"
$exitButton.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Exit Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $form.Close()
    }
})
$form.Controls.Add($exitButton)

# Apply hover effects to buttons
Set-HoverEffect -button $connectButton
Set-HoverEffect -button $fetchButton
Set-HoverEffect -button $exportButton
Set-HoverEffect -button $exitButton

# Show the form
$form.Add_Shown({ $form.Activate() })
[void] $form.ShowDialog()
