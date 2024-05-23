# Ensure necessary assemblies are loaded
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$form.BackColor = "#063970"
$form.Text = "Permission Scanner v3.0"
$form.Size = New-Object System.Drawing.Size(992, 706)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $true
$form.MinimizeBox = $true
$form.ShowIcon = $false

# Create Picture Box
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Point(775, 10)
$imageURL = "C:\ccadmin\Disk Permission Scanner v1.0\PermissionsReporter.png"
$image = [System.Net.WebRequest]::Create($imageUrl).GetResponse().GetResponseStream()
$pictureBox.Image = [System.Drawing.Image]::FromStream($image)
$pictureBox.SizeMode = "Stretch"  # Fit the image within the PictureBox
$pictureBox.Size = New-Object System.Drawing.Size(170, 120)
$form.Controls.Add($pictureBox)

# Create controls
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Select Folder/Drive:"
$pathLabel.ForeColor = "#ffffff"
$pathLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$pathLabel.Location = New-Object System.Drawing.Point(20, 20)
$pathLabel.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($pathLabel)

$pathComboBox = New-Object System.Windows.Forms.ComboBox
$pathComboBox.Location = New-Object System.Drawing.Point(150, 20)
$pathComboBox.Size = New-Object System.Drawing.Size(500, 20)
$pathComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($pathComboBox)

# Populate ComboBox with Drives
Get-PSDrive -PSProvider FileSystem | ForEach-Object {
    $pathComboBox.Items.Add($_.Root)
}

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse"
$browseButton.BackColor = [System.Drawing.Color]::LightGray
$browseButton.Location = New-Object System.Drawing.Point(670, 20)
$browseButton.Size = New-Object System.Drawing.Size(75, 23)
$form.Controls.Add($browseButton)

$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Text = "Scan"
$scanButton.Location = New-Object System.Drawing.Point(20, 60)
$scanButton.Size = New-Object System.Drawing.Size(75, 23)
$scanButton.BackColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($scanButton)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "Stop Scan"
$stopButton.BackColor = [System.Drawing.Color]::LightGray
$stopButton.Location = New-Object System.Drawing.Point(110, 60)
$stopButton.Size = New-Object System.Drawing.Size(75, 23)
$stopButton.Enabled = $false
$form.Controls.Add($stopButton)

$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Exit"
$exitButton.BackColor = [System.Drawing.Color]::LightGray
$exitButton.Location = New-Object System.Drawing.Point(200, 60)
$exitButton.Size = New-Object System.Drawing.Size(75, 23)
$form.Controls.Add($exitButton)

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.View = [System.Windows.Forms.View]::Details
$listView.Location = New-Object System.Drawing.Point(20, 140)
$listView.Size = New-Object System.Drawing.Size(940, 510)
$listView.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Columns.Add("Path", 400)
$listView.Columns.Add("Account", 150)
$listView.Columns.Add("Permissions", 200)
$listView.Columns.Add("Inheritance", 150)  # New column for Inheritance
$listView.Anchor = 'Top, Bottom, Left, Right'
$form.Controls.Add($listView)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export to File"
$exportButton.BackColor = [System.Drawing.Color]::LightGray
$exportButton.Location = New-Object System.Drawing.Point(290, 60)
$exportButton.Size = New-Object System.Drawing.Size(100, 23)
$form.Controls.Add($exportButton)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 100)
$progressBar.Size = New-Object System.Drawing.Size(725, 23)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($progressBar)

# Hidden signature label
$signatureLabel = New-Object System.Windows.Forms.Label
$signatureLabel.Text = "C.S"
$signatureLabel.Location = New-Object System.Drawing.Point(750, 570)
$signatureLabel.Size = New-Object System.Drawing.Size(30, 20)
$signatureLabel.Visible = $false
$form.Controls.Add($signatureLabel)

# Browse button click event
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $pathComboBox.Items.Add($folderBrowser.SelectedPath)
        $pathComboBox.SelectedItem = $folderBrowser.SelectedPath
    }
})

# Define scan function
$global:stopScan = $false

function Scan-Disk {
    param (
        [string]$path
    )

    $listView.Items.Clear()
    $items = Get-ChildItem -Path $path -Recurse -ErrorAction SilentlyContinue
    $totalItems = $items.Count
    $currentItem = 0

    foreach ($item in $items) {
        if ($global:stopScan) {
            break
        }

        while ($global:pauseScan) {
            Start-Sleep -Milliseconds 500
        }

        $currentItem++
        $progressBar.Value = [math]::Round(($currentItem / $totalItems) * 100)
        $acl = Get-Acl -Path $item.FullName
        $acl.Access | ForEach-Object {
            $account = $_.IdentityReference.ToString()
            $permissions = $_.FileSystemRights.ToString()

            # Check if permissions are inherited
            if ($_.IsInherited) {
                $inheritance = "Inherited"
            } else {
                $inheritance = "Not Inherited"
            }

            $listItem = New-Object System.Windows.Forms.ListViewItem($item.FullName)
            $listItem.SubItems.Add($account)
            $listItem.SubItems.Add($permissions)
            $listItem.SubItems.Add($inheritance)  # Add inheritance information

            # Highlight FullControl permissions in red
            if ($permissions -match "FullControl") {
                $listItem.ForeColor = [System.Drawing.Color]::Red
            }

            $listView.Items.Add($listItem)
            # Update the UI
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    # Autoresize columns after populating
    $listView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
}


# Define button click events
$scanButton.Add_Click({
    $path = $pathComboBox.SelectedItem
    if ($path) {
        $scanButton.Text = "Scanning..."
        $scanButton.BackColor = [System.Drawing.Color]::Green
        $scanButton.ForeColor = [System.Drawing.Color]::White
        $scanButton.Enabled = $true
        $stopButton.Enabled = $true
        $global:stopScan = $false
        Scan-Disk -path $path
        $scanButton.Text = "Scan"
        $scanButton.BackColor = [System.Drawing.Color]::LightGray
        $scanButton.ForeColor = [System.Drawing.Color]::Black
        $scanButton.Enabled = $true
        $stopButton.Enabled = $true
        $progressBar.Value = 0
        [System.Windows.Forms.MessageBox]::Show("Scan Completed", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a folder or drive.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$stopButton.Add_Click({
    $global:stopScan = $true
    $stopButton.Enabled = $false
})

$exitButton.Add_Click({
    $form.Close()
})

$exportButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save Scan Results"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvContent = "Path,Account,Permissions`n"
        foreach ($item in $listView.Items) {
            $csvContent += "$($item.Text),$($item.SubItems[1].Text),$($item.SubItems[2].Text)`n"
        }
        $csvContent | Out-File -FilePath $saveFileDialog.FileName
        [System.Windows.Forms.MessageBox]::Show("File saved successfully.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Show the form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
