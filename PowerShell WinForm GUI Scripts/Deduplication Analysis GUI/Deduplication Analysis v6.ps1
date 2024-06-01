Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore

# Create transcript log folder if not exist
$folderPath = "C:\RTAdmin\Reports Toolset Files\Transcript logs"
if (-not (Test-Path -Path $folderPath)) {  
    New-Item -Path $folderPath -ItemType Directory
}


# Start transcript log  file and write to path C:\RTAdmin\
Start-Transcript -Path "C:\RTAdmin\Reports Toolset Files\Transcript logs\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Log1.txt"

function Create-DeduplicationAnalysisForm {
    param (
        [string]$initialDirectory
    )

    # Create form object
    $form = New-Object System.Windows.Forms.Form
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Text = "Deduplication Analysis v1.0"
    $form.Size = New-Object System.Drawing.Size(1000, 850)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    #$form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -id $pid | Select-Object -ExpandProperty Path))
    $form.ShowIcon = $false

    # Create Picture Box
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Location = New-Object System.Drawing.Point(540, 10)
    $imageURL = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Deduplication%20Analysis%20GUI/Header.png"
    $image = [System.Net.WebRequest]::Create($imageUrl).GetResponse().GetResponseStream()
    $pictureBox.Image = [System.Drawing.Image]::FromStream($image)
    $pictureBox.SizeMode = "Stretch"  # Fit the image within the PictureBox
    $pictureBox.Size = New-Object System.Drawing.Size(410, 195)
    $form.Controls.Add($pictureBox)
    
    # Create Browse Button
    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(20, 135)
    $browseButton.Size = New-Object System.Drawing.Size(100, 30)
    $browseButton.Text = "Browse"
    $browseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
        if ($folderBrowser.ShowDialog() -eq 'OK') {
            $inputBox.Text = $folderBrowser.SelectedPath
        }
    })
    $form.Controls.Add($browseButton)

    # Create Input Box for selected Path
    $inputBox = New-Object System.Windows.Forms.TextBox
    $inputBox.Location = New-Object System.Drawing.Point(130, 144)
    $inputBox.Size = New-Object System.Drawing.Size(400, 30)
    $form.Controls.Add($inputBox)

    # Create Progress Bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 180)
    $progressBar.Size = New-Object System.Drawing.Size(510, 20)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.ForeColor = [System.Drawing.Color]::Green
    $form.Controls.Add($progressBar)

    # Create Run Button
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Location = New-Object System.Drawing.Point(330, 20)
    $runButton.Size = New-Object System.Drawing.Size(90, 30)
    $runButton.Text = "Run Analysis"
    $runButton.BackColor = [System.Drawing.Color]::LightGray
    $runButton.ForeColor = [System.Drawing.Color]::Black
    $runButton.Add_Click({
        function Get-FileSizeMB {
            param ($file)
            $size = (Get-Item $file).Length / 1MB
            return [math]::Round($size, 2)
        }

        function Convert-Size {
            param (
                [double]$sizeMB
            )
            if ($sizeMB -lt 1024) {
                return "$sizeMB MB"
            } else {
                $sizeGB = $sizeMB / 1024
                return [math]::Round($sizeGB, 2).ToString() + " GB"
            }
        }

        $folderPath = $inputBox.Text
        $files = Get-ChildItem -Path $folderPath -File -Recurse
        $duplicateFiles = @{}
        $totalFiles = 0
        $totalFileSize = 0
        $totalDuplicateSize = 0

        $outputGrid.Rows.Clear()
        $progressBar.Value = 0

        $runButton.Text = "Running"
        $runButton.BackColor = [System.Drawing.Color]::Green
        $runButton.ForeColor = [System.Drawing.Color]::White

        $pauseButton.Enabled = $true

        $totalFilesCount = $files.Count
        $progressIncrement = 100 / $totalFilesCount

        foreach ($file in $files) {
            $totalFiles++
            $fileName = $file.Name
            $filePath = $file.FullName
            $fileSizeMB = Get-FileSizeMB -file $file.FullName

            if ($duplicateFiles.ContainsKey($fileName)) {
                $duplicateFiles[$fileName] += ,$file
                $totalDuplicateSize += $fileSizeMB
            } else {
                $duplicateFiles[$fileName] = ,$file
                $totalFileSize += $fileSizeMB
            }

            $rowIndex = $outputGrid.Rows.Add($fileName, $filePath, $fileSizeMB, "Scan Completed")
            $outputGrid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::White

            $progressBar.Value += $progressIncrement
            [System.Windows.Forms.Application]::DoEvents()
        }

        foreach ($fileName in $duplicateFiles.Keys) {
            if ($duplicateFiles[$fileName].Count -gt 1) {
                foreach ($file in $duplicateFiles[$fileName]) {
                    $rowIndex = $outputGrid.Rows.Add($file.Name, $file.FullName, (Get-FileSizeMB -file $file.FullName), "Duplicate")
                    $outputGrid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::Red
                }
            }
        }

        $totalSizeFormatted = Convert-Size -sizeMB ($totalFileSize + $totalDuplicateSize)
        $duplicateSizeFormatted = Convert-Size -sizeMB $totalDuplicateSize
        $savingSizeFormatted = Convert-Size -sizeMB $totalDuplicateSize

        $drive = Get-PSDrive -Name (Get-Item $folderPath).PSDrive.Name
        $usedSpaceMB = [math]::Round(($drive.Used / 1MB), 2)
        $freeSpaceMB = [math]::Round(($drive.Free / 1MB), 2)
        $totalSizeMB = [math]::Round((($drive.Used + $drive.Free) / 1MB), 2)

        $usedSpaceFormatted = if ($usedSpaceMB -lt 1024) { "$usedSpaceMB MB" } else { "$([math]::Round($usedSpaceMB / 1024, 2)) GB" }
        $freeSpaceFormatted = if ($freeSpaceMB -lt 1024) { "$freeSpaceMB MB" } else { "$([math]::Round($freeSpaceMB / 1024, 2)) GB" }
        $totalSizeFormatted = if ($totalSizeMB -lt 1024) { "$totalSizeMB MB" } else { "$([math]::Round($totalSizeMB / 1024, 2)) GB" }

        $resultsLabel.Text = "Total number of files: $totalFiles`n" +
                             "Total size of all files: $totalSizeFormatted`n" +
                             "Total size of duplicate files: $duplicateSizeFormatted`n" +
                             "File size saving if duplicates removed: $savingSizeFormatted`n" +
                             "Current disk size: $totalSizeFormatted`n" +
                             "Current used space: $usedSpaceFormatted`n" +
                             "Current free space: $freeSpaceFormatted"

        $progressBar.Value = 100

        $runButton.Text = "Run Analysis"
        $runButton.BackColor = [System.Drawing.Color]::Gray
        $pauseButton.Enabled = $false
    })
    $form.Controls.Add($runButton)

    # Create Pause Button
    $pauseButton = New-Object System.Windows.Forms.Button
    $pauseButton.Location = New-Object System.Drawing.Point(330, 60)
    $pauseButton.Size = New-Object System.Drawing.Size(90, 30)
    $pauseButton.Text = "Pause"
    $pauseButton.BackColor = [System.Drawing.Color]::Orange
    $pauseButton.ForeColor = [System.Drawing.Color]::White
    $pauseButton.Enabled = $true
    $pauseButton.Add_Click({
    $pauseButton.Text = "Paused"
    $pauseButton.BackColor = [System.Drawing.Color]::Gray
    })
    $form.Controls.Add($pauseButton)

    # Output grid view
    $outputGrid = New-Object System.Windows.Forms.DataGridView
    $outputGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $outputGrid.Location = New-Object System.Drawing.Point(20, 210)
    $outputGrid.Size = New-Object System.Drawing.Size(940, 550)
    $outputGrid.ReadOnly = $true
    $outputGrid.AllowUserToAddRows = $false
    $outputGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $outputGrid.ColumnHeadersVisible = $true
    $outputGrid.RowHeadersVisible = $false
    $outputGrid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    $outputGrid.DefaultCellStyle.BackColor = [System.Drawing.Color]::White  # Set default back color to white
    $outputGrid.BackgroundColor = [System.Drawing.Color]::White # Ensure the entire background is white
    $form.Controls.Add($outputGrid)

    $outputGrid.Columns.Add("FileName", "File Name")
    $outputGrid.Columns.Add("Path", "Path")
    $outputGrid.Columns.Add("Size", "Size (MB)")
    $outputGrid.Columns.Add("Status", "Status")

    # Create Results Panel
    $resultsPanel = New-Object System.Windows.Forms.Panel
    $resultsPanel.BackColor = [System.Drawing.Color]::White
    $resultsPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $resultsPanel.Location = New-Object System.Drawing.Point(20, 20)
    $resultsPanel.Size = New-Object System.Drawing.Size(300, 105)
    $resultsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $form.Controls.Add($resultsPanel)

    # Create Results Label
    $resultsLabel = New-Object System.Windows.Forms.Label
    $resultsLabel.Size = $resultsPanel.Size
    $resultsLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $resultsPanel.Controls.Add($resultsLabel)

    # Create Export button
    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "Export"
    $exportButton.Location = New-Object System.Drawing.Point(20, 770)
    $exportButton.Size = New-Object System.Drawing.Size(80, 30)
    $exportButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|PDF files (*.pdf)|*.pdf|HTML files (*.html)|*.html"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $content = "Path,Owner,Permissions,Last Accessed,File Size`n"
        $listView.Items | ForEach-Object {
            $content += "$($_.SubItems[0].Text),$($_.SubItems[1].Text),$($_.SubItems[2].Text),$($_.SubItems[3].Text),$($_.SubItems[4].Text)`n"
        }
        $content | Out-File -FilePath $saveFileDialog.FileName
         }
    })
    $form.Controls.Add($exportButton)

    # Create Enable DeDuplication Feature button
    $enableDedupButton = New-Object System.Windows.Forms.Button
    $enableDedupButton.Text = "Enable DeDuplication Feature"
    $enableDedupButton.Location = New-Object System.Drawing.Point(110, 770)
    $enableDedupButton.Size = New-Object System.Drawing.Size(190, 30)
    $enableDedupButton.Add_Click({
            # Define the command to run
            $command = "Install-WindowsFeature -Name FS-Data-Deduplication"

            # Open a new PowerShell window and run the command
            Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", $command
    })            

    $form.Controls.Add($enableDedupButton)

   # Create Backup Duplicate Files button
$backupButton = New-Object System.Windows.Forms.Button
$backupButton.Text = "Backup Duplicate Files"
$backupButton.Location = New-Object System.Drawing.Point(310, 770)
$backupButton.Size = New-Object System.Drawing.Size(180, 30)
$backupButton.Add_Click({
    $backupFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $backupFolderDialog.Description = "Select backup folder"
    if ($backupFolderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $backupFolderPath = $backupFolderDialog.SelectedPath

        Write-Host "Selected backup folder: $backupFolderPath"

        # Loop through DataGridView rows
        foreach ($row in $outputGrid.Rows) {
            $fileName = $row.Cells["FileName"].Value
            $filePath = $row.Cells["Path"].Value
            $status = $row.Cells["Status"].Value

            # Check if status is "Duplicate" (assuming status is set as "Duplicate" for red rows)
            if ($status -eq "Duplicate") {
                Write-Host "Copying duplicate file and directory structure: $filePath to $backupFolderPath"

                # Copy the entire directory structure recursively
                $directoryPath = [System.IO.Path]::GetDirectoryName($filePath)
                $relativePath = $directoryPath.Replace((Get-Item $inputBox.Text).FullName, "").TrimStart("\")
                $destinationDir = Join-Path -Path $backupFolderPath -ChildPath $relativePath

                if (-not (Test-Path -Path $destinationDir)) {
                    New-Item -ItemType Directory -Path $destinationDir | Out-Null
                }

                # Copy the duplicate file itself
                $destinationFile = Join-Path -Path $destinationDir -ChildPath $fileName
                Copy-Item -Path $filePath -Destination $destinationFile -Force
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Backup completed.", "Backup", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$form.Controls.Add($backupButton)



    # Create Exit button
    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.BackColor = [System.Drawing.Color]::Red
    $exitButton.ForeColor = [System.Drawing.Color]::White
    $exitButton.Text = "Exit"
    $exitButton.Location = New-Object System.Drawing.Point(892, 770)
    $exitButton.Size = New-Object System.Drawing.Size(70, 30)
    $exitButton.Add_Click({
    $form.Close()
    })
    $form.Controls.Add($exitButton)

    $form.ShowDialog() | Out-Null
}

Create-DeduplicationAnalysisForm
