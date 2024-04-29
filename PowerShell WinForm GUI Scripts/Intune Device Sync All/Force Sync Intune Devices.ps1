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

#region Check if Active Directory module is installed
$MsGraphModule = Get-Module -Name ActiveDirectory -ListAvailable
#
# If not installed, install Microsoft Graph module
if (-not $MsGraphModule) {
    Write-Host "Installing Microsoft Graph module..."
    Install-Module Microsoft.Graph.Intune -Force

    #Import Microsoft Graph module
    Import-Module Microsoft.Graph.Intune
}
#endregion Check if Active Directory module is installed

#region Create form
$Form = New-Object System.Windows.Forms.Form
$Form.ShowIcon = $false
$Form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Show
$Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Form.BackColor = [System.Drawing.Color]::AliceBlue
$Form.Text = "Force Sync Intune Devices"
$Form.Size = New-Object System.Drawing.Size(400,550)
$Form.MaximizeBox = $false
$Form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
#endregion Create form

#region Create Header Picture Box
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Location = New-Object System.Drawing.Point(10, 30)
    $imageURL = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Intune%20Device%20Sync%20All/Header3.png"
    $image = [System.Net.WebRequest]::Create($imageUrl).GetResponse().GetResponseStream()
    $pictureBox.Image = [System.Drawing.Image]::FromStream($image)
    $pictureBox.SizeMode = "Stretch"  # Fit the image within the PictureBox
    $pictureBox.Size = New-Object System.Drawing.Size(370, 70)
    $form.Controls.Add($pictureBox)
#endregion Create Picture Box

#region Create Windows Image Box
# Create a picture box (image button)
$windowsimageBox = New-Object System.Windows.Forms.PictureBox
$windowsimageBox.Size = New-Object System.Drawing.Size(100, 100)
$windowsimageBox.Location = New-Object System.Drawing.Point(20, 120)
$windowsimageBox.SizeMode = "Stretch"
$windowsimageBox.add_MouseEnter($handlerEnter)
$windowsimageBox.add_MouseLeave($handlerLeave)
$form.Controls.Add($windowsimageBox)

# Download the image from the URL
$imageUrl = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Intune%20Device%20Sync%20All/windows.png"
$image = [System.Drawing.Image]::FromStream((New-Object System.Net.WebClient).OpenRead($imageUrl))


# Set the image for the picture box
$windowsimageBox.Image = $image

# Add event handlers for hover effects
$handlerEnter = {
    $this.Width = $this.Width + 5
    $this.Height = $this.Height + 5
}

$handlerLeave = {
    $this.Width = $this.Width - 5
    $this.Height = $this.Height - 5
}

$windowsimageBox.Add_Click({
        Connect-MsGraph
        $outputtextbox.AppendText("You are now connecting to MSGraph, please sign in.`r`n")
        # Open new window for initiating sync
        $windowssyncForm = New-Object System.Windows.Forms.Form
        $windowssyncForm.ShowIcon = $false
        $windowssyncForm.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Show
        $windowssyncForm.BackColor = [System.Drawing.Color]::AliceBlue
        $windowssyncForm.Text = "Initiate Intune Windows Device Sync"
        $windowssyncForm.Size = New-Object System.Drawing.Size(400, 260)
        $windowssyncForm.StartPosition = "CenterScreen"
        $windowssyncForm.MaximizeBox = $false
        $windowssyncForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        
        # Add label
        $windowssyncLabel = New-Object System.Windows.Forms.Label
        $windowssyncLabel.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
        $windowssyncLabel.Text = "Initiate Intune Windows Device Sync"
        $windowssyncLabel.AutoSize = $true
        $windowssyncLabel.Location = New-Object System.Drawing.Point(10, 10)
        $windowssyncForm.Controls.Add($windowssyncLabel)
        
        # Add sync button
        $windowssyncButton = New-Object System.Windows.Forms.Button
        $windowssyncButton.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
        $windowssyncButton.Text = "Sync"
        $windowssyncButton.Size = New-Object System.Drawing.Size(50, 30)
        $windowssyncButton.BackColor = "#01a1ee"
        $windowssyncButton.ForeColor = [System.Drawing.Color]::White
        $windowssyncButton.Location = New-Object System.Drawing.Point(10, 40)
        $windowssyncButton.Add_Click({
            $windowsprogressBar = New-Object System.Windows.Forms.ProgressBar
            $windowsprogressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
            $windowsprogressBar.Location = New-Object System.Drawing.Point(70, 40)
            $windowsprogressBar.Size = New-Object System.Drawing.Size(300, 30)
            $windowssyncForm.Controls.Add($windowsprogressBar)
            
            # Simulating sync process
            $windowsprogressBar.Visible = $true
            Start-Sleep -Seconds 5
            $windowstextBox.AppendText("Syncing Windows Intune devices...`r`n")
            # Simulating completion
            Start-Sleep -Seconds 5
            $windowstextBox.AppendText("Sync completed.`r`n")
            $windowsprogressBar.Visible = $false
        })
        $windowssyncForm.Controls.Add($windowssyncButton)
        
        # Add text output box
        $windowstextBox = New-Object System.Windows.Forms.TextBox
        $windowstextBox.Multiline = $true
        $windowstextBox.ScrollBars = "Vertical"
        $windowstextBox.Location = New-Object System.Drawing.Point(10, 85)
        $windowstextBox.Size = New-Object System.Drawing.Size(360, 100)
        $windowssyncForm.Controls.Add($windowstextBox)

        $exitbutton = New-Object System.Windows.Forms.Button
        $exitbutton.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
        $exitbutton.Location = New-Object System.Drawing.Point(320, 205)
        $exitbutton.Name = "formexitbutton"
        $exitbutton.Size = New-Object System.Drawing.Size(50, 30)
        $exitbutton.BackColor = "#01a1ee"
        $exitbutton.ForeColor = [System.Drawing.Color]::White
        $exitbutton.TabIndex = 8
        $exitbutton.Text = "Exit"
        $exitbutton.Add_Click({
        $windowssyncForm.Close()
        
})
        $windowssyncForm.Controls.Add($exitbutton)

        $windowssyncForm.ShowDialog() | Out-Null
    })
    
#endregion Create Windows Image Box

#region Create Android Image Box
# Create a picture box (image button)
$androidimageBox = New-Object System.Windows.Forms.PictureBox
$androidimageBox.Size = New-Object System.Drawing.Size(100, 100)
$androidimageBox.Location = New-Object System.Drawing.Point(150, 120)
$androidimageBox.SizeMode = "Stretch"
$androidimageBox.add_MouseEnter($handlerEnter)
$androidimageBox.add_MouseLeave($handlerLeave)
$form.Controls.Add($androidimageBox)

# Download the image from the URL
$imageUrl = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Intune%20Device%20Sync%20All/android.png"
$image = [System.Drawing.Image]::FromStream((New-Object System.Net.WebClient).OpenRead($imageUrl))


# Set the image for the picture box
$androidimageBox.Image = $image

# Add event handlers for hover effects
$handlerEnter = {
    $this.Width = $this.Width + 5
    $this.Height = $this.Height + 5
}

$handlerLeave = {
    $this.Width = $this.Width - 5
    $this.Height = $this.Height - 5
}

$androidimageBox.Add_Click({Connect-MsGraph
        $outputtextbox.AppendText("You are now connecting to MSGraph, please sign in.`r`n")
        # Open new window for initiating sync
        $androidsyncForm = New-Object System.Windows.Forms.Form
        $androidsyncForm.ShowIcon = $false
        $androidsyncForm.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Show
        $androidsyncForm.BackColor = [System.Drawing.Color]::AliceBlue
        $androidsyncForm.Text = "Initiate Intune Android Device Sync"
        $androidsyncForm.Size = New-Object System.Drawing.Size(400, 260)
        $androidsyncForm.StartPosition = "CenterScreen"
        $androidsyncForm.MaximizeBox = $false
        $androidsyncForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

        
        # Add label
        $androidsyncLabel = New-Object System.Windows.Forms.Label
        $androidsyncLabel.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
        $androidsyncLabel.Text = "Initiate Intune Android Device Sync"
        $androidsyncLabel.AutoSize = $true
        $androidsyncLabel.Location = New-Object System.Drawing.Point(10, 10)
        $androidsyncForm.Controls.Add($androidsyncLabel)
        
        # Add sync button
        $androidsyncButton = New-Object System.Windows.Forms.Button
        $androidsyncButton.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
        $androidsyncButton.Text = "Sync"
        $androidsyncButton.Size = New-Object System.Drawing.Size(50, 30)
        $androidsyncButton.BackColor = "#01a1ee"
        $androidsyncButton.ForeColor = [System.Drawing.Color]::White
        $androidsyncButton.Location = New-Object System.Drawing.Point(10, 40)
        $androidsyncButton.Add_Click({
            $androidprogressBar = New-Object System.Windows.Forms.ProgressBar
            $androidprogressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
            $androidprogressBar.Location = New-Object System.Drawing.Point(70, 40)
            $androidprogressBar.Size = New-Object System.Drawing.Size(300, 30)
            $androidsyncForm.Controls.Add($androidprogressBar)
            
            # Simulating sync process
            $androidprogressBar.Visible = $true
            Start-Sleep -Seconds 5
            $androidtextBox.AppendText("Syncing all Android Intune devices...`r`n")
            # Simulating completion
            Start-Sleep -Seconds 5
            $androidtextBox.AppendText("Sync completed.`r`n")
            $androidprogressBar.Visible = $false
        })
        $androidsyncForm.Controls.Add($androidsyncButton)
        
        # Add text output box
        $androidtextBox = New-Object System.Windows.Forms.TextBox
        $androidtextBox.Multiline = $true
        $androidtextBox.ScrollBars = "Vertical"
        $androidtextBox.Location = New-Object System.Drawing.Point(10, 85)
        $androidtextBox.Size = New-Object System.Drawing.Size(360, 100)
        $androidsyncForm.Controls.Add($androidtextBox)

        $exitbutton = New-Object System.Windows.Forms.Button
        $exitbutton.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
        $exitbutton.Location = New-Object System.Drawing.Point(320, 205)
        $exitbutton.Name = "formexitbutton"
        $exitbutton.Size = New-Object System.Drawing.Size(50, 30)
        $exitbutton.BackColor = "#01a1ee"
        $exitbutton.ForeColor = [System.Drawing.Color]::White
        $exitbutton.TabIndex = 8
        $exitbutton.Text = "Exit"
        $exitbutton.Add_Click({
        $androidsyncForm.Close()
        
})
        $androidsyncForm.Controls.Add($exitbutton)

        $androidsyncForm.ShowDialog() | Out-Null
    })
#endregion Create Andoid Image Box

#region Create iOS Image Box
# Create a picture box (image button)
$iosimageBox = New-Object System.Windows.Forms.PictureBox
$iosimageBox.Size = New-Object System.Drawing.Size(80, 80)
$iosimageBox.Location = New-Object System.Drawing.Point(280, 125)
$iosimageBox.SizeMode = "Stretch"
$iosimageBox.add_MouseEnter($handlerEnter)
$iosimageBox.add_MouseLeave($handlerLeave)
$form.Controls.Add($iosimageBox)

# Download the image from the URL
$imageUrl = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Intune%20Device%20Sync%20All/ios.png"
$image = [System.Drawing.Image]::FromStream((New-Object System.Net.WebClient).OpenRead($imageUrl))


# Set the image for the picture box
$iosimageBox.Image = $image

# Add event handlers for hover effects
$handlerEnter = {
    $this.Width = $this.Width + 5
    $this.Height = $this.Height + 5
}

$handlerLeave = {
    $this.Width = $this.Width - 5
    $this.Height = $this.Height - 5
}

$iosimageBox.Add_Click({Connect-MsGraph
        $outputtextbox.AppendText("You are now connecting to MSGraph, please sign in.`r`n")
        # Open new window for initiating sync
        $iossyncForm = New-Object System.Windows.Forms.Form
        $iossyncForm.ShowIcon = $false
        $iossyncForm.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Show
        $iossyncForm.BackColor = [System.Drawing.Color]::AliceBlue
        $iossyncForm.Text = "Initiate Intune Device Sync"
        $iossyncForm.Size = New-Object System.Drawing.Size(400, 260)
        $iossyncForm.StartPosition = "CenterScreen"
        $iossyncForm.MaximizeBox = $false
        $iossyncForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        
        # Add label
        $iossyncLabel = New-Object System.Windows.Forms.Label
        $iossyncLabel.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
        $iossyncLabel.Text = "Initiate Intune iOS Device Sync"
        $iossyncLabel.AutoSize = $true
        $iossyncLabel.Location = New-Object System.Drawing.Point(10, 10)
        $iossyncForm.Controls.Add($iossyncLabel)
        
        # Add sync button
        $iossyncButton = New-Object System.Windows.Forms.Button
        $iossyncButton.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
        $iossyncButton.Text = "Sync"
        $iossyncButton.Size = New-Object System.Drawing.Size(50, 30)
        $iossyncButton.BackColor = "#01a1ee"
        $iossyncButton.ForeColor = [System.Drawing.Color]::White
        $iossyncButton.Location = New-Object System.Drawing.Point(10, 40)
        $iossyncButton.Add_Click({
            $iosprogressBar = New-Object System.Windows.Forms.ProgressBar
            $iosprogressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
            $iosprogressBar.Location = New-Object System.Drawing.Point(70, 40)
            $iosprogressBar.Size = New-Object System.Drawing.Size(300, 30)
            $iossyncForm.Controls.Add($iosprogressBar)
            
            # Simulating sync process
            $iosprogressBar.Visible = $true
            Start-Sleep -Seconds 5
            $textBox.AppendText("Syncing all iOS Intune devices...`r`n")
            # Simulating completion
            Start-Sleep -Seconds 5
            $textBox.AppendText("Sync completed.`r`n")
            $iosprogressBar.Visible = $false
        })
        $iossyncForm.Controls.Add($iossyncButton)
        
        # Add text output box
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Location = New-Object System.Drawing.Point(10, 85)
        $textBox.Size = New-Object System.Drawing.Size(360, 100)
        $iossyncForm.Controls.Add($textBox)

        $exitbutton = New-Object System.Windows.Forms.Button
        $exitbutton.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
        $exitbutton.Location = New-Object System.Drawing.Point(320, 205)
        $exitbutton.Name = "formexitbutton"
        $exitbutton.Size = New-Object System.Drawing.Size(50, 30)
        $exitbutton.BackColor = "#01a1ee"
        $exitbutton.ForeColor = [System.Drawing.Color]::White
        $exitbutton.TabIndex = 8
        $exitbutton.Text = "Exit"
        $exitbutton.Add_Click({
        $iossyncForm.Close()
        
})
        $iossyncForm.Controls.Add($exitbutton)

        $iossyncForm.ShowDialog() | Out-Null
    })
#endregion Create iOS Image Box

#region Output Text Box
#
$outputtextbox = New-Object System.Windows.Forms.RichTextBox
$outputtextbox.Location = New-Object System.Drawing.Point(10, 240)
$outputtextbox.AutoSize = $true
$outputtextbox.Name = "outputtextbox"
$outputtextbox.Size = New-Object System.Drawing.Size(380, 260)
$outputtextbox.Multiline = $true
$outputtextbox.ScrollBars = "Vertical"
$outputtextbox.TabIndex = 1
$outputtextbox.Text = ""
$outputtextbox.Font = New-Object System.Drawing.Font("Consolas", 10)
function Add-outputtextbox{
		[CmdletBinding()]
		param ($text)
		$outputtextbox.Text += "$text;"
		$outputtextbox.Text += "`n# # # # # # # # # #`n"
	}
$form.Controls.Add($outputtextbox)
#endregion Output Text Box

#region Copy Button
#
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$copyButton.Location = New-Object System.Drawing.Point(10, 510)
$copyButton.Name = "copyButton"
$copyButton.Size = New-Object System.Drawing.Size(130, 30)
$copyButton.BackColor = "#01a1ee"
$copyButton.ForeColor = [System.Drawing.Color]::White
$copyButton.TabIndex = 7
$copyButton.Text = "Copy to Clipboard"
$copyButton.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($outputtextbox[0].Text)
})
$Form.Controls.Add($copyButton)
#endregion Copy Button

#region Form Exit Button
#
$formexitbutton = New-Object System.Windows.Forms.Button
$formexitbutton.Font = New-Object System.Drawing.Font("Arial", 9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$formexitbutton.Location = New-Object System.Drawing.Point(340, 510)
$formexitbutton.Name = "formexitbutton"
$formexitbutton.Size = New-Object System.Drawing.Size(50, 30)
$formexitbutton.BackColor = "#01a1ee"
$formexitbutton.ForeColor = [System.Drawing.Color]::White
$formexitbutton.TabIndex = 8
$formexitbutton.Text = "Exit"
$formexitbutton.Add_Click({$confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Exit Confirmation", "YesNo", "Question")
    if ($confirm -eq "Yes") {
        $form.Close()
    }
})
$Form.Controls.Add($formexitbutton)
#endregion Form Exit Button

# Show form
$form.ShowDialog() | Out-Null
