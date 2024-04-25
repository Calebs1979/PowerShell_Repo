Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to generate a random password
function Generate-Password {
    param (
        [int]$Length,
        [bool]$UseAlphaNumerics,
        [bool]$UseSpecialChars
    )

    $alphaNumerics = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
    $specialChars = '!@#$%^&*()_+-=[]{}|;:,.<>?'

    $chars = $alphaNumerics
    if ($UseAlphaNumerics -and $UseSpecialChars) {
        $chars += $specialChars
    } elseif ($UseSpecialChars) {
        $chars = $specialChars
    }

    $password = ''
    $random = New-Object System.Random
    for ($i = 0; $i -lt $Length; $i++) {
        $password += $chars[$random.Next(0, $chars.Length)]
    }

    return $password
}

# Function to create GUI
function Create-GUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Password Generator"
    $form.Width = 500
    $form.Height = 400
    $form.BackColor = "#0d3659"
    $form.ForeColor = "#FFFFFF"
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

    $dropdownlabel = New-Object System.Windows.Forms.Label
    $dropdownlabel.Font = New-Object System.Drawing.Font("Arial Black", 12,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
    $dropdownlabel.Text = "Password Generator v1.0"
    $dropdownlabel.Location = New-Object System.Drawing.Point(10, 20)
    $dropdownlabel.AutoSize = $true
    $form.Controls.Add($dropdownlabel)

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
    $checkboxNumbers = New-Object System.Windows.Forms.CheckBox
    $checkboxNumbers.Text = "Include Numbers"
    $checkboxNumbers.Location = New-Object System.Drawing.Point(10, 130)
    $checkboxNumbers.Size = New-Object System.Drawing.Size(180, 30)
    $form.Controls.Add($checkboxNumbers)

    # Create Uppercase Checkbox
    $UppercaseCheckBox = New-Object System.Windows.Forms.CheckBox
    $UppercaseCheckBox.Location = New-Object System.Drawing.Point(10, 170)
    #$UppercaseCheckBox.Size = New-Object System.Drawing.Size(180, 30)
    $UppercaseCheckBox.AutoSize = $true
    $UppercaseCheckBox.Text = "Include Uppercase"
    $form.Controls.Add($UppercaseCheckBox)

    # Create Lowercase Checkbox
    $LowercaseCheckBox = New-Object System.Windows.Forms.CheckBox
    $LowercaseCheckBox.Location = New-Object System.Drawing.Point(10, 210)
    #$LowercaseCheckBox.Size = New-Object System.Drawing.Size(180, 30)
    $LowercaseCheckBox.AutoSize = $true
    $LowercaseCheckBox.Text = "Include Lowercase"
    $form.Controls.Add($LowercaseCheckBox)

    # Create Special Characters
    $specialCharsCheckBox = New-Object System.Windows.Forms.CheckBox
    $specialCharsCheckBox.Location = New-Object System.Drawing.Point(10, 250)
    #$specialCharsCheckBox.Size = New-Object System.Drawing.Size(180, 30)
    $specialCharsCheckBox.AutoSize = $true
    $specialCharsCheckBox.Text = "Include Special Characters"
    $form.Controls.Add($specialCharsCheckBox)

    $outputTextBox = New-Object System.Windows.Forms.TextBox
    $outputTextBox.Location = New-Object System.Drawing.Point(10, 300)
    $outputTextBox.Size = New-Object System.Drawing.Size(400, 30)
    $form.Controls.Add($outputTextBox)

    $generateButton = New-Object System.Windows.Forms.Button
    $generateButton.Location = New-Object System.Drawing.Point(10, 350)
    $generateButton.AutoSize = $true
    #$generateButton.Size = New-Object System.Drawing.Size(120, 30)
    $generateButton.BackColor = [System.Drawing.Color]::White
    $generateButton.ForeColor = "#000000"
    $generateButton.Text = "Generate Password"
    $generateButton.Add_Click({
        $length = $dropdown.SelectedItem
        $UppercaseCheckBox = $UppercaseCheckBox.Checked
        $LowercaseCheckBox = $LowercaseCheckBox.Checked
        $checkboxNumbers = $checkboxNumbers.Checked
        $specialCharsCheckBox = $specialCharsCheckBox.Checked
        $outputTextBox.Text = Generate-Password -Length $length -UseUppercase $UppercaseCheckBox -UseLowercase $LowercaseCheckBox -UsecheckboxNumbers $checkboxNumbers -UsespecialCharsCheckBox $specialCharsCheckBox
    })
    $form.Controls.Add($generateButton)

    # Add Password copy button control to Password Form
    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Location = New-Object System.Drawing.Point(150, 350)
    $copyButton.Name = "copyButton"
    #$copyButton.Size = New-Object System.Drawing.Size(117, 30)
    $copyButton.AutoSize = $true
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
    $exitButton.Location = New-Object System.Drawing.Point(420, 350)
    $exitButton.Name = "exitButton"
    #$exitButton.Size = New-Object System.Drawing.Size(70, 30)
    $exitButton.AutoSize = $true
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
    $imageURL = "C:\ccadmin\Password Generator GUI\Logo.v1.png"
    $image = [System.Net.WebRequest]::Create($imageUrl).GetResponse().GetResponseStream()
    $pictureBox.Image = [System.Drawing.Image]::FromStream($image)
    $pictureBox.SizeMode = "Stretch"  # Fit the image within the PictureBox
    $pictureBox.Size = New-Object System.Drawing.Size(200, 200)
    $form.Controls.Add($pictureBox)

    $form.ShowDialog() | Out-Null
}

# Run the GUI
Create-GUI
