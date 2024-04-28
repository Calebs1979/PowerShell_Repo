Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Enable Visual Styles
[Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion Enable Visual Styles

#region Get the current logged-on user and display informational window
$currentUserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
# Extract only the username from the full user name
$currentUserName = $currentUserName -replace ".*\\" 
# Capitalize the first letter of the username
$capitalizedUserName = $currentUserName.Substring(0,1).ToUpper() + $currentUserName.Substring(1)
# Create the message
$message = "Hello, $capitalizedUserName!`n`nPlease ensure that you create a strong password, based on Microsoft's recommendations.`n`n* At least 12 characters long but 14 or more is better.`n`n* A combination of uppercase letters, lowercase letters, numbers, and symbols.`n`n* Not a word that can be found in a dictionary or the name of a person, character, product, or organization.`n`nThanks you $capitalizedUserName"
# Display the popup window
Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show($message, "Microsoft's password complexity recommendations ", "OK", "Information")
#endregion Get the current logged-on user and display informational window

# Password Generator

#regionFunction to create GUI
function Create-GUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Password Generator"
    $form.Width = 500
    $form.Height = 400
    $form.BackColor = "#0d3659"
    $form.ForeColor = "#FFFFFF"
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
    $imageURL = "https://raw.githubusercontent.com/Calebs1979/PowerShell_Repo/main/PowerShell%20WinForm%20GUI%20Scripts/Password%20Generator/Logo.v1.png"
    $image = [System.Net.WebRequest]::Create($imageUrl).GetResponse().GetResponseStream()
    $pictureBox.Image = [System.Drawing.Image]::FromStream($image)
    $pictureBox.SizeMode = "Stretch"  # Fit the image within the PictureBox
    $pictureBox.Size = New-Object System.Drawing.Size(200, 200)
    $form.Controls.Add($pictureBox)

    $form.ShowDialog() | Out-Null
}

# Continue running the script in the background
Create-GUI