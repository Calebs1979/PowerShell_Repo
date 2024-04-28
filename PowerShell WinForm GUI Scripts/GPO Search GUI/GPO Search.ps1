Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
 
 # Create a form
            $gposearchform = New-Object System.Windows.Forms.Form
            $gposearchform.Text = "Search AD Group Policies"
            $gposearchform.BackColor = [System.Drawing.Color]::AliceBlue
            $gposearchform.Forecolor="#000000"
            $gposearchform.Size = New-Object System.Drawing.Size(300, 370)
            $gposearchform.StartPosition = "CenterScreen"
            $gposearchform.BackColor = "#0d3659"
            $gposearchform.ForeColor = "#FFFFFF"
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
