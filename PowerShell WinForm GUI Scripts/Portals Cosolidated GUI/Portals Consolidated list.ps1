################################################################################ 
#
#  Name    : Caleb Smith   
#  Version : 1.130124
#  Author  :
#  Date    : 2024/04/13
#  Designed using Visual Studio 2022
#  Converted to PowerShell with ConvertForm module version 2.0.0
#  PowerShell version 5.1.19041.4291
#
################################################################################

#regionLoading external assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#endregion assmblies

$Form1 = New-Object System.Windows.Forms.Form

#region components
$generalcomboBox = New-Object System.Windows.Forms.ComboBox
$microsoft365comboBox = New-Object System.Windows.Forms.ComboBox
$intunecomboBox = New-Object System.Windows.Forms.ComboBox
$defendercomboBox = New-Object System.Windows.Forms.ComboBox
$entraidcomboBox = New-Object System.Windows.Forms.ComboBox
$azurecomboBox = New-Object System.Windows.Forms.ComboBox
$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox2 = New-Object System.Windows.Forms.ComboBox
$label1 = New-Object System.Windows.Forms.Label
$label2 = New-Object System.Windows.Forms.Label
$label3 = New-Object System.Windows.Forms.Label
$label4 = New-Object System.Windows.Forms.Label
$label5 = New-Object System.Windows.Forms.Label
$label6 = New-Object System.Windows.Forms.Label
$label7 = New-Object System.Windows.Forms.Label
$button2 = New-Object System.Windows.Forms.Button
$button5 = New-Object System.Windows.Forms.Button
$button6 = New-Object System.Windows.Forms.Button
$button7 = New-Object System.Windows.Forms.Button
$panel1 = New-Object System.Windows.Forms.Panel
$panel2 = New-Object System.Windows.Forms.Panel
$panel3 = New-Object System.Windows.Forms.Panel
$tabControl1 = New-Object System.Windows.Forms.TabControl
$tabPage1 = New-Object System.Windows.Forms.TabPage
$monthCalendar1 = New-Object System.Windows.Forms.MonthCalendar
$webBrowser1 = New-Object System.Windows.Forms.WebBrowser
#endregion components

#regionpanel1
#
$panel1.AutoSize = $true
$panel1.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$panel1.BackColor = [System.Drawing.Color]::FromArgb(44,62,80)
$panel1.Controls.Add($label5)
$panel1.Controls.Add($button4)
$panel1.Controls.Add($button2)
$panel1.Controls.Add($label1)
$panel1.Dock = [System.Windows.Forms.DockStyle]::Top
$panel1.Location = New-Object System.Drawing.Point(0, 0)
$panel1.Name = "panel1"
$panel1.Size = New-Object System.Drawing.Size(1151, 58)
$panel1.TabIndex = 0
#endregion panel1

#regionlabel5 Consolidated Portal Links
#
$label5.Anchor = [System.Windows.Forms.AnchorStyles]::Top
$label5.AutoSize = $true
$label5.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$label5.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 15.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label5.ForeColor = [System.Drawing.Color]::FromArgb(255,128,0)
$label5.Location = New-Object System.Drawing.Point(450, 15)
$label5.Name = "label5"
$label5.Size = New-Object System.Drawing.Size(282, 25)
$label5.TabIndex = 5
$label5.Text = "Consolidated Portal Links"
$label5.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
#endregionlabel5

#regionbutton2 Designer icon button
#
$button2.Anchor = [System.Windows.Forms.AnchorStyles]"Top,Right"
$button2.BackgroundImage= [System.Drawing.Image]::FromFile("C:\ccadmin\Modern Flat UI Design Dashboard\man.png")
$button2.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
$button2.FlatAppearance.BorderSize = 0
$button2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button2.Location = New-Object System.Drawing.Point(1071, 15)
$button2.Name = "Designericonbutton"
$button2.Size = New-Object System.Drawing.Size(43, 40)
$button2.TabIndex = 2
$button2.UseVisualStyleBackColor = $true
function OnClick_button2 {
    # Create the About popup window
    $developerpopupform = New-Object System.Windows.Forms.Form
    $developerpopupform.Text = "Designed By"
    $developerpopupform.Size = New-Object System.Drawing.Size(420, 220)
    $developerpopupform.BackColor = [System.Drawing.Color]::FromArgb(44,62,80)
    $developerpopupform.StartPosition = "CenterScreen"
    
    # Add header label
    $labelHeader1 = New-Object System.Windows.Forms.Label
    $labelHeader1.AutoSize = $true
    $labelHeader1.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $labelHeader1.Location = New-Object System.Drawing.Point(10, 10)
    $developerpopupform.Controls.Add($labelHeader1)
    
    # Add Name label
    $Namelabel = New-Object System.Windows.Forms.Label
    $Namelabel.ForeColor = "#ffffff"
    $Namelabel.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $Namelabel.Text = "Name: Caleb" 
    $Namelabel.AutoSize = $true
    $Namelabel.Location = New-Object System.Drawing.Point(10, 20)

    # Add Surname label
    $Surnamelabel = New-Object System.Windows.Forms.Label
    $Surnamelabel.ForeColor = "#ffffff"
    $Surnamelabel.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $Surnamelabel.Text = "Surname: Smith" 
    $Surnamelabel.AutoSize = $true
    $Surnamelabel.Location = New-Object System.Drawing.Point(250, 20)

    # Add Email label
    $emaillabel = New-Object System.Windows.Forms.Label
    $emaillabel.ForeColor = "#ffffff"
    $emaillabel.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $emaillabel.Text = "Email: Caleb.s1979@gmail.com" 
    $emaillabel.AutoSize = $true
    $emaillabel.Location = New-Object System.Drawing.Point(10, 60)
    
    # Add Label Developer

    $Developerlabel = New-Object System.Windows.Forms.Label
    $Developerlabel.ForeColor = "#ffffff"
    $Developerlabel.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $Developerlabel.Location = New-Object System.Drawing.Point(250, 60)
    $Developerlabel.AutoSize = $true
    $Developerlabel.Text = "Developer: Caleb Smith"

    # Add Label Build version

    $Buildversionlabel = New-Object System.Windows.Forms.Label
    $Buildversionlabel.ForeColor = "#ffffff"
    $Buildversionlabel.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $Buildversionlabel.Location = New-Object System.Drawing.Point(250, 60)
    $Buildversionlabel.AutoSize = $true
    $Buildversionlabel.Text = "Build version 1.130424"

     # Add Label Copyright
    $CopyrightDesignerlabel = New-Object System.Windows.Forms.Label
    $CopyrightDesignerlabel.ForeColor = "#ffffff"
    $CopyrightDesignerlabel.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $CopyrightDesignerlabel.Location = New-Object System.Drawing.Point(10, 100)
    $CopyrightDesignerlabel.AutoSize = $true
    $CopyrightDesignerlabel.Text = "All rights reserved.Copyright © $(Get-Date -Format 'yyyy')."

    # Add labels to the About form
    $developerpopupform.Controls.Add($labelHeader1)
    $developerpopupform.Controls.Add($Namelabel)
    $developerpopupform.Controls.Add($Surnamelabel)
    $developerpopupform.Controls.Add($emaillabel)
    $developerpopupform.Controls.Add($Developerlabel)
    $developerpopupform.Controls.Add($Buildversionlabel)
    $developerpopupform.Controls.Add($CopyrightDesignerlabel)
    $developerpopupform.Controls.Add($Designedbylabel)
    # Show the About popup window
    $developerpopupform.ShowDialog() | Out-Null
}
$button2.Add_Click( { OnClick_button2 } )
#endregion button2

#regionlabel1 Admin Dashboard
#
$label1.CausesValidation = $false
$label1.Font = New-Object System.Drawing.Font("Segoe UI Black", 16,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label1.ForeColor = [System.Drawing.Color]::FromArgb(255,128,0)
$label1.Location = New-Object System.Drawing.Point(12, 15)
$label1.Name = "label1"
$label1.Size = New-Object System.Drawing.Size(268, 34)
$label1.TabIndex = 0
$label1.Text = "Admin Dashboard"
$label1.UseMnemonic = $false
#endregion label1

#regionpanel2 Panel 2
#
$panel2.BackColor = [System.Drawing.Color]::FromArgb(44,62,80)
$panel2.Controls.Add($monthCalendar1)
$panel2.Controls.Add($comboBox2)
$panel2.Controls.Add($label7)
$panel2.Controls.Add($comboBox1)
$panel2.Controls.Add($label6)
$panel2.Controls.Add($label3)
$panel2.Controls.Add($generalcomboBox)
$panel2.Controls.Add($microsoft365comboBox)
$panel2.Controls.Add($intunecomboBox)
$panel2.Controls.Add($defendercomboBox)
$panel2.Controls.Add($entraidcomboBox)
$panel2.Controls.Add($azurecomboBox)
$panel2.Controls.Add($button7)
$panel2.Controls.Add($button6)
$panel2.Controls.Add($button5)
$panel2.Dock = [System.Windows.Forms.DockStyle]::Left
$panel2.Location = New-Object System.Drawing.Point(0, 58)
$panel2.Name = "panel2"
$panel2.Size = New-Object System.Drawing.Size(243, 707)
$panel2.TabIndex = 1
#endregion panel2

#regionlabel3 Portal Links
#
$label3.AutoSize = $true
$label3.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label3.ForeColor = [System.Drawing.Color]::FromArgb(255,128,0)
$label3.Location = New-Object System.Drawing.Point(6, 75)
$label3.Name = "label3"
$label3.Size = New-Object System.Drawing.Size(81, 17)
$label3.TabIndex = 13
$label3.Text = "Portal Links"
#
#regionAdd click action to ComboBox selection
$generalcomboBox.Add_SelectedIndexChanged({
    OpenPortal
})
#endregion Click action to Combobox selection

#regionFunction to open portal
function OpenPortal{
    $selectedPortal = $generalcomboBox.SelectedItem.ToString()
    switch ($selectedPortal) {
        "Microsoft Partner Center" {
            Start-Process "https://partner.microsoft.com/dashboard"
            break
        }
        "training" {
            Start-Process "https://learn.microsoft.com/training/"
            break
        }
        "Volume Licensing Service Center" {
            Start-Process "https://www.microsoft.com/Licensing/servicecenter"
            break
        }
        "AkaSearch" {
            Start-Process "https://akaSearch.net"
            break
        }
    }
}
#endregion Function to open portal
#endregion label3

#region Add click action to ComboBox selection generalcomboBox -  - Open Portal8
$generalcomboBox.Add_SelectedIndexChanged({
    OpenPortal8
})
#endregion Click action to Combobox selection

#region Function to open portal generalcomboBox - Open Portal8
function OpenPortal8{
    $selectedPortal = $generalcomboBox.SelectedItem.ToString()
    switch ($selectedPortal) {
        "Microsoft Partner Center" {
            Start-Process "https://partner.microsoft.com/dashboard"
            break
        }
        "Training" {
            Start-Process "https://learn.microsoft.com/training/"
            break
        }
        "Volume Licensing Service Center" {
            Start-Process "https://www.microsoft.com/Licensing/servicecenter"
            break
        }
        "AkaSearch" {
            Start-Process "https://akaSearch.net"
            break
        }
     }
}
#endregion Function to open portal

#region general comboBox generalcomboBox  - Open Portal8
#
$generalcomboBox.AccessibleDescription = "General"
$generalcomboBox.AccessibleName = "General"
$generalcomboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$generalcomboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$generalcomboBox.FormattingEnabled = $true
$generalcomboBox.Items.AddRange(@(
"Microsoft Partner Center",
"Training",
"Volume Licensing Service Center",
"AkaSearch"))
$generalcomboBox.Location = New-Object System.Drawing.Point(9, 302)
$generalcomboBox.Name = "generalcomboBox"
$generalcomboBox.Size = New-Object System.Drawing.Size(225, 25)
$generalcomboBox.TabIndex = 12
$generalcomboBox.Text = "General"
#endregion general comboBox

#region Add click action to microsoft365 comboBox selection microsoft365comboBox - Open Portal2
$microsoft365comboBox.Add_SelectedIndexChanged({
    OpenPortal2
})
#endregion Click action to Combobox selection

#region Function to open portal microsoft365comboBox  - Open Portal2
function OpenPortal2{
    $selectedPortal = $microsoft365comboBox.SelectedItem.ToString()
    switch ($selectedPortal) {
        "Microsoft 365 Admin" {
            Start-Process "https://admin.microsoft.com"
            break
        }
        "Exchange" {
            Start-Process "https://admin.exchange.microsoft.com"
            break
        }
        "SharePoint" {
            Start-Process "https://admin.microsoft.com/sharepoint"
            break
        }
        "Teams" {
            Start-Process "	https://admin.teams.microsoft.com/"
            break
        }
        "OneDrive" {
            Start-Process "https://admin.onedrive.com"
            break
        }
        "Power BI" {
            Start-Process "https://app.powerbi.com/admin-portal"
            break
        }
        "Power Platform" {
            Start-Process "https://admin.powerplatform.microsoft.com"
            break
        }
        "Power Automate" {
            Start-Process "https://make.powerautomate.com"
            break
        }
        "Yammer" {
            Start-Process "https://www.yammer.com/office365/admin"
            break
        }
        "Graph Explorer" {
            Start-Process "	https://developer.microsoft.com/en-us/graph/graph-explorer"
            break
        }
        "Outlook" {
            Start-Process "https://outlook.office365.com/mail/"
            break
        }
        "Purview Compliance manager" {
            Start-Process "https://compliance.microsoft.com/compliancemanager"
            break
        }
        "Purview App Governance" {
            Start-Process "https://compliance.microsoft.com/cloudapps/app-governance"
            break
        }
        "Purview Audit" {
            Start-Process "https://compliance.microsoft.com/auditlogsearch"
            break
        }
        "Purview Communication Compliance" {
            Start-Process "https://compliance.microsoft.com/supervisoryreview"
            break
        }
        "Purview Data Loss Prevention" {
            Start-Process "https://compliance.microsoft.com/datalossprevention"
            break
        }
        "Purview eDiscovery" {
            Start-Process "https://compliance.microsoft.com/advancedediscovery"
            break
        }
        "Purview Data Lifecycle Management" {
            Start-Process "https://compliance.microsoft.com/informationgovernance"
            break
        }
        "Purview Activity Explorer" {
            Start-Process "https://compliance.microsoft.com/dataclassification/activityexplorer"
            break
        }
        "Purview Content Search" {
            Start-Process "https://compliance.microsoft.com/contentsearchv2"
            break
        }
        "Purview Content Explorer" {
            Start-Process "https://compliance.microsoft.com/dataclassification/contentexplorer"
            break
        }
        "Purview Data Classification" {
            Start-Process "https://compliance.microsoft.com/dataclassificationclassifiers"
            break
        }
        "PowerApps" {
            Start-Process "https://make.powerapps.com/"
            break
        }
        "Teams Web Client" {
            Start-Process "https://teams.microsoft.com"
            break
        }
        "Teams Rooms Pro Management" {
            Start-Process "https://portal.rooms.microsoft.com/"
            break
        }
        "Office Customization Tool" {
            Start-Process "https://config.office.com/deploymentsettings"
            break
        }
        "Microsoft 365 Apps admin center" {
            Start-Process "https://config.office.com/officeSettings"
            break
        }
        "Microsoft Forms" {
            Start-Process "https://forms.office.com/"
            break
        }
        "Microsoft Loop" {
            Start-Process "https://loop.microsoft.com/"
            break
        }
        "Power BI GCC" {
            Start-Process "https://app.powerbigov.us/"
            break
        }
        "PowerAutomate GCC" {
            Start-Process "https://make.gov.powerautomate.us/"
            break
        }
        "PowerPages GCC" {
            Start-Process "https://make.gov.powerpages.microsoft.us/"
            break
        }
        "Windows 365" {
            Start-Process "https://windows365.microsoft.com/"
            break
        }
        "Volume Licensing Keys" {
            Start-Process "https://admin.microsoft.com/Adminportal/Home#/subscriptions/vlnew/downloadsandkeys"
            break
        }
        "Message Center" {
            Start-Process "https://admin.microsoft.com/#/MessageCenter"
            break
        }
    }
}
#endregion Function to open portal

#region microsoft365 comboBox microsoft365comboBox  - Open Portal2
#
$microsoft365comboBox.AccessibleDescription = "Microsoft365"
$microsoft365comboBox.AccessibleName = "Microsoft365"
$microsoft365comboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$microsoft365comboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$microsoft365comboBox.FormattingEnabled = $true
$microsoft365comboBox.Items.AddRange(@(
"Microsoft 365 Admin",
"Exchange",
"SharePoint",
"Teams",
"OneDrive",
"Power BI",
"Power Platform",
"Power Automate",
"Yammer",
"Graph Explorer",
"Outlook",
"Purview Permissions manager",
"Purview Compliance manager",
"Purview App Governance",
"Purview Audit",
"Purview Communication Compliance",
"Purview Data Loss Prevention",
"Purview eDiscovery",
"Purview Data Lifecycle Management",
"Purview Activity Explorer",
"Purview Content Search",
"Purview Content Explorer",
"Purview Data Classification",
"PowerApps",
"Teams Web Client",
"Teams Rooms Pro Management",
"Office Customization Tool",
"Microsoft 365 Apps admin center",
"Microsoft Forms",
"Power BI GCC",
"PowerAutomate GCC",
"PowerPages GCC",
"Windows 365",
"Volume Licensing Keys",
"Message Center"
))
$microsoft365comboBox.Location = New-Object System.Drawing.Point(9, 222)
$microsoft365comboBox.Name = "microsoft365comboBox"
$microsoft365comboBox.Size = New-Object System.Drawing.Size(225, 25)
$microsoft365comboBox.TabIndex = 11
$microsoft365comboBox.Text = "Microsoft365"
#endregion micrsoft365 comboBox

#region Add click action to Intune ComboBox selection intunecomboBox - Open Portal3
$intunecomboBox.Add_SelectedIndexChanged({
    OpenPortal3
})
#endregion Click action to Combobox selection

#region Function to open portal intunecomboBox  - Open Portal3
function OpenPortal3{
    $selectedPortal = $intunecomboBox.SelectedItem.ToString()
    switch ($selectedPortal) {
        "Intune Portal" {
            Start-Process "https://intune.microsoft.com"
            break
        }
        "Devices" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/overview"
            break
        }
        "Apps" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/AppsMenu/~/overview"
            break
        }
        "Security" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_Workflows/SecurityManagementMenu/~/overview"
            break
        }
        "Users" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers"
            break
        }
        "Groups" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_AAD_IAM/GroupsManagementMenuBlade/~/AllGroups"
            break
        }
        "Tenant admin" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/TenantAdminMenu/~/tenantStatus"
            break
        }
        "Windows devices" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesWindowsMenu/~/windowsDevices"
            break
        }
        "iOS/iPad devices" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesIosMenu/~/iosDevices"
            break
        }
        "macOS devices" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMacOsMenu/~/macOsDevices"
            break
        }
        "Android" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesAndroidMenu/~/overview"
            break
        }
        "Enroll devices" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesEnrollmentMenu/~/windowsEnrollment"
            break
        }
        "Compliance policies" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesComplianceMenu/~/policies"
            break
        }
        "Configuration profiles" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/configurationProfiles"
            break
        }
        "Scripts" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/powershell"
            break
        }
        "App configuration policies" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/AppsMenu/~/appConfig"
            break
        }
        "Windows Autopilot devices" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_Enrollment/AutopilotDevices.ReactView/filterOnManualRemediation"
            break
        }
        "Education" {
            Start-Process "https://intuneeducation.portal.azure.com"
            break
        }
        "Proactive Remediations" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/remediations?#view/Microsoft_IntuneDeviceSettings"
            break
        }
        "Windows Update Rings" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/windows10UpdateRings?"
            break
        }"Troubleshooting" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/SupportMenu/~/troubleshooting"
            break
        }
        "Feature updates for Windows" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/featureUpdateDeployments"
            break
        }
        "Driver updates for Windows" {
            Start-Process "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/windowsDriverUpdate?#view/"
            break
        }
    }
}
#endregion Function to open portal

#regionintunecomboBox - comboBox Intune ComboBox - Open Portal3
#
$intunecomboBox.AccessibleDescription = "Intune"
$intunecomboBox.AccessibleName = "Intune"
$intunecomboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$intunecomboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$intunecomboBox.FormattingEnabled = $true
$intunecomboBox.Items.AddRange(@(
"Intune Portal",
"Devices",
"Apps",
"Security",
"Users",
"Groups",
"Tenant admin",
"Windows devices",
"iOS/iPad devices",
"macOS devices",
"Android",
"Enroll devices",
"Compliance policies",
"Configuration profiles",
"Scripts",
"App configuration policies",
"Windows Autopilot devices",
"Education",
"Proactive Remediations",
"Windows Update Rings",
"Troubleshooting",
"Feature updates for Windows",
"Driver updates for Windows"
))
$intunecomboBox.Location = New-Object System.Drawing.Point(9, 262)
$intunecomboBox.Name = "intunecomboBox"
$intunecomboBox.Size = New-Object System.Drawing.Size(225, 25)
$intunecomboBox.TabIndex = 10
$intunecomboBox.Text = "Intune"
#endregion intunecomboBox

#region Add click action to Defender ComboBox selection DefendercomboBox - Open Portal4
$DefendercomboBox.Add_SelectedIndexChanged({
    OpenPortal4
})
#endregion Click action to Combobox selection

#region Function to open portal DefendercomboBox  - Open Portal4
function OpenPortal4{
    $selectedPortal = $DefendercomboBox.SelectedItem.ToString()
    switch ($selectedPortal) {
        "Microsoft 365 Defender" {
            Start-Process "https://security.microsoft.com"
            break
        }
        "Incidents" {
            Start-Process "https://security.microsoft.com/incidents"
            break
        }
        "Hunting" {
            Start-Process "https://security.microsoft.com/v2/advanced-hunting"
            break
        }
        "Action centre" {
            Start-Process "https://security.microsoft.com/action-center"
            break
        }
        "Explorer" {
            Start-Process "https://security.microsoft.com/threatexplorer"
            break
        }
        "Restricted entities" {
            Start-Process "https://security.microsoft.com/restrictedentities"
            break
        }
        "Investigations" {
            Start-Process "https://security.microsoft.com/airinvestigation"
            break
        }
        "Threat Intelligence" {
            Start-Process "https://ti.defender.microsoft.com"
            break
        }
        "Quarantined Emails" {
            Start-Process "https://security.microsoft.com/quarantine"
            break
        }
        "Defender Tenant Allow/Block List" {
            Start-Process "https://security.microsoft.com/tenantAllowBlockList"
            break
        }
        "Microsoft Secure Score" {
            Start-Process "https://security.microsoft.com/securescore"
            break
        }
        "Cloud app catalog" {
            Start-Process "https://security.microsoft.com/cloudapps/app-catalog"
            break
        }
        "Manage OAuth apps" {
            Start-Process "https://security.microsoft.com/cloudapps/oauth-apps"
            break
        }
        "Device Inventory" {
            Start-Process "https://security.microsoft.com/machines"
            break
        }
        "Identities" {
            Start-Process "https://security.microsoft.com/cloudapps/users-and-accounts"
            break
        }
        "Microsoft Defender for Cloud Apps" {
            Start-Process "https://portal.cloudappsecurity.com/"
            break
        }
        
    }
}
#endregion Function to open portal

#regiondefendercomboBox - Defender -Open Portal4
#
$defendercomboBox.AccessibleDescription = "Defender"
$defendercomboBox.AccessibleName = "Defender"
$defendercomboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$defendercomboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$defendercomboBox.FormattingEnabled = $true
$defendercomboBox.Items.AddRange(@(
"Microsoft 365 Defender",
"Incidents",
"Hunting",
"Action centre",
"Explorer",
"Restricted entities",
"Investigations",
"Threat Intelligence",
"Quarantined Emails",
"Microsoft Secure Score",
"Cloud app catalog",
"Manage OAuth apps",
"Device Inventory",
"Identities",
"Microsoft Defender for Cloud Apps"))
$defendercomboBox.Location = New-Object System.Drawing.Point(9, 142)
$defendercomboBox.Name = "defendercomboBox"
$defendercomboBox.Size = New-Object System.Drawing.Size(225, 25)
$defendercomboBox.TabIndex = 9
$defendercomboBox.Text = "Defender"
#endregion defendercomboBox

#region Add click action to EntraID ComboBox selection EntraIDcomboBox - Open Portal5
$EntraIDcomboBox.Add_SelectedIndexChanged({
    OpenPortal5
})
#endregion Click action to Combobox selection

#region Function to open portal EntraIDcomboBox  - Open Portal5
function OpenPortal5{
    $selectedPortal = $EntraIDcomboBox.SelectedItem.ToString()
    switch ($selectedPortal) {
        "Microsoft Entra" {
            Start-Process "https://entra.microsoft.com"
            break
        }
        "Microsoft Entra (Azure Portal)" {
            Start-Process "https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/Overview?"
            break
        }
        "Entra Admins Centre" {
            Start-Process "https://entra.microsoft.us/"
            break
        }
        "App registrations" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade/qu"
            break
        }
        "Enterprise applications" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/StartboardApplicationsMenuBlade/~/AppApps"
            break
        }
        "Groups" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupsManagementMenuBlade/~/AllG"
            break
        }
        "Users" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade"
            break
        }
        "Devices" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_Devices/DevicesMenuBlade/~/DeviceSettings/men"
            break
        }
        "Device settings" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_Devices/DevicesMenuBlade/~/DeviceSetting"
            break
        }
        "External Identities" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/CompanyRelationshipsMenuBlade/~/Settings/menuId"
            break
        }
        "Conditional Access" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/~/Polic"
            break
        }
        "Cross-tenant access settings" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/CompanyRelationshipsMenuBlade/~/"
            break
        }
        "ADFS services" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_Azure_ADHybridHealth/AadHealthMenuBlade/~/"
            break
        }
        "Authentication methods" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AuthenticationMethodsMenuBlade/~/Ad"
            break
        }
        "Consent and permissions" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ConsentPoliciesMenuBlade/~/UserSettings"
            break
        }
        "Legacy MFA" {
            Start-Process "https://account.activedirectory.windowsazure.com/UserManagement/MfaSettings.aspx"
            break
        }
        "Licenses" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/LicensesMenuBlade/~/Products"
            break
        }
        "Sign-in logs" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/SignInEventsV3Blade/timeRangeType/last24hou"
            break
        }
        "MFA unblock" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/MultifactorAuthenticationMenuBlade/~/BlockedUsers/fro"
            break
        }
        "PIM" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/CommonMenuBlade/~/quickStart"
            break
        }
        "PIM for Roles" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadmi"
            break
        }
        "PIM for Groups" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadg"
            break
        }
        "PIM for Azure resources" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_ERM/DashboardBlade/~/Controls"
            break
        }
        "Access reviews" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RolesManagementMenuBlade/~/AllRoles"
            break
        }
        "Roles and administrators" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RolesManagementMenuBlade/~/AllRoles"
            break
        }
        "Protected actions" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RolesManagementMenuBlade/~/Protecte"
            break
        }
        "Authentication context" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/~/A"
            break
        }
        "Secure score" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/SecurityMenuBlade/~/IdentitySecureScore/menu"
            break
        }
        "Security" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/SecurityMenuBlade/~/GettingStarted/menuI"
            break
        }
        "Password reset" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/PasswordResetMenuBlade/~/Properties"
            break
        }
        "Support" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_Azure_Support/NewSupportRequestV3Blade/caller"
            break
        }
        "AAD Connect - Sync errors" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_Azure_ADHybridHealth/AadHealthMenuBlade/~/S"
            break
        }
        "Azure B2B Invitations" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/UserCreateBlade/initialMode~/1/b2cEnabl"
            break
        }
        "Azure AD Entitlement Management" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_ERM/DashboardBlade/~/elmEntitlement"
            break
        }
        "Authentication strengths" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AuthenticationMethodsMenuBlade/~/A"
            break
        }
        "Authentication methods activity" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AuthenticationMethodsMenuBlade/~/AuthMe"
            break
        }
        "Login error lookup" {
            Start-Process "https://login.microsoftonline.com/error"
            break
        }
        "Login page" {
            Start-Process "https://microsoft.com/devicelogin"
            break
        }
        "Device login" {
            Start-Process "https://microsoft.com/devicelogin"
            break
        }
        "Login page" {
            Start-Process "https://portal.cloudappsecurity.com/"
            break
        }
        "Azure AD Identity Protection" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/IdentityProtectionMenuBlade/~/Overview"
            break
        }
        "LAPS" {
            Start-Process "https://entra.microsoft.com/#view/Microsoft_AAD_Devices/DevicesMenuBlade/~/LocalAdminPassword"
            break
        }
    }
}
#endregion Function to open portal

#regionentraidcomboBox - EntraID - Open Portal5
#
$entraidcomboBox.AccessibleDescription = "EntraID "
$entraidcomboBox.AccessibleName = "EntraID"
$entraidcomboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$entraidcomboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$entraidcomboBox.FormattingEnabled = $true
$entraidcomboBox.Items.AddRange(@(
"Microsoft Entra",
"Microsoft Entra (Azure Portal)",
"Entra Admins Centre",
"App registrations",
"Enterprise applications",
"Groups",
"Users",
"Devices",
"Device settings",
"External Identities",
"Conditional Access",
"Cross-tenant access settings",
"ADFS services",
"Authentication methods",
"Consent and permissions",
"Legacy MFA",
"Licenses",
"Sign-in logs",
"MFA unblock",
"PIM",
"PIM for Roles",
"PIM for Groups",
"PIM for Azure resources",
"Access reviews",
"Roles and administrators",
"Protected actions",
"Authentication context",
"Secure score",
"Security",
"Password reset",
"Support",
"AAD Connect - Sync errors",
"Azure B2B Invitations",
"Azure AD Entitlement Management",
"Authentication strengths",
"Authentication methods activity\t",
"Login error lookup",
"Login page",
"Device login",
"Azure AD Identity Protection",
"LAPS"))
$entraidcomboBox.Location = New-Object System.Drawing.Point(9, 182)
$entraidcomboBox.Name = "entraidcomboBox"
$entraidcomboBox.Size = New-Object System.Drawing.Size(225, 25)
$entraidcomboBox.TabIndex = 9
$entraidcomboBox.Text = "Entra ID"
#endregion entraid comboBox

#regionazurecomboBox - Azure
#
$azurecomboBox.AccessibleDescription = "Azure"
$azurecomboBox.AccessibleName = "Azure"
$azurecomboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$azurecomboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$azurecomboBox.FormattingEnabled = $true
$azurecomboBox.Items.AddRange(@(
"Azure",
"Azure Preview Portal",
"Sentinel",
"Azure Resource Graph Explorer",
"Purview Unified Data Governance",
"Automanage",
"Azure Virtual Desktop",
"Remote Desktop Web Client",
"Azure Automation",
"Function App",
"Azure Cosmos DB",
"CosmosDB Explorer",
"Monitor",
"Application Insights",
"Log Analytics",
"App Service plans",
"App Services",
"SQL servers",
"SQL databases",
"SQL elastic pools",
"SQL managed instance",
"Cost Management + Billing",
"Load balancing",
"Storage accounts",
"Disks",
"Virtual networks",
"Microsoft Defender for Cloud",
"Data factories",
"Help + support",
"Subscriptions",
"Resource groups",
"Network security groups",
"Key Vault",
"Data Explorer",
"Virtual machines",
"Kubernetes services",
"Azure Cache for Redis",
"Service Bus",
"Application Gateway",
"Web Application Firewall policies",
"Servers - Azure Arc",
"SSH Keys",
"Data Collection Rules",
"Container Registries",
"Azure Policy",
"Azure Cloud Shell",
"Azure Firewall Manager",
"Public IP Addresses",
"Azure Management Groups",
"Azure Update Management Center",
"Backup Center",
"Azure Synapse Studio",
"API Management services",
"Azure Front Door and CDN profiles",
"Azure Managed Identities",
"Azure Service Health",
"Azure - Create a Resource",
"Network Watcher",
"Azure Advisor",
"Availability Sets",
"Scale Sets",
"Azure Machine Learning Studio",
"DB for PostgreSQL servers",
"DB for PostgreSQL flexible servers",
"DB for PostgreSQL Hyperscale (Citus)",
"PostgreSQL servers - Azure Arc",
"Azure Databricks",
"Azure Synapse Analytics",
"Templates",
"Shared dashboards",
"Universal Print Console",
"Azure DevOps Portal",
"Azure Data Factory",
"Azure Price Calculator",
"Azure Logic Apps",
"Azure Resources PIM",
"Azure OpenAI Studio",
"Azure Content Safety",
"Azure Workbooks",
"Route tables"))
$azurecomboBox.Location = New-Object System.Drawing.Point(9, 101)
$azurecomboBox.Name = "azurecomboBox"
$azurecomboBox.Size = New-Object System.Drawing.Size(225, 25)
$azurecomboBox.TabIndex = 8
$azurecomboBox.Text = "Azure"
#endregion azure comboBox

#regionbutton7 - Exit
#
$button7.Dock = [System.Windows.Forms.DockStyle]::Bottom
$button7.FlatAppearance.BorderSize = 0
$button7.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button7.Font = New-Object System.Drawing.Font("Segoe UI", 12,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$button7.Location = New-Object System.Drawing.Point(0, 666)
$button7.Name = "button7"
$button7.Size = New-Object System.Drawing.Size(243, 41)
$button7.TabIndex = 7
$button7.Text = "Exit"
$button7.UseMnemonic = $false
$button7.UseVisualStyleBackColor = $false

function OnClick_button7 {$confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Exit Confirmation", "YesNo", "Question")
    if ($confirm -eq "Yes") {
        $form1.Close()
    }
}
$button7.Add_Click( { OnClick_button7 } )
#endregion button7

#regionbutton6 - Dashboard
#
$button6.Anchor = [System.Windows.Forms.AnchorStyles]"Top,Left"
$button6.FlatAppearance.BorderSize = 0
$button6.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button6.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
#$button6.ForeColor = [System.Drawing.Color]::FromArgb(255,128,0)
$button6.Location = New-Object System.Drawing.Point(73, 12)
$button6.Name = "button6"
$button6.Size = New-Object System.Drawing.Size(118, 31)
$button6.TabIndex = 6
$button6.Text = "Dashboard"
$button6.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$button6.UseVisualStyleBackColor = $true
#endregion button6

#regionbutton5 - button5
#
$button5.BackColor = [System.Drawing.Color]::FromArgb(44,62,80)
$button5.BackgroundImage = [System.Drawing.Image]::FromFile("C:\ccadmin\Modern Flat UI Design Dashboard\home.png")
$button5.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Zoom
$button5.FlatAppearance.BorderSize = 0
$button5.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button5.ForeColor = [System.Drawing.Color]::Transparent
$button5.Location = New-Object System.Drawing.Point(6, 6)
$button5.Name = "button5"
$button5.Size = New-Object System.Drawing.Size(51, 43)
$button5.TabIndex = 5
$button5.UseVisualStyleBackColor = $true
#endregion button5

#region panel3 - panel3
#
$panel3.Controls.Add($label4)
$panel3.Controls.Add($label2)
$panel3.Dock = [System.Windows.Forms.DockStyle]::Top
$panel3.Location = New-Object System.Drawing.Point(243, 58)
$panel3.Name = "panel3"
$panel3.Size = New-Object System.Drawing.Size(908, 43)
$panel3.TabIndex = 2
#endregion panel3

#regionlabel4 - Home / Dashboard
#
$label4.Anchor = [System.Windows.Forms.AnchorStyles]"Top,Right"
$label4.AutoSize = $true
$label4.Font = New-Object System.Drawing.Font("Segoe UI", 10,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label4.ForeColor = [System.Drawing.Color]::FromArgb(255,128,0)
$label4.Location = New-Object System.Drawing.Point(766, 12)
$label4.Name = "label4"
$label4.Size = New-Object System.Drawing.Size(137, 19)
$label4.TabIndex = 1
$label4.Text = "Home / Dashboard"
#endregion label4

#regionlabel2 - Dashboard
#
$label2.AutoSize = $true
$label2.Font = New-Object System.Drawing.Font("Segoe UI", 10,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label2.ForeColor = [System.Drawing.Color]::FromArgb(255,128,0)
$label2.Location = New-Object System.Drawing.Point(18, 12)
$label2.Name = "label2"
$label2.Size = New-Object System.Drawing.Size(82, 19)
$label2.TabIndex = 0
$label2.Text = "Dashboard"
#endregion label2

#regiontabControl1
#
$tabControl1.Controls.Add($tabPage1)
$tabControl1.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl1.Location = New-Object System.Drawing.Point(243, 101)
$tabControl1.Name = "tabControl1"
$tabControl1.SelectedIndex = 0
$tabControl1.Size = New-Object System.Drawing.Size(908, 664)
$tabControl1.TabIndex = 3
#endregion tabcontrol1

#regiontabPage1 Azure Global Health Status

$tabPage1.Controls.Add($webBrowser1)
$tabPage1.Location = New-Object System.Drawing.Point(4, 22)
$tabPage1.Name = "tabPage1"
$tabPage1.Padding = New-Object System.Windows.Forms.Padding(3)
$tabPage1.Size = New-Object System.Drawing.Size(900, 638)
$tabPage1.TabIndex = 0
$tabPage1.Text = "Azure Global Health Status"
$tabPage1.UseVisualStyleBackColor = $true
#endregion tabpage1

#regionwebBrowser1 webBrowser1
#
$webBrowser1.Dock = [System.Windows.Forms.DockStyle]::Fill
$webBrowser1.Location = New-Object System.Drawing.Point(3, 3)
$webBrowser1.MinimumSize = New-Object System.Drawing.Size(20, 20)
$webBrowser1.Name = "webBrowser1"
$webBrowser1.ScriptErrorsSuppressed = $true
$webBrowser1.Size = New-Object System.Drawing.Size(894, 632)
$webBrowser1.TabIndex = 0
$webBrowser1.Url = New-Object System.Uri("https://azure.status.microsoft/en-us/status")
#endregion Browser1

#regionlabel6 Office365 Troubleshooting
#
$label6.AutoSize = $true
$label6.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label6.ForeColor = [System.Drawing.Color]::FromArgb(255,128,0)
$label6.Location = New-Object System.Drawing.Point(6, 358)
$label6.Name = "label6"
$label6.Size = New-Object System.Drawing.Size(172, 17)
$label6.TabIndex = 14
$label6.Text = "Office365 Troubleshooting"
#endregion label6

#regionAdd click action to ComboBox selection Office Troubleshooting - Open Portal6
$comboBox1.Add_SelectedIndexChanged({
    OpenPortal6
})
#endregion Click action to Combobox1 selection

#regionFunction to open portal Office Troubleshooting - Open Portal6
function OpenPortal6{
    $selectedPortal = $comboBox1.SelectedItem.ToString()
    switch ($selectedPortal) {
        "DNSSEC and DANE Validation Test" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365DaneValidation/input"
            break
        }
        "Exchange Online Custom Domains DNS Connectivity Test" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365ExchangeDns/input"
            break
        }
        "Teams DNS Connectivity Test" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365TeamsDns/input"
            break
        }
        "Exchange ActiveSync" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365Eas/input"
            break
        }
        "Synchronization, Notification, Availability, and Automatic Replies" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365EwsTask/input"
            break
        }
        "Service Account Access (Developers)" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365EwsAccess/input"
            break
        }
        "Outlook Connectivity" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365Ola/input"
            break
        }
        "Inbound SMTP Email" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365InboundSmtp/input"
            break
        }
        "Outbound SMTP Email" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365OutboundSmtp/input"
            break
        }
        "Free/Busy" {
            Start-Process "https://testconnectivity.microsoft.com/tests/FreeBusy/input"
            break
        }
        "Outlook Mobile Hybrid Modern Authentication Test" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365OlkMobHma/input"
            break
        }
        "Mailbox Provisioning Test" {
            Start-Process "https://testconnectivity.microsoft.com/tests/O365ProvisioningUserMaibox/input"
            break
        }
        "Message Header Analyzer" {
            Start-Process "https://mha.azurewebsites.net/"
            break
        }
        "Exchange ActiveSync" {
            Start-Process "https://testconnectivity.microsoft.com/tests/Eas/input"
            break
        }
        "Synchronization, Notification, Availability, and Automatic Replies" {
            Start-Process "https://testconnectivity.microsoft.com/tests/EwsTask/input"
            break
        }
        "SSL Server Test" {
            Start-Process "https://testconnectivity.microsoft.com/tests/SslServer/input"
            break
        }
        "Teams Sign in" {
            Start-Process "https://testconnectivity.microsoft.com/tests/TeamsSignin/input"
            break
        }
        "Teams Calendar Tab" {
            Start-Process "https://testconnectivity.microsoft.com/tests/TeamsCalendarMissing/input"
            break
        }
        "Teams Meeting Delegation" {
            Start-Process "https://testconnectivity.microsoft.com/tests/TeamsMeetingDelegation/input"
            break
        }
        "Teams Channel Meeting" {
            Start-Process "https://testconnectivity.microsoft.com/tests/TeamsChannelMeeting/input"
            break
        }
        "Microsoft Teams Room Sign in" {
            Start-Process "https://testconnectivity.microsoft.com/tests/TeamsMTRDeviceSignIn/input"
            break
        }
    }
}
#endregion Function to open portal

#regioncomboBox1 Office Troubleshooting - Open Portal6
#
$comboBox1.AccessibleDescription = "Office Troubleshooting"
$comboBox1.AccessibleName = "Troubleshooting"
$comboBox1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$comboBox1.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$comboBox1.FormattingEnabled = $true
$comboBox1.Items.AddRange(@(
"DNSSEC and DANE Validation Test",
"Exchange Online Custom Domains DNS Connectivity Test",
"Teams DNS Connectivity Test",
"Exchange ActiveSync",
"Synchronization, Notification, Availability, and Automatic Replies",
"Service Account Access (Developers)",
"Outlook Connectivity",
"Inbound SMTP Email",
"Outbound SMTP Email",
"Free/Busy",
"Outlook Mobile Hybrid Modern Authentication Test",
"Mailbox Provisioning Test",
"Message Header Analyzer",
"Exchange ActiveSync",
"Synchronization, Notification, Availability, and Automatic Replies",
"SSL Server Test",
"Teams Sign in",
"Teams Calendar Tab",
"Teams Meeting Delegation",
"Teams Channel Meeting",
"Microsoft Teams Room Sign in"
))
$comboBox1.Location = New-Object System.Drawing.Point(8, 383)
$comboBox1.Name = "comboBox1"
$comboBox1.Size = New-Object System.Drawing.Size(226, 25)
$comboBox1.TabIndex = 15
$comboBox1.Text = "Office Troubleshooting"
#endregion comboBox1

#regionlabel7 - Support and Recovery Assistant
#
$label7.AutoSize = $true
$label7.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label7.ForeColor = [System.Drawing.Color]::FromArgb(255,128,0)
$label7.Location = New-Object System.Drawing.Point(6, 427)
$label7.Name = "label7"
$label7.Size = New-Object System.Drawing.Size(204, 17)
$label7.TabIndex = 16
$label7.Text = "Support and Recovery Assistant"
#endregion label7

#regionAdd click action to ComboBox2 selection - Support and Recovery Assistant - Open Portal7
$comboBox2.Add_SelectedIndexChanged({
    OpenPortal7
})
#endregion Click action to Combobox2 selection

#regionFunction to open portal - Support and Recovery Assistant - Open Portal7
function OpenPortal7{
    $selectedPortal = $comboBox2.SelectedItem.ToString()
    switch ($selectedPortal) {
        "Fix my Excel startup issue" {
            Start-Process "https://aka.ms/SaRA-ExcelCrash-sarahome"
            break
        }
        "Let's uninstall Office" {
            Start-Process "https://aka.ms/SaRA-OfficeUninstall-sarahome"
            break
        }
        "Let's setup your Office" {
            Start-Process "https://aka.ms/SaRA-OfficeSetup-sarahome"
            break
        }
        "Let's Activate your Office" {
            Start-Process "https://aka.ms/SaRA-OfficeActivation-sarahome"
            break
        }
        "Resolve Office Sign in issues" {
            Start-Process "https://aka.ms/SaRA-OfficeSignIn-sarahome"
            break
        }
        "Enable Office Shared Computer Activation" {
            Start-Process "https://aka.ms/SaRA-OfficeSCA-sarahome"
            break
        }
        "Scan Office - ROI Normal Scan" {
            Start-Process "https://aka.ms/SaRA-RoiscanNormal-sarahome"
            break
        }
        "Scan Office - ROI Full Scan" {
            Start-Process "https://aka.ms/SaRA-RoiscanFull-sarahome"
            break
        }
        "Let's setup your Outlook profile" {
            Start-Process "https://aka.ms/SaRA-OutlookSetupProfile-sarahome"
            break
        }
        "Troubleshoot Outlook won't start issue" {
            Start-Process "https://aka.ms/SaRA-olkstart-sarahome"
            break
        }
        "Resolve Outlook password issues" {
            Start-Process "https://aka.ms/SaRA-OutlookPwdPrompt-sarahome"
            break
        }
        "Resolve Outlook connection issues" {
            Start-Process "https://aka.ms/SaRA-OutlookDisconnect-sarahome"
            break
        }
        "Create detailed Outlook diagnostic report" {
            Start-Process "https://aka.ms/SaRA-OutlookAdvDiagExpExp-sarahome"
            break
        }
        "Find Outlook calendar issues" {
            Start-Process "https://aka.ms/SaRA-CalCheck-sarahome"
            break
        }
        "Fix Teams Meeting add-in issues" {
            Start-Process "https://aka.ms/SaRA-TeamsAddin-sarahome"
            break
        }
        "Fix Teams user presence information" {
            Start-Process "https://aka.ms/SaRA-TeamsPresence-sarahome"
            break
        }
        "Resolve Team Sign Issues" {
            Start-Process "https://aka.ms/SaRA-TeamsSignIn-sarahome"
            break
        }
        "Outlook Advanced diagnostic solution" {
            Start-Process "https://aka.ms/SaRA-OutlookAdvDiagExpExp-sarahome"
            break
        }
        "Outlook Calendar check solution" {
            Start-Process "https://aka.ms/SaRA-CalCheck-sarahome"
            break
        }
        "Network Connectivity Test" {
            Start-Process "https://aka.ms/SaRA-NetworkConnectivity-sarahome"
            break
        }
    }
}
#endregion Function to open portal

#regioncomboBox2 Support and Recovery Assistant - Open Portal7
#
$comboBox2.AccessibleDescription = "Support and Recovery Assistant"
$comboBox2.AccessibleName = "Support and Recovery Assistant"
$comboBox2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$comboBox2.Font = New-Object System.Drawing.Font("Segoe UI", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$comboBox2.FormattingEnabled = $true
$comboBox2.Items.AddRange(@(
"Fix my Excel startup issue",
"Let's uninstall Office",
"Let's setup your Office",
"Let's Activate your Office",
"Resolve Office Sign in issues",
"Enable Office Shared Computer Activation",
"Scan Office - ROI Normal Scan",
"Scan Office - ROI Full Scan",
"Let's setup your Outlook profile",
"Troubleshoot Outlook won\'t start issue",
"Resolve Outlook password issues",
"Resolve Outlook connection issues",
"Create detailed Outlook diagnostic report",
"Find Outlook calendar issues",
"Fix Teams Meeting add-in issues",
"Fix Teams user presence information",
"Resolve Team Sign Issues",
"Outlook Advanced diagnostic solution",
"Outlook Calendar check solution ",
"Network Connectivity Test"))
$comboBox2.Location = New-Object System.Drawing.Point(8, 452)
$comboBox2.Name = "comboBox2"
$comboBox2.Size = New-Object System.Drawing.Size(226, 25)
$comboBox2.TabIndex = 17
$comboBox2.Text = "Support and Recovery Assistant"
#endregion comboBox2

#regionmonthCalendar1
#
$monthCalendar1.Anchor = [System.Windows.Forms.AnchorStyles]"Top,Bottom,Left,Right"
$monthCalendar1.FirstDayOfWeek = [System.Windows.Forms.Day]::Monday
$monthCalendar1.Location = New-Object System.Drawing.Point(6, 504)
$monthCalendar1.Margin = New-Object System.Windows.Forms.Padding(0)
$monthCalendar1.MinimumSize = New-Object System.Drawing.Size(218, 162)
$monthCalendar1.Name = "monthCalendar1"
$monthCalendar1.TabIndex = 18

function OnDateChanged_monthCalendar1 {
	[void][System.Windows.Forms.MessageBox]::Show("The event handler monthCalendar1.Add_DateChanged is not implemented.")
}

$monthCalendar1.Add_DateChanged( { OnDateChanged_monthCalendar1 } )
#endregion calendar1

#regionForm1
#
#$Form1.WindowState = "Maximized"
$Form1.AutoSize = $true
$Form1.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Zoom
$Form1.ClientSize = New-Object System.Drawing.Size(1151, 765)
$Form1.Controls.Add($tabControl1)
$Form1.Controls.Add($panel3)
$Form1.Controls.Add($panel2)
$Form1.Controls.Add($panel1)
$Form1.ForeColor = [System.Drawing.Color]::WhiteSmoke
$Form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$Form1.Name = "Form1"
$Form1.ShowIcon = $false
$Form1.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Show
$Form1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Form1.Text = "Consolidated Portal Links"
$Form1.ShowDialog() | Out-Null
#endregion form1
#

