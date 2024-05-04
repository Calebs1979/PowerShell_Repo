# Loading external assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# Enable Visual Styles
[Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::EnableVisualStyles()

# Get the current logged-on user
$currentUserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name 
#
#
# Extract only the username from the full user name
$currentUserName = $currentUserName -replace ".*\\"  
#
#
# Capitalize the first letter of the username
$capitalizedUserName = $currentUserName.Substring(0,1).ToUpper() + $currentUserName.Substring(1)

# Create the message
$message = "Hello, $capitalizedUserName!`n`nPlease ensure that you close the Reporting Tool when you have completed.`n`nThank you $capitalizedUserName" 
                                                                                                                                                                
# Display the popup window                                                                                                                                       
Add-Type -AssemblyName PresentationFramework                                                                                                                     
[System.Windows.MessageBox]::Show($message, "Informational Message", "OK", "Information")
#
#
# Create transcript log folder if not exist
$folderPath = "C:\RTAdmin\Reports Toolset Files\Transcript logs"
if (-not (Test-Path -Path $folderPath)) {  
    New-Item -Path $folderPath -ItemType Directory
}
#
# Create Reports log folder if not exist
$folderPath = "C:\RTAdmin\Reports Toolset Files\Reports"
if (-not (Test-Path -Path $folderPath)) {
    New-Item -Path $folderPath -ItemType Directory
}
#
# Start transcript log  file and write to path C:\RTAdmin\
Start-Transcript -Path "C:\RTAdmin\Reports Toolset Files\Transcript logs\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Log1.txt"

# Install needed modules
Install-Module -Name Msonline

# Check if Active Directory module is installed #######
$adModule = Get-Module -Name ActiveDirectory -ListAvailable
#
# If not installed, install Active Directory module
if (-not $adModule) {
    Write-Host "Installing Active Directory module..."
    Install-WindowsFeature RSAT-AD-PowerShell

    #Import Active Directory module
    Import-Module ActiveDirectory
}

# Creation of the components.
$Form1 = New-Object System.Windows.Forms.Form

$BottomToolStripPanel = New-Object System.Windows.Forms.ToolStripPanel
$TopToolStripPanel = New-Object System.Windows.Forms.ToolStripPanel
$RightToolStripPanel = New-Object System.Windows.Forms.ToolStripPanel
$LeftToolStripPanel = New-Object System.Windows.Forms.ToolStripPanel
$ContentPanel = New-Object System.Windows.Forms.ToolStripContentPanel
$menuStrip1 = New-Object System.Windows.Forms.MenuStrip
$fileToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$computerManagementToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$computerManagementToolStripMenuItem1 = New-Object System.Windows.Forms.ToolStripMenuItem
$diskManagmentToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$eventViewerToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$taskSchedulerToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$regeditToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$certificateManagerToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$localUsersAndComputersToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$peformanceMonitorToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$comandPromptToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$powerShellToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$powerShellISEToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$getHotfixToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$checksmbstatusToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$azureToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllSubscriptionsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$getMFAToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$azurelinksToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$IntunelinksToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$office365linksToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$defenderlinksToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listallasrvaultsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listentraidguestusersToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listallresourcesinsubscriptionToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listvirtualmachinesinsubscriptionToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$intuneToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllWindows10DevicesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllIOSDevicesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllAndroidDevicesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listWin10NonCompliantDevicesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listIOSNonCompliantDevicesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAndroidNonCompliantDevicesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllIntuneApplicationsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$forceSyncAllDevicesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$enrollmentStatusToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$office365ToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$licensesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllOffice365LicensesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listUnlicensedUsersToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$mailboxReportsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$office365troubleshootingToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$office365connectivitytestsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$ExchangeconnectivitytestsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$microsoftteamsconnectivitytestsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$microsoft365networkconnectivitytestToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$microsoft365networkhealthstatusToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$microsoftmessageheaderanalysisToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$supporttoolsStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$outlookadvanceddiagnosticsStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$outlookcalendardiagnosticsStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$normalscanofficeinstallationStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$fullscanofficeinstallationStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$networkconnectivitytestStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllSharedMailboxesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllMailboxSizeToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listInactiveMailboxesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listDeletedMailboxesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$dynamicDistributionGroupsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$distributionGroupsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$activeDirectoryToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$generatePasswordToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$backupallgrouppoliciesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$unlockandenableuseraccountToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$lockedOutUsersToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$disabledUsersToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$allUsersLogonHistoryToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$searchGroupPolicyToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$domainDefaultPasswordPolicyToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllSecurityGroupsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllActiveDirectoryUsersToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllDomainControllersToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$listAllDomainSSLCertificatesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$checkActiveDirectoryReplicationToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$dNSToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$ExportdnsZonesToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$dHCPToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$toolStripSeparator1 = New-Object System.Windows.Forms.ToolStripSeparator
$toolStripSeparator2 = New-Object System.Windows.Forms.ToolStripSeparator
$toolStripSeparator3 = New-Object System.Windows.Forms.ToolStripSeparator
$reportsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$usersNotLoggedOnIn30DaysToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$usersNotLoggedOnIn60DaysToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$usersNotLoggedOnIn90DaysToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$usersCreatedInPast30daysToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$linksToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$azureloginportalToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$intuneloginportalToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$office365loginportalToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$azureglobalhealthstatusToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$sentinelloginportalToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$365defenderportalloginToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$azureresourcegraphexplorerToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$graphexplorerToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$purviewCompliancemanagerToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$purvieweDiscoveryToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$advancedhuntingToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$authenticationmethodsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$conditionalaccessToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aadconnectsyncerrorsToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$azuresupportportalToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$infoToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$contactToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$supportToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$monthCalendar1 = New-Object System.Windows.Forms.MonthCalendar
$toolStripContainer1 = New-Object System.Windows.Forms.ToolStripContainer
$richTextBox1 = New-Object System.Windows.Forms.RichTextBox
$exportButton = New-Object System.Windows.Forms.Button
$clearButton = New-Object System.Windows.Forms.Button
$pwdclearButton = New-Object System.Windows.Forms.Button
$copyButton = New-Object System.Windows.Forms.Button
$copyButton1 = New-Object System.Windows.Forms.Button
$button_formexit = New-Object System.Windows.Forms.Button
$pictureBox1 = New-Object System.Windows.Forms.PictureBox
#
# BottomToolStripPanel
#
$BottomToolStripPanel.Location = New-Object System.Drawing.Point(0, 0)
$BottomToolStripPanel.BackColor = [System.Drawing.Color]::AliceBlue
$BottomToolStripPanel.Name = "BottomToolStripPanel"
$BottomToolStripPanel.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$BottomToolStripPanel.RowMargin = New-Object System.Windows.Forms.Padding(3, 0, 0, 0)
$BottomToolStripPanel.Size = New-Object System.Drawing.Size(0, 0)
#
# TopToolStripPanel
#
$TopToolStripPanel.Location = New-Object System.Drawing.Point(0, 0)
$TopToolStripPanel.BackColor = [System.Drawing.Color]::AliceBlue
$TopToolStripPanel.Name = "TopToolStripPanel"
$TopToolStripPanel.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$TopToolStripPanel.RowMargin = New-Object System.Windows.Forms.Padding(3, 0, 0, 0)
$TopToolStripPanel.Size = New-Object System.Drawing.Size(0, 0)
#
# RightToolStripPanel
#
$RightToolStripPanel.Location = New-Object System.Drawing.Point(0, 0)
$RightToolStripPanel.BackColor = [System.Drawing.Color]::AliceBlue
$RightToolStripPanel.Name = "RightToolStripPanel"
$RightToolStripPanel.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$RightToolStripPanel.RowMargin = New-Object System.Windows.Forms.Padding(3, 0, 0, 0)
$RightToolStripPanel.Size = New-Object System.Drawing.Size(0, 0)
#
# LeftToolStripPanel
#
$LeftToolStripPanel.Location = New-Object System.Drawing.Point(0, 0)
$LeftToolStripPanel.BackColor = [System.Drawing.Color]::AliceBlue
$LeftToolStripPanel.Name = "LeftToolStripPanel"
$LeftToolStripPanel.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$LeftToolStripPanel.RowMargin = New-Object System.Windows.Forms.Padding(3, 0, 0, 0)
$LeftToolStripPanel.Size = New-Object System.Drawing.Size(0, 0)
#
# ContentPanel
#
$ContentPanel.BackColor = [System.Drawing.Color]::AliceBlue
$ContentPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$ContentPanel.Size = New-Object System.Drawing.Size(1106, 146)
#
# menuStrip1
#
$menuStrip1.AllowDrop = $true
$menuStrip1.BackColor = [System.Drawing.Color]::AliceBlue
$menuStrip1.Items.AddRange(@(
$fileToolStripMenuItem,
$computerManagementToolStripMenuItem,
$azureToolStripMenuItem,
$intuneToolStripMenuItem,
$office365ToolStripMenuItem,
$activeDirectoryToolStripMenuItem,
$LinksToolStripMenuItem,
$aboutToolStripMenuItem))
$menuStrip1.Location = New-Object System.Drawing.Point(0, 0)
$menuStrip1.Name = "menuStrip1"
$menuStrip1.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::Professional
$menuStrip1.Size = New-Object System.Drawing.Size(856, 24)
$menuStrip1.TabIndex = 0
$menuStrip1.Text = "menuStrip1"
#
# fileToolStripMenuItem
#
$fileToolStripMenuItem.DropDownItems.AddRange(@(
$exitToolStripMenuItem))
$fileToolStripMenuItem.Name = "fileToolStripMenuItem"
$fileToolStripMenuItem.Size = New-Object System.Drawing.Size(37, 20)
$fileToolStripMenuItem.Text = "File"
#
# exitToolStripMenuItem
#
$exitToolStripMenuItem.Name = "exitToolStripMenuItem"
$exitToolStripMenuItem.Size = New-Object System.Drawing.Size(180, 22)
$exitToolStripMenuItem.Text = "Exit"
$exitToolStripMenuItem.Add_Click({$confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Exit Confirmation", "YesNo", "Question")
    if ($confirm -eq "Yes") {
        $form1.Close()
    }
})
#
# computerManagementToolStripMenuItem
#
$computerManagementToolStripMenuItem.DropDownItems.AddRange(@(
$certificateManagerToolStripMenuItem,
$checksmbstatusToolStripMenuItem,
$comandPromptToolStripMenuItem,
$computerManagementToolStripMenuItem1,
$diskManagmentToolStripMenuItem,
$eventViewerToolStripMenuItem,
$getHotfixToolStripMenuItem
$localUsersAndComputersToolStripMenuItem,
$peformanceMonitorToolStripMenuItem,
$powerShellToolStripMenuItem,
$powerShellISEToolStripMenuItem,
$regeditToolStripMenuItem,
$taskSchedulerToolStripMenuItem))
$computerManagementToolStripMenuItem.Name = "computerManagementToolStripMenuItem"
$computerManagementToolStripMenuItem.Size = New-Object System.Drawing.Size(147, 20)
$computerManagementToolStripMenuItem.Text = "Computer Management"
$computerManagementToolStripMenuItem.BackColor = [System.Drawing.Color]::AliceBlue
#
# computerManagementToolStripMenuItem1
#
$computerManagementToolStripMenuItem1.Name = "computerManagementToolStripMenuItem1"
$computerManagementToolStripMenuItem1.Size = New-Object System.Drawing.Size(218, 22)
$computerManagementToolStripMenuItem1.Text = "Computer Management"
$computerManagementToolStripMenuItem1.Add_Click({
    # Open Computer Management
    Start-Process "compmgmt.msc"
})
#
# diskManagmentToolStripMenuItem
#
$diskManagmentToolStripMenuItem.Name = "diskManagmentToolStripMenuItem"
$diskManagmentToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$diskManagmentToolStripMenuItem.Text = "Disk Managment"
$diskManagmentToolStripMenuItem.Add_Click({
    #Open Disk Management
    Start-Process diskmgmt.msc
})
#
# eventViewerToolStripMenuItem
#
$eventViewerToolStripMenuItem.Name = "eventViewerToolStripMenuItem"
$eventViewerToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$eventViewerToolStripMenuItem.Text = "Event Viewer"
$eventViewerToolStripMenuItem.Add_Click({
    # Open Event Viewer
    Start-Process -FilePath "eventvwr.exe"
})
#
# taskSchedulerToolStripMenuItem
#
$taskSchedulerToolStripMenuItem.Name = "taskSchedulerToolStripMenuItem"
$taskSchedulerToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$taskSchedulerToolStripMenuItem.Text = "Task Scheduler"
$taskSchedulerToolStripMenuItem.Add_Click({
    # Open Task Scheduler
    Start-Process "taskschd.msc"
})
#
# regeditToolStripMenuItem
#
$regeditToolStripMenuItem.Name = "regeditToolStripMenuItem"
$regeditToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$regeditToolStripMenuItem.Text = "Regedit"
$regeditToolStripMenuItem.Add_Click({
    # Open Task Scheduler
    Start-Process "regedit.exe"
})
#
# certificateManagerToolStripMenuItem
#
$certificateManagerToolStripMenuItem.Name = "certificateManagerToolStripMenuItem"
$certificateManagerToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$certificateManagerToolStripMenuItem.Text = "Certificate Manager"
$certificateManagerToolStripMenuItem.Add_Click({
    # Open Task Scheduler
    Start-Process "certmgr.msc"
})
#
# localUsersAndComputersToolStripMenuItem
#
$localUsersAndComputersToolStripMenuItem.Name = "localUsersAndComputersToolStripMenuItem"
$localUsersAndComputersToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$localUsersAndComputersToolStripMenuItem.Text = "Local Users and Computers"
$localUsersAndComputersToolStripMenuItem.Add_Click({
    # Open Task Scheduler
    Start-Process -FilePath "lusrmgr.msc"
})
#
# peformanceMonitorToolStripMenuItem
#
$peformanceMonitorToolStripMenuItem.Name = "peformanceMonitorToolStripMenuItem"
$peformanceMonitorToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$peformanceMonitorToolStripMenuItem.Text = "Peformance Monitor"
$peformanceMonitorToolStripMenuItem.Add_Click({
    # Start Performance Monitor
    Start-Process -FilePath "perfmon.exe"
})
#
# comandPromptToolStripMenuItem
#
$comandPromptToolStripMenuItem.Name = "comandPromptToolStripMenuItem"
$comandPromptToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$comandPromptToolStripMenuItem.Text = "Comand Prompt"
$comandPromptToolStripMenuItem.Add_Click({
    # Start Performance Monitor
    Start-Process -FilePath "cmd.exe"
})
#
# powerShellToolStripMenuItem
#
$powerShellToolStripMenuItem.Name = "powerShellToolStripMenuItem"
$powerShellToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$powerShellToolStripMenuItem.Text = "PowerShell"
$powerShellToolStripMenuItem.Add_Click({
    # Start Performance Monitor
    Start-Process -FilePath "powershell.exe"
})
#
# powerShellISEToolStripMenuItem
#
$powerShellISEToolStripMenuItem.Name = "powerShellISEToolStripMenuItem"
$powerShellISEToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$powerShellISEToolStripMenuItem.Text = "PowerShell ISE"
$powerShellISEToolStripMenuItem.Add_Click({
    # Start Performance Monitor
    Start-Process -FilePath "powershell_ise.exe"
})
#
# getHotfixToolStripMenuItem
#
$getHotfixToolStripMenuItem.Name = "getHotfixToolStripMenuItem"
$getHotfixToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$getHotfixToolStripMenuItem.Text = "Get Hotfix"
$getHotfixToolStripMenuItem.Add_Click({
    $output =  Get-HotFix | Format-Table -AutoSize | Out-String
    $richtextbox1[0].AppendText($output)
})
#
# checksmbstatusToolStripMenuItem
#
$checksmbstatusToolStripMenuItem.Name = "checksmbstatusToolStripMenuItem"
$checksmbstatusToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 22)
$checksmbstatusToolStripMenuItem.Text = "Check SMB Protocol Status"
$checksmbstatusToolStripMenuItem.Add_Click({
    $output =  Get-SmbServerConfiguration | Select-Object EnableSMB1Protocol, EnableSMB2Protocol, EnableSMB3Protocol | Format-Table -AutoSize | Out-String
    $richtextbox1[0].AppendText($output)
})
#
#
# azureToolStripMenuItem
#
$azureToolStripMenuItem.DropDownItems.AddRange(@(
$getMFAToolStripMenuItem,
$listallSubscriptionsToolStripMenuItem,
$listallasrvaultsToolStripMenuItem,
$listallresourcesinsubscriptionToolStripMenuItem,
$listentraidguestusersToolStripMenuItem,
$listvirtualmachinesinsubscriptionToolStripMenuItem))
$azureToolStripMenuItem.Name = "azureToolStripMenuItem"
$azureToolStripMenuItem.Size = New-Object System.Drawing.Size(49, 20)
$azureToolStripMenuItem.Text = "Azure"
#
# listAllSubscriptionsToolStripMenuItem
#
$listAllSubscriptionsToolStripMenuItem.Name = "listAllSubscriptionsToolStripMenuItem"
$listAllSubscriptionsToolStripMenuItem.Size = New-Object System.Drawing.Size(183, 22)
$listAllSubscriptionsToolStripMenuItem.Text = "List All Subscriptions"
$listAllSubscriptionsToolStripMenuItem.Add_Click({ Connect-AzAccount
           $Output = Get-AzSubscription | Format-Table -AutoSize | Out-String
           $richtextbox1[0].AppendText($output)
           
})
#
# getMFAToolStripMenuItem
#
$getMFAToolStripMenuItem.Name = "getMFAToolStripMenuItem"
$getMFAToolStripMenuItem.Size = New-Object System.Drawing.Size(183, 22)
$getMFAToolStripMenuItem.Text = "Get MFA Status"
$getMFAToolStripMenuItem.Add_Click({ Connect-AzAccount
$users = Get-ADUser -Filter 'enabled -eq $true'

# Initialize array to store MFA status
$mfaStatus = @()

# Loop through each user and get MFA status
foreach ($user in $users) {
    $mfaStatus += [PSCustomObject]@{
        DisplayName = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        MFAStatus = (Get-AzADUser -ObjectId $user.ObjectId).StrongAuthenticationMethods.Count -gt 0
    }
}
# Display MFA status
$mfaStatus | Out-GridView -Title "MFS Status of Azure AD Users" -Passthru
if ($mfaStatus) {
            $mfaStatus | Select DisplayName,UserPrincipalName | Export-CSV "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_AzureAD Users MFA Status.csv" -NoTypeInformation -Delimiter ";"
            }
 # Disconnect from Azure account
Disconnect-AzAccount
})
#
# getMFAToolStripMenuItem
#
$listallasrvaultsToolStripMenuItem.Name = "getMFAToolStripMenuItem"
$listallasrvaultsToolStripMenuItem.Size = New-Object System.Drawing.Size(183, 22)
$listallasrvaultsToolStripMenuItem.Text = "List All ASR Vaults"
$listallasrvaultsToolStripMenuItem.Add_Click({Connect-AzAccount
$subscriptions = Get-AzSubscription
        # loop through the subscriptions and collect the output in variable $result
            $result = foreach ($Subscription in $Subscriptions) {
            $null   = Set-AzContext -SubscriptionId $Subscription.Id
            $vaults = Get-AzRecoveryServicesVault

        foreach ($vault in $vaults) {
        Set-AzRecoveryServicesVaultContext -Vault $vault
        # output a PSObject with the chosen properties
        Get-AzRecoveryServicesVault | Select-Object "Name","ResourceGroupName","Location","SubscriptionId" 
           }
        } 
          $result | Out-GridView -Title "$selectedItem" -PassThru | Export-Csv "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_ListConditionalAccessPoliciesList.csv" -NoTypeInformation -Encoding ASCII
})  
#
$listentraidguestusersToolStripMenuItem.Name = "getMFAToolStripMenuItem"
$listentraidguestusersToolStripMenuItem.Size = New-Object System.Drawing.Size(183, 22)
$listentraidguestusersToolStripMenuItem.Text = "List EntraID Guest Users"
$listentraidguestusersToolStripMenuItem.Add_Click({ Connect-AzureAD
$Output = Get-AzureADUser -Filter "UserType eq 'Guest'" | Select DisplayName,UserPrincipalName, UserType, ObjectID
         $textBox2.Text = $Output | Out-GridView -Title "List EntraID Guest Users" -PassThru
        if ($Output) {
            $Output | Select DisplayName,UserPrincipalName | Export-CSV "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List EntraID Guest Users.csv" -NoTypeInformation -Delimiter ";"
            }
 })
#
$listallresourcesinsubscriptionToolStripMenuItem.Name = "getMFAToolStripMenuItem"
$listallresourcesinsubscriptionToolStripMenuItem.Size = New-Object System.Drawing.Size(183, 22)
$listallresourcesinsubscriptionToolStripMenuItem.Text = "List All Resources in Subscription"
$listallresourcesinsubscriptionToolStripMenuItem.Add_Click({ Connect-AzAccount

# Get all Azure subscriptions
$subscriptions = Get-AzSubscription | Select-Object -Property Name, SubscriptionId

# Display subscriptions in a GridView and allow selection
$selectedSubscription = $subscriptions | Out-GridView -Title "Select Azure Subscription" -PassThru

if ($selectedSubscription) {
    # Set the selected subscription context
    Set-AzContext -SubscriptionId $selectedSubscription.SubscriptionId

    # Get resources of the selected subscription
    $resources = Get-AzResource | Out-GridView -PassThru

    # Output resources to CSV
    $csvPath = "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_AllResources.csv"
    $resources | Export-Csv -Path $csvPath -NoTypeInformation

    # Display confirmation message
    Write-Host "$selectedSubscription Resources exported to $csvPath" -ForegroundColor Green
} else {
    Write-Host "No subscription selected. Exiting script." -ForegroundColor Yellow
}
})
#
$listvirtualmachinesinsubscriptionToolStripMenuItem.Name = "getMFAToolStripMenuItem"
$listvirtualmachinesinsubscriptionToolStripMenuItem.Size = New-Object System.Drawing.Size(183, 22)
$listvirtualmachinesinsubscriptionToolStripMenuItem.Text = "List All Virtual Machines in Subscription"
$listvirtualmachinesinsubscriptionToolStripMenuItem.Add_Click({ Connect-AzAccount

# Get all Azure subscriptions
$subscriptions = Get-AzSubscription | Select-Object -Property Name, SubscriptionId

# Display subscriptions in a GridView and allow selection
$selectedSubscription = $subscriptions | Out-GridView -Title "Select Azure Subscription" -PassThru

if ($selectedSubscription) {
    # Set the selected subscription context
    Set-AzContext -SubscriptionId $selectedSubscription.SubscriptionId

    # Get resources of the selected subscription
    $resources = Get-AzVM | Select-Object -Property Name, ResourceGroup, Location, VmSize, OSType | Out-GridView -PassThru

    # Output resources to CSV
    $csvPath = "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_VirtualMachines.csv"
    $resources | Export-Csv -Path $csvPath -NoTypeInformation

    # Display confirmation message
    Write-Host "$selectedSubscription Virtual Machines have been exported to $csvPath" -ForegroundColor Green
} else {
    Write-Host "No subscription selected. Exiting script." -ForegroundColor Yellow
}
})
#
# intuneToolStripMenuItem
#
$intuneToolStripMenuItem.DropDownItems.AddRange(@(
$enrollmentStatusToolStripMenuItem,
$forceSyncAllDevicesToolStripMenuItem,
$listAndroidNonCompliantDevicesToolStripMenuItem,
$listAllAndroidDevicesToolStripMenuItem,
$listAllIntuneApplicationsToolStripMenuItem,
$listAllIOSDevicesToolStripMenuItem,
$listAllWindows10DevicesToolStripMenuItem,
$listIOSNonCompliantDevicesToolStripMenuItem,
$listWin10NonCompliantDevicesToolStripMenuItem))
$intuneToolStripMenuItem.Name = "intuneToolStripMenuItem"
$intuneToolStripMenuItem.Size = New-Object System.Drawing.Size(53, 20)
$intuneToolStripMenuItem.Text = "Intune"
#
# listAllWindows10DevicesToolStripMenuItem
#
$listAllWindows10DevicesToolStripMenuItem.Name = "listAllWindows10DevicesToolStripMenuItem"
$listAllWindows10DevicesToolStripMenuItem.Size = New-Object System.Drawing.Size(263, 22)
$listAllWindows10DevicesToolStripMenuItem.Text = "List All Windows 10 Devices"
$listAllWindows10DevicesToolStripMenuItem.Add_Click(
{$Output = Get-MgDeviceManagementManagedDevice -Filter "contains(operatingSystem, 'Windows')" | Get-MSGraphAllPages | Select-Object DeviceName, userPrincipalName, serialNumber, ID, OperatingSystem, ComplianceState
        $richtextbox1.Text = $Output | Out-GridView -Title "$selectedItem" -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All Windows 10 Devices.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "List All Windows 10 Devices exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All Windows 10 Devices.csv'"
        }
})
#
# listAllIOSDevicesToolStripMenuItem
#
$listAllIOSDevicesToolStripMenuItem.Name = "listAllIOSDevicesToolStripMenuItem"
$listAllIOSDevicesToolStripMenuItem.Size = New-Object System.Drawing.Size(263, 22)
$listAllIOSDevicesToolStripMenuItem.Text = "List All IOS Devices"
$listAllIOSDevicesToolStripMenuItem.Add_Click({
$Output = Get-MgDeviceManagementManagedDevice -Filter "contains(operatingSystem, 'iOS')" | Get-MSGraphAllPages | Select-Object DeviceName, userPrincipalName, serialNumber, ID, OperatingSystem, ComplianceState
        $richtextbox1.Text = $Output | Out-GridView -Title "$selectedItem" -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All IOS Devices.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "List All IOS Devices exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All IOS Devices.csv'"
        }
})
#
# listAllAndroidDevicesToolStripMenuItem
#
$listAllAndroidDevicesToolStripMenuItem.Name = "listAllAndroidDevicesToolStripMenuItem"
$listAllAndroidDevicesToolStripMenuItem.Size = New-Object System.Drawing.Size(263, 22)
$listAllAndroidDevicesToolStripMenuItem.Text = "List All Android Devices"
$listAllAndroidDevicesToolStripMenuItem.Add_Click({
$Output = Get-MgDeviceManagementManagedDevice -Filter "contains(operatingSystem, 'Android')" | Get-MSGraphAllPages | Select-Object DeviceName, userPrincipalName, serialNumber, ID, OperatingSystem, ComplianceState
        $richtextbox1.Text = $Output | Out-GridView -Title "$selectedItem" -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All Android Devices.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "List All Android Devices exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All Android Devices.csv'"
        }
 })
#
# listWin10NonCompliantDevicesToolStripMenuItem
#
$listWin10NonCompliantDevicesToolStripMenuItem.Name = "listWin10NonCompliantDevicesToolStripMenuItem"
$listWin10NonCompliantDevicesToolStripMenuItem.Size = New-Object System.Drawing.Size(263, 22)
$listWin10NonCompliantDevicesToolStripMenuItem.Text = "List Win10 NonCompliant Devices"
$listWin10NonCompliantDevicesToolStripMenuItem.Add_Click({
$Output = Get-MgDeviceManagementManagedDevice -Filter "contains(operatingSystem, 'Windows')" | Get-MSGraphAllPages | Where-Object -Property compliancestate -eq noncompliant | Select-Object DeviceName, userPrincipalName, serialNumber, ID, OperatingSystem, ComplianceState
        $richtextbox1.Text = $Output | Out-GridView -Title "$selectedItem" -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List Win10 NonCompliant Devices.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "Win10 NonCompliant Devices List exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List Win10 NonCompliant Devices.csv'"
        }
})
#
# listIOSNonCompliantDevicesToolStripMenuItem
#
$listIOSNonCompliantDevicesToolStripMenuItem.Name = "listIOSNonCompliantDevicesToolStripMenuItem"
$listIOSNonCompliantDevicesToolStripMenuItem.Size = New-Object System.Drawing.Size(263, 22)
$listIOSNonCompliantDevicesToolStripMenuItem.Text = "List IOS NonCompliant Devices"
$listIOSNonCompliantDevicesToolStripMenuItem.Add_Click({
$Output = Get-MgDeviceManagementManagedDevice -Filter "contains(operatingSystem, 'iOS')" | Get-MSGraphAllPages | Where-Object -Property compliancestate -eq noncompliant | Select-Object DeviceName, userPrincipalName, serialNumber, ID, OperatingSystem, ComplianceState
        $richtextbox1.Text = $Output | Out-GridView -Title "$selectedItem" -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List IOS NonCompliant Devices.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "List IOS NonCompliant Devices exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List IOS NonCompliant Devices.csv'"
        }
})
#
# listAndroidNonCompliantDevicesToolStripMenuItem
#
$listAndroidNonCompliantDevicesToolStripMenuItem.Name = "listAndroidNonCompliantDevicesToolStripMenuItem"
$listAndroidNonCompliantDevicesToolStripMenuItem.Size = New-Object System.Drawing.Size(263, 22)
$listAndroidNonCompliantDevicesToolStripMenuItem.Text = "List Android NonCompliant Devices"
$listAndroidNonCompliantDevicesToolStripMenuItem.Add_Click({
$Output = Get-MgDeviceManagementManagedDevice -Filter "contains(operatingSystem, 'Android')" | Get-MSGraphAllPages | Where-Object -Property compliancestate -eq noncompliant | Select-Object DeviceName, userPrincipalName, serialNumber, ID, OperatingSystem, ComplianceState
        $richtextbox1.Text = $Output | Out-GridView -Title "$selectedItem" -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List Android NonCompliant Devices.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "List Android NonCompliant Devices exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List Android NonCompliant Devices.csv'"
        }
})
#
# listAllIntuneApplicationsToolStripMenuItem
#
$listAllIntuneApplicationsToolStripMenuItem.Name = "listAndroidNonCompliantDevicesToolStripMenuItem"
$listAllIntuneApplicationsToolStripMenuItem.Size = New-Object System.Drawing.Size(263, 22)
$listAllIntuneApplicationsToolStripMenuItem.Text = "List Android NonCompliant Devices"
$listAllIntuneApplicationsToolStripMenuItem.Add_Click({
$Output = Get-IntuneMobileApp 
        foreach ($app in $applications) {
        $assignmentStatus = "Not Assigned"
        if ($app.Assignments.Count -gt 0) {
        $assignmentStatus = "Assigned"
        }
        Write-Host "Application: $($app.DisplayName), Assignment Status: $assignmentStatus"
}        
        $textBox2.Text = $Output | Out-GridView -Title "$selectedItem" -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_IntuneAppsList.csv" -NoTypeInformation -Delimiter ";"
        }
})
#
# forceSyncAllDevicesToolStripMenuItem
#
$forceSyncAllDevicesToolStripMenuItem.Name = "forceSyncAllDevicesToolStripMenuItem"
$forceSyncAllDevicesToolStripMenuItem.Size = New-Object System.Drawing.Size(263, 22)
$forceSyncAllDevicesToolStripMenuItem.Text = "Force Sync All Devices"
$forceSyncAllDevicesToolStripMenuItem.Add_Click({Connect-MsGraph
#Force Intune All Device Sync
$Output = Get-IntuneManagedDevice -Filter "contains(OperatingSystem, 'Windows')" | Get-MSGraphAllPages
Foreach ($Device in $Devices)
            {
            Invoke-IntuneManagedDeviceSyncDevice -managedDeviceId $Device.managedDeviceId
            Write-Host "Sending Sync request to Device with DeviceID $($Device.managedDeviceId)" -ForegroundColor Yellow
            } {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_ForceIntuneSynctoAllDevicesList.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "ForceIntuneSynctoAllDevices exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_ForceIntuneSynctoAllDevicesList.csv'"
            }

})
#
# enrollmentStatusToolStripMenuItem
#
$enrollmentStatusToolStripMenuItem.Name = "enrollmentStatusToolStripMenuItem"
$enrollmentStatusToolStripMenuItem.Size = New-Object System.Drawing.Size(263, 22)
$enrollmentStatusToolStripMenuItem.Text = "Enrollment Status"
$enrollmentStatusToolStripMenuItem.Add_Click({})
#
# office365ToolStripMenuItem
#
$office365ToolStripMenuItem.DropDownItems.AddRange(@(
$distributionGroupsToolStripMenuItem,
$dynamicDistributionGroupsToolStripMenuItem,
$licensesToolStripMenuItem,
$mailboxReportsToolStripMenuItem,
$toolStripSeparator2,
$office365troubleshootingToolStripMenuItem,
$toolStripSeparator3,
$supporttoolsStripMenuItem))
$office365ToolStripMenuItem.Name = "office365ToolStripMenuItem"
$office365ToolStripMenuItem.Size = New-Object System.Drawing.Size(72, 20)
$office365ToolStripMenuItem.Text = "Office 365"
#
# licensesToolStripMenuItem
#
$licensesToolStripMenuItem.DropDownItems.AddRange(@(
$listAllOffice365LicensesToolStripMenuItem,
$listUnlicensedUsersToolStripMenuItem))
$licensesToolStripMenuItem.Name = "licensesToolStripMenuItem"
$licensesToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$licensesToolStripMenuItem.Text = "Licenses"
#
# listAllOffice365LicensesToolStripMenuItem
#
$listAllOffice365LicensesToolStripMenuItem.Name = "listAllOffice365LicensesToolStripMenuItem"
$listAllOffice365LicensesToolStripMenuItem.Size = New-Object System.Drawing.Size(236, 22)
$listAllOffice365LicensesToolStripMenuItem.Text = "List All Office365 Subscriptions"
$listAllOffice365LicensesToolStripMenuItem.Add_Click({
$Output = Get-MsolAccountSku
         $richtextbox1.Text = $Output | Out-GridView -Title "List All Office 365 Subscriptions" -PassThru
        if ($Output) {
            $Output | Select AccountSkuID,ActiveUnits,WarningUnits,ConsumedUnits | Export-CSV "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All Office Subscriptions.csv" -NoTypeInformation -Delimiter ";"
       }
})
#
# listUnlicensedUsersToolStripMenuItem
#
$listUnlicensedUsersToolStripMenuItem.Name = "listUnlicensedUsersToolStripMenuItem"
$listUnlicensedUsersToolStripMenuItem.Size = New-Object System.Drawing.Size(236, 22)
$listUnlicensedUsersToolStripMenuItem.Text = "List Unlicensed Users"
$listUnlicensedUsersToolStripMenuItem.Add_Click({
$Output = Get-MsolUser -All | Where-Object { $_.IsLicensed -match $false }
         $richtextbox1.Text = $Output | Out-GridView -Title "List Unlicensed Users" -PassThru
        if ($Output) {
            $Output | Select DisplayName,UserPrincipalName,isLicensed | Export-CSV "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List Unlicensed Users.csv" -NoTypeInformation -Delimiter ";"
       }
})
#
# mailboxReportsToolStripMenuItem
#
$mailboxReportsToolStripMenuItem.DropDownItems.AddRange(@(
$listAllSharedMailboxesToolStripMenuItem,
$listAllMailboxSizeToolStripMenuItem,
$listInactiveMailboxesToolStripMenuItem,
$listDeletedMailboxesToolStripMenuItem))
$mailboxReportsToolStripMenuItem.Name = "mailboxReportsToolStripMenuItem"
$mailboxReportsToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$mailboxReportsToolStripMenuItem.Text = "Mailbox Reports"
#
# listAllSharedMailboxesToolStripMenuItem
#
$listAllSharedMailboxesToolStripMenuItem.Name = "listAllSharedMailboxesToolStripMenuItem"
$listAllSharedMailboxesToolStripMenuItem.Size = New-Object System.Drawing.Size(205, 22)
$listAllSharedMailboxesToolStripMenuItem.Text = "List All Shared Mailboxes"
$listAllSharedMailboxesToolStripMenuItem.Add_Click({
$Output = Get-EXOMailbox –ResultSize Unlimited –RecipientTypeDetails SharedMailbox
         $richtextbox1.Text = $Output | Out-GridView -Title "List All Shared Mailboxes" -PassThru
        if ($Output) {
            $Output | Select Name,Database,ProhibitSendQuota,ExternalDirectoryObjectID | Export-CSV "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All Shared Mailboxes.csv" -NoTypeInformation -Delimiter ";"
      }
})
#
# listAllMailboxSizeToolStripMenuItem
#
$listAllMailboxSizeToolStripMenuItem.Name = "listAllMailboxSizeToolStripMenuItem"
$listAllMailboxSizeToolStripMenuItem.Size = New-Object System.Drawing.Size(205, 22)
$listAllMailboxSizeToolStripMenuItem.Text = "List All Mailbox Size"
$listAllMailboxSizeToolStripMenuItem.Add_Click({
$Output = Get-EXOMailbox -ResultSize Unlimited | foreach { Get-MailboxStatistics -identity $_.userprincipalName | select Displayname,TotalItemSize}
         $richtextbox1.Text = $Output | Out-GridView -Title "List All Mailbox Size" -PassThru
        if ($Output) {
            $Output | Select Name,Database,ProhibitSendQuota,ExternalDirectoryObjectID | Export-CSV "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List All Maibox Size.csv" -NoTypeInformation -Delimiter ";"
      }
})
#
# listInactiveMailboxesToolStripMenuItem
#
$listInactiveMailboxesToolStripMenuItem.Name = "listInactiveMailboxesToolStripMenuItem"
$listInactiveMailboxesToolStripMenuItem.Size = New-Object System.Drawing.Size(205, 22)
$listInactiveMailboxesToolStripMenuItem.Text = "List Inactive Mailboxes"
$listInactiveMailboxesToolStripMenuItem.Add_Click({})
#
# listDeletedMailboxesToolStripMenuItem
#
$listDeletedMailboxesToolStripMenuItem.Name = "listDeletedMailboxesToolStripMenuItem"
$listDeletedMailboxesToolStripMenuItem.Size = New-Object System.Drawing.Size(205, 22)
$listDeletedMailboxesToolStripMenuItem.Text = "List Deleted Mailboxes"
$listDeletedMailboxesToolStripMenuItem.Add_Click({
$Output = Get-Mailbox -ResultSize Unlimited -SoftDeletedMailbox -Filter "WhenSoftDeleted -ge '$startDate'" | Select-Object UserPrincipalName, WhenSoftDeleted | Sort-Object UserPrincipalName
         $richtextbox1.Text = $Output | Out-GridView -Title "$List Deleted Mailboxes" -PassThru
        if ($Output) {
            $Output | Select Name,Database,ProhibitSendQuota,ExternalDirectoryObjectID | Export-CSV "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List Deleted Mailboxes.csv" -NoTypeInformation -Delimiter ";"
      }
})
#
# dynamicDistributionGroupsToolStripMenuItem
#
$dynamicDistributionGroupsToolStripMenuItem.Name = "dynamicDistributionGroupsToolStripMenuItem"
$dynamicDistributionGroupsToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$dynamicDistributionGroupsToolStripMenuItem.Text = "List Dynamic Distribution Groups"
$dynamicDistributionGroupsToolStripMenuItem.Add_Click({
$Output = Get-DynamicDistributionGroup -ResultSize Unlimited | Select Name, PrimarySMTPAddress, ManagedBy
         $richtextbox1.Text = $Output | Out-GridView -Title "$Dynamic Distribution Groups" -PassThru
        if ($Output) {
            $Output | Select Name,ManagedBy | Export-CSV "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List Dynamic Distribution Groups.csv" -NoTypeInformation -Delimiter ";"
      }
})
#
# distributionGroupsToolStripMenuItem
#
$distributionGroupsToolStripMenuItem.Name = "distributionGroupsToolStripMenuItem"
$distributionGroupsToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$distributionGroupsToolStripMenuItem.Text = "List Distribution Groups"
$distributionGroupsToolStripMenuItem.Add_Click({
$Output = Get-DistributionGroup -ResultSize Unlimited | Select Name, PrimarySMTPAddress, ManagedBy
         $richtextbox1.Text = $Output | Out-GridView -Title "$List Distribution Groups" -PassThru
        if ($Output) {
            $Output | Select Name,ManagedBy | Export-CSV "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_List Distribution Groups.csv" -NoTypeInformation -Delimiter ";"
      }
})
#
# office365troubleshootingToolStripMenuItem
#
$office365troubleshootingToolStripMenuItem.DropDownItems.AddRange(@(
$office365connectivitytestsToolStripMenuItem,
$ExchangeconnectivitytestsToolStripMenuItem,
$microsoftteamsconnectivitytestsToolStripMenuItem,
$microsoft365networkconnectivitytestToolStripMenuItem,
$microsoft365networkhealthstatusToolStripMenuItem,
$microsoftmessageheaderanalysisToolStripMenuItem))
$office365troubleshootingToolStripMenuItem.Name = "office365troubleshootingToolStripMenuItem"
$office365troubleshootingToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$office365troubleshootingToolStripMenuItem.Text = "Office365 Troubleshooting"
#
# ExchangeconnectivitytestsToolStripMenuItem
#
$ExchangeconnectivitytestsToolStripMenuItem.Name = "ExchangeconnectivitytestsToolStripMenuItem"
$ExchangeconnectivitytestsToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$ExchangeconnectivitytestsToolStripMenuItem.Text = "Exchange Connectivity Tests"
$ExchangeconnectivitytestsToolStripMenuItem.Add_Click({
Start-Process msedge "https://testconnectivity.microsoft.com/tests/exchange"

})
#
# office365connectivitytestsToolStripMenuItem
#
$office365connectivitytestsToolStripMenuItem.Name = "office365connectivitytestsToolStripMenuItem"
$office365connectivitytestsToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$office365connectivitytestsToolStripMenuItem.Text = "Office365 Connectivity Tests"
$office365connectivitytestsToolStripMenuItem.Add_Click({
Start-Process msedge "https://testconnectivity.microsoft.com/tests/o365"

})
#
# microsoftteamsconnectivitytestsToolStripMenuItem
#
$microsoftteamsconnectivitytestsToolStripMenuItem.Name = "microsoftteamsconnectivitytestsToolStripMenuItem"
$microsoftteamsconnectivitytestsToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$microsoftteamsconnectivitytestsToolStripMenuItem.Text = "Microsoft Teams Connectivity Tests"
$microsoftteamsconnectivitytestsToolStripMenuItem.Add_Click({
Start-Process msedge "https://testconnectivity.microsoft.com/tests/teams"

})
#
# microsoft365networkconnectivitytestToolStripMenuItem
#
$microsoft365networkconnectivitytestToolStripMenuItem.Name = "microsoft365networkconnectivitytestToolStripMenuItem"
$microsoft365networkconnectivitytestToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$microsoft365networkconnectivitytestToolStripMenuItem.Text = "Microsoft 365 Network Connectivity Tests"
$microsoft365networkconnectivitytestToolStripMenuItem.Add_Click({
Start-Process msedge "https://connectivity.office.com/"

})
#
# microsoft365networkhealthstatusToolStripMenuItem
#
$microsoft365networkhealthstatusToolStripMenuItem.Name = "microsoft365networkhealthstatusToolStripMenuItem"
$microsoft365networkhealthstatusToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$microsoft365networkhealthstatusToolStripMenuItem.Text = "Microsoft 365 Network Health Status"
$microsoft365networkhealthstatusToolStripMenuItem.Add_Click({
Start-Process msedge "https://connectivity.office.com/status"

})
#
# microsoftmessageheaderanalysisToolStripMenuItem
#
$microsoftmessageheaderanalysisToolStripMenuItem.Name = "microsoftmessageheaderanalysisToolStripMenuItem"
$microsoftmessageheaderanalysisToolStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$microsoftmessageheaderanalysisToolStripMenuItem.Text = "Microsoft Message Header Analysis"
$microsoftmessageheaderanalysisToolStripMenuItem.Add_Click({
Start-Process msedge "https://mha.azurewebsites.net/"
})
#
# $supporttoolsStripMenuItem
#
$supporttoolsStripMenuItem.DropDownItems.AddRange(@(
$outlookadvanceddiagnosticsStripMenuItem,
$outlookcalendardiagnosticsStripMenuItem,
$normalscanofficeinstallationStripMenuItem,
$fullscanofficeinstallationStripMenuItem,
$networkconnectivitytestStripMenuItem))
$supporttoolsStripMenuItem.Name = "supporttoolsStripMenuItem"
$supporttoolsStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$supporttoolsStripMenuItem.Text = "Support Tools"
#
#
# outlookadvanceddiagnosticsStripMenuItem
#
$outlookadvanceddiagnosticsStripMenuItem.Name = "outlookadvanceddiagnosticsStripMenuItem"
$outlookadvanceddiagnosticsStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$outlookadvanceddiagnosticsStripMenuItem.Text = "Outlook Advanced Diagnostics"
$outlookadvanceddiagnosticsStripMenuItem.Add_Click({
Start-Process msedge "https://aka.ms/SaRA-OutlookAdvDiagExpExp-sarahome"
})
#
# outlookcalendardiagnosticsStripMenuItem
#
$outlookcalendardiagnosticsStripMenuItem.Name = "outlookcalendardiagnosticsStripMenuItem"
$outlookcalendardiagnosticsStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$outlookcalendardiagnosticsStripMenuItem.Text = "Outlook Calendar Diagnostics"
$outlookcalendardiagnosticsStripMenuItem.Add_Click({
Start-Process msedge "https://aka.ms/SaRA-CalCheck-sarahome"
})
#
# normalscanofficeinstallationStripMenuItem
#
$normalscanofficeinstallationStripMenuItem.Name = "normalscanofficeinstallationStripMenuItem"
$normalscanofficeinstallationStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$normalscanofficeinstallationStripMenuItem.Text = "Normal Scan Office Installation"
$normalscanofficeinstallationStripMenuItem.Add_Click({
Start-Process msedge "https://aka.ms/SaRA-RoiscanNormal-sarahome"
})
#
# fullscanofficeinstallationStripMenuItem
#
$fullscanofficeinstallationStripMenuItem.Name = "fullscanofficeinstallationStripMenuItem"
$fullscanofficeinstallationStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$fullscanofficeinstallationStripMenuItem.Text = "Full Scan Office Installation"
$fullscanofficeinstallationStripMenuItem.Add_Click({
Start-Process msedge "https://aka.ms/SaRA-RoiscanFull-sarahome"
})
#
# networkconnectivitytestStripMenuItem
#
$networkconnectivitytestStripMenuItem.Name = "networkconnectivitytestStripMenuItem"
$networkconnectivitytestStripMenuItem.Size = New-Object System.Drawing.Size(227, 22)
$networkconnectivitytestStripMenuItem.Text = "Network Connectivity Test"
$networkconnectivitytestStripMenuItem.Add_Click({
Start-Process msedge "https://aka.ms/SaRA-NetworkConnectivity-sarahome"
})
#
#
# activeDirectoryToolStripMenuItem
#
$activeDirectoryToolStripMenuItem.DropDownItems.AddRange(@(
$allUsersLogonHistoryToolStripMenuItem,
$backupallgrouppoliciesToolStripMenuItem,
$checkActiveDirectoryReplicationToolStripMenuItem,
$dHCPToolStripMenuItem,
$disabledUsersToolStripMenuItem,
$domainDefaultPasswordPolicyToolStripMenuItem,
$dnsToolStripMenuItem,
$ExportdnsZonesToolStripMenuItem,
$generatePasswordToolStripMenuItem,
$listAllActiveDirectoryUsersToolStripMenuItem,
$listAllDomainControllersToolStripMenuItem,
$listAllDomainSSLCertificatesToolStripMenuItem,
$listAllSecurityGroupsToolStripMenuItem,
$lockedOutUsersToolStripMenuItem,
$searchGroupPolicyToolStripMenuItem,
$unlockandenableuseraccountToolStripMenuItem,
$toolStripSeparator1,
$reportsToolStripMenuItem))
$activeDirectoryToolStripMenuItem.Name = "activeDirectoryToolStripMenuItem"
$activeDirectoryToolStripMenuItem.Size = New-Object System.Drawing.Size(103, 20)
$activeDirectoryToolStripMenuItem.Text = "Active Directory"
#
# generatePasswordToolStripMenuItem
#
$generatePasswordToolStripMenuItem.Name = "generatePasswordToolStripMenuItem"
$generatePasswordToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$generatePasswordToolStripMenuItem.Text = "Generate Password"
$generatePasswordToolStripMenuItem.Add_Click({
    # Create Generate Password form
    $generatePasswordForm = New-Object System.Windows.Forms.Form
    $generatePasswordForm.Text = "Generate Password"
    $generatePasswordForm.Size = New-Object System.Drawing.Size(400, 300)
    $generatePasswordForm.BackColor = [System.Drawing.Color]::AliceBlue
    $generatePasswordForm.StartPosition = "CenterScreen"

    # Create Password Length label
    $labelPasswordLength = New-Object System.Windows.Forms.Label
    $labelPasswordLength.Text = "Password Length:"
    $labelPasswordLength.Location = New-Object System.Drawing.Point(10, 10)
    $labelPasswordLength.AutoSize = $true

    # Create Password Length dropdown
    $dropdownPasswordLength = New-Object System.Windows.Forms.ComboBox
    $dropdownPasswordLength.Location = New-Object System.Drawing.Point(120, 10)
    $dropdownPasswordLength.Size = New-Object System.Drawing.Size(100, 20)
    1..100 | ForEach-Object {
        $dropdownPasswordLength.Items.Add($_)
    }
    $dropdownPasswordLength.SelectedIndex = 0

    # Create Lowercase Characters checkbox
    $checkboxLowercase = New-Object System.Windows.Forms.CheckBox
    $checkboxLowercase.Text = "Lowercase Characters"
    $checkboxLowercase.Location = New-Object System.Drawing.Point(10, 40)
    $checkboxLowercase.Size = New-Object System.Drawing.Size(180, 30)

    # Create Uppercase Characters checkbox
    $checkboxUppercase = New-Object System.Windows.Forms.CheckBox
    $checkboxUppercase.Text = "Uppercase Characters"
    $checkboxUppercase.Location = New-Object System.Drawing.Point(10, 75)
    $checkboxUppercase.Size = New-Object System.Drawing.Size(250, 30)

    # Create Numbers checkbox
    $checkboxNumbers = New-Object System.Windows.Forms.CheckBox
    $checkboxNumbers.Text = "Numbers"
    $checkboxNumbers.Location = New-Object System.Drawing.Point(10, 110)
    $checkboxNumbers.Size = New-Object System.Drawing.Size(180, 30)

    # Create Symbols checkbox
    $checkboxSymbols = New-Object System.Windows.Forms.CheckBox
    $checkboxSymbols.Text = "Symbols"
    $checkboxSymbols.Location = New-Object System.Drawing.Point(10, 145)
    $checkboxSymbols.Size = New-Object System.Drawing.Size(180, 30)

    # Create Password textbox
    $textboxPassword = New-Object System.Windows.Forms.TextBox
    $textboxPassword.Location = New-Object System.Drawing.Point(10, 185)
    $textboxPassword.Size = New-Object System.Drawing.Size(230, 30)
    $textboxPassword.ReadOnly = $false
    $textboxPassword.BackColor = [System.Drawing.Color]::White

    # Create Generate Password button
    $buttonGenerate = New-Object System.Windows.Forms.Button
    $buttonGenerate.Text = "Generate Password"
    $buttonGenerate.Location = New-Object System.Drawing.Point(250, 180)
    $buttonGenerate.BackColor = [System.Drawing.Color]::White
    $buttonGenerate.Size = New-Object System.Drawing.Size(115, 30)

    # Add click event handler for Generate button
    $buttonGenerate.Add_Click({
        # Function to generate password based on selected options
        function Generate-Password {
            $length = $dropdownPasswordLength.SelectedItem
            $characters = @()
            if ($checkboxLowercase.Checked) { $characters += [char[]]'abcdefghijklmnopqrstuvwxyz' }
            if ($checkboxUppercase.Checked) { $characters += [char[]]'ABCDEFGHIJKLMNOPQRSTUVWXYZ' }
            if ($checkboxNumbers.Checked) { $characters += [char[]]'0123456789' }
            if ($checkboxSymbols.Checked) { $characters += [char[]]'!@#$%^&*()-_=+[]{};:,.<>?/' }

            $password = ''
            for ($i = 0; $i -lt $length; $i++) {
                $password += $characters[(Get-Random -Minimum 0 -Maximum $characters.Length)]
            }
            $textboxPassword.Text = $password
        }

        # Generate password
        Generate-Password
    })

    # Add controls to Generate Password form
    $generatePasswordForm.Controls.Add($labelPasswordLength)
    $generatePasswordForm.Controls.Add($dropdownPasswordLength)
    $generatePasswordForm.Controls.Add($checkboxLowercase)
    $generatePasswordForm.Controls.Add($checkboxUppercase)
    $generatePasswordForm.Controls.Add($checkboxNumbers)
    $generatePasswordForm.Controls.Add($checkboxSymbols)
    $generatePasswordForm.Controls.Add($textboxPassword)
    $generatePasswordForm.Controls.Add($buttonGenerate)
    $generatePasswordForm.Controls.Add($copyButton1)
    $generatePasswordForm.Controls.Add($pwdclearButton)


    # Add Password copy button control to Password Form
    $copyButton1.Location = New-Object System.Drawing.Point(10, 220)
    $copyButton1.Name = "copyButton"
    $copyButton1.Size = New-Object System.Drawing.Size(110, 30)
    $copyButton1.BackColor = [System.Drawing.Color]::White
    $copyButton1.TabIndex = 7
    $copyButton1.Text = "Copy to Clipboard"
    $copyButton1.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($textboxPassword[0].Text)
    })
    
    # Add Password Clear Text button control to Password Form
    $pwdclearButton.Location = New-Object System.Drawing.Point(140, 220)
    $pwdclearButton.Name = "pwdclearButton"
    $pwdclearButton.Size = New-Object System.Drawing.Size(110, 30)
    $pwdclearButton.BackColor = [System.Drawing.Color]::White
    $pwdclearButton.TabIndex = 7
    $pwdclearButton.Text = "Clear Password"
    $pwdclearButton.Add_Click({
                    $textboxPassword[0].Clear()
    })
    # Show Generate Password form
    $generatePasswordForm.ShowDialog() | Out-Null
})
#
# lockedOutUsersToolStripMenuItem
#
$lockedOutUsersToolStripMenuItem.Name = "lockedOutUsersToolStripMenuItem"
$lockedOutUsersToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$lockedOutUsersToolStripMenuItem.Text = "Locked Out Users List"
$lockedOutUsersToolStripMenuItem.Add_Click({$Output = Search-ADAccount -Locked | Select Name,DistinguishedName, LockedOut, LastLogon, SID | Out-GridView -Title "Search results for locked out users." -PassThru | Export-Csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\ListAllLockedUserAccounts.csv" -NoTypeInformation -Delimiter ";"})
#
# backupallgrouppoliciesToolStripMenuItem
$backupallgrouppoliciesToolStripMenuItem.Name = "backupallgrouppoliciesToolStripMenuItem"
$backupallgrouppoliciesToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$backupallgrouppoliciesToolStripMenuItem.Text = "Backup All Group Policies"
$backupallgrouppoliciesToolStripMenuItem.Add_Click({
    # Define output directory
    $outputDirectory = "C:\Temp\Group Policy Backups"

# Create the output directory if it doesn't exist
if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
   $BackupGpoOutput = Backup-Gpo -All -Path $outputDirectory
   $BackupGpoOutput = "All Group Policies have been exported to C:\Temp\Group Policy Backups"
    $label.Text = $BackupGpoOutput

    # Display a pop-up window
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup($BackupGpoOutput, 0, "Export Completed", 0x1)
})

# unlockandenableuseraccountToolStripMenuItem
#
$unlockandenableuseraccountToolStripMenuItem.Name = "unlockandenableuseraccountToolStripMenuItem"
$unlockandenableuseraccountToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$unlockandenableuseraccountToolStripMenuItem.Text = "Unlock and Enable User Account"
$unlockandenableuseraccountToolStripMenuItem.Add_Click({

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
#
# Get locked user accounts
$lockedUsers = Search-ADAccount -LockedOut | Select-Object Name, SamAccountName | Sort-Object Name
foreach ($user in $lockedUsers) {
    $lockedUsersComboBox.Items.Add($user.SamAccountName)
}
#
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

})
#
#
# disabledUsersToolStripMenuItem
#
$disabledUsersToolStripMenuItem.Name = "disabledUsersToolStripMenuItem"
$disabledUsersToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$disabledUsersToolStripMenuItem.Text = "Disabled Users List"
$disabledUsersToolStripMenuItem.Add_Click({$Output = Search-ADAccount -AccountDisabled | select Name, DistinguishedName, Enabled, LastLogon, SID | Out-GridView -Title "Search results for disabled users." -PassThru | Export-Csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\ListAllDisabledUserAccounts.csv" -NoTypeInformation -Delimiter ";"})
#
# allUsersLogonHistoryToolStripMenuItem
#
$allUsersLogonHistoryToolStripMenuItem.Name = "allUsersLogonHistoryToolStripMenuItem"
$allUsersLogonHistoryToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$allUsersLogonHistoryToolStripMenuItem.Text = "All Users Logon History"
$allUsersLogonHistoryToolStripMenuItem.Add_Click({$Output = Get-ADUser -Filter * -Properties LastLogon | Select-Object Name, DistinguishedName, SID, @{Name="LastLogon"; Expression={[DateTime]::FromFileTime($_.LastLogon)}} | Out-GridView -Title "Search results for all users logon history." -PassThru | Export-Csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_ListAllUsersLogonHistory.csv" -NoTypeInformation -Delimiter ";"})
#
# searchGroupPolicyToolStripMenuItem
#
$searchGroupPolicyToolStripMenuItem.Name = "searchGroupPolicyToolStripMenuItem"
$searchGroupPolicyToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$searchGroupPolicyToolStripMenuItem.Text = "Search Group Policy"
$searchGroupPolicyToolStripMenuItem.Add_Click({
# Create a form
            $gposearchform = New-Object System.Windows.Forms.Form
            $gposearchform.Text = "Search AD Group Policies"
            $gposearchform.BackColor = [System.Drawing.Color]::AliceBlue
            $gposearchform.Forecolor="#000000"
            $gposearchform.Size = New-Object System.Drawing.Size(400, 200)
            $gposearchform.StartPosition = "CenterScreen"

# Create label for input
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10, 20)
            $label.Size = New-Object System.Drawing.Size(200, 20)
            $label.Text = "Enter search criteria:"
            $gposearchform.Controls.Add($label)

# Create text box for input
            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Location = New-Object System.Drawing.Point(10, 40)
            $textBox.Size = New-Object System.Drawing.Size(200, 20)
            $gposearchform.Controls.Add($textBox)

# Create submit button
            $gposearchbutton = New-Object System.Windows.Forms.Button
            $gposearchbutton.Location = New-Object System.Drawing.Point(10, 70)
            $gposearchbutton.Size = New-Object System.Drawing.Size(80, 30)
            $gposearchbutton.Text = "Submit"
            $gposearchbutton.BackColor = [System.Drawing.Color]::White

$gposearchbutton.Add_Click({
    # Get input from the text box
    $searchCriteria = $textBox.Text
    
    # Search Active Directory group policies
    $groupPolicies = Get-GPO -All | Where-Object { $_.DisplayName -like "*$searchCriteria*" } | Select-Object DisplayName, DomainName, GpoStatus, CreationTime, Modification, WmiFilter | Out-GridView -Title "Search results for $searchCriteria" -Passthru
  
})
$gposearchform.Controls.Add($gposearchbutton)
$gposearchform.Controls.Add($searchCriteria)
$gposearchform.Controls.Add($searchCriteria)

# Show the form
$gposearchform.ShowDialog() | Out-Null
})
#
# domainDefaultPasswordPolicyToolStripMenuItem
#
$domainDefaultPasswordPolicyToolStripMenuItem.Name = "domainDefaultPasswordPolicyToolStripMenuItem"
$domainDefaultPasswordPolicyToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$domainDefaultPasswordPolicyToolStripMenuItem.Text = "Domain Default Password Policy"
$domainDefaultPasswordPolicyToolStripMenuItem.Add_Click({$Output = Get-ADDefaultDomainPasswordPolicy -Verbose | Format-list | Out-String 
                                              $richtextbox1[0].AppendText($output)
                                              })
#
# listAllSecurityGroupsToolStripMenuItem
#
$listAllSecurityGroupsToolStripMenuItem.Name = "listAllSecurityGroupsToolStripMenuItem"
$listAllSecurityGroupsToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$listAllSecurityGroupsToolStripMenuItem.Text = "List All Security Groups"
$listAllSecurityGroupsToolStripMenuItem.Add_Click({$Output = Get-ADGroup -Filter '*' | Select-Object -Property Name, GroupCategory, GroupScope, DistinguishedName | Sort-Object -Property Name | Out-GridView -Title 'Select a group to display members.' -PassThru | ForEach { Get-ADGroupMember -Identity ($SelectedGroup = $PSItem.Name) | Where-Object -Property ObjectClass -eq 'User'} | Select-Object -Property @{Name = 'GroupName';Expression = {$SelectedGroup}}, SamAccountName, DistinguishedName | Out-GridView -Title 'Select a member or members to export' -PassThru  | Select SamAccountName, DistinguishedName | Export-csv "C:\smiadmin\Scripts\Listmembersofgroup.csv" -NoTypeInformation -Delimiter ";"
                                     Write-Host "List members of group" saved to "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_GetAllSecurityGroups.csv"})
#
#
# listAllActiveDirectoryUsersToolStripMenuItem
#
$listAllActiveDirectoryUsersToolStripMenuItem.Name = "listAllActiveDirectoryUsersToolStripMenuItem"
$listAllActiveDirectoryUsersToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$listAllActiveDirectoryUsersToolStripMenuItem.Text = "List All Active Directory Users"
$listAllActiveDirectoryUsersToolStripMenuItem.Add_Click({ $Output = Get-ADUser -Filter * -Property * | Select-object CN, CanonicalName, DistinguishedName, Created, DisplayName, LastLogonDate, Modified, PasswordLastSet, PasswordNotRequired, UserAccountControl, whenChanged  | Out-GridView -PassThru
                                           $Output | Export-Csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_GetAllADUsers.csv" -NoTypeInformation -Delimiter ";"
                                           #$richtextbox1[0].AppendText($output)
                                           Write-Host "List of all Active Directory users exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))__GetAllADUsers.csv'"})
#
# listAllDomainControllersToolStripMenuItem
#
$listAllDomainControllersToolStripMenuItem.Name = "listAllDomainControllersToolStripMenuItem"
$listAllDomainControllersToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$listAllDomainControllersToolStripMenuItem.Text = "List All Domain Controllers"
$listAllDomainControllersToolStripMenuItem.Add_Click({ $output = Get-ADDomainController -Filter * | Select-Object Name, Domain, IPvAddress, Forest, ComputerObjectDN, OperationMasterRoles, OperatingSystem  | Out-String
                                                    $richtextbox1[0].AppendText($output)                    
 })
#
# listAllDomainSSLCertificatesToolStripMenuItema
#
$listAllDomainSSLCertificatesToolStripMenuItem.Name = "listAllDomainSSLCertificatesToolStripMenuItem"
$listAllDomainSSLCertificatesToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$listAllDomainSSLCertificatesToolStripMenuItem.Text = "List All Domain SSL Certificates"
#
# checkActiveDirectoryReplicationToolStripMenuItem
#
$checkActiveDirectoryReplicationToolStripMenuItem.Name = "checkActiveDirectoryReplicationToolStripMenuItem"
$checkActiveDirectoryReplicationToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$checkActiveDirectoryReplicationToolStripMenuItem.Text = "Check Active Directory Replication"
$checkActiveDirectoryReplicationToolStripMenuItem.Add_Click({$output = repadmin /replsum | Format-list | Out-string
    # Display results in text output box
    $richtextbox1.Text = $output
})
#
# dnsToolStripMenuItem
#
$dnsToolStripMenuItem.Name = "dnsToolStripMenuItem"
$dnsToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$dnsToolStripMenuItem.Text = "DNS"
$dnsToolStripMenuItem.Add_Click({
    Start-Process dnsmgmt.msc
})
#
#
# ExportdnsZonesToolStripMenuItem
#
$ExportdnsZonesToolStripMenuItem.Name = "ExportdnsZonesToolStripMenuItem"
$ExportdnsZonesToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$ExportdnsZonesToolStripMenuItem.Text = "Export DNS Zones"
$ExportdnsZonesToolStripMenuItem.Add_Click({
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
    $label.Text = $dnsRecords

    # Display a pop-up window
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup($dnsRecords, 0, "Export Completed", 0x1)
#
})
#
#
# dhcpToolStripMenuItem
#
$dhcpToolStripMenuItem.Name = "dhcpToolStripMenuItem"
$dhcpToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$dhcpToolStripMenuItem.Text = "DHCP"
$dhcpToolStripMenuItem.Add_Click({
    Start-Process dhcpmgmt.msc
})
#
# toolStripSeparator1
#
$toolStripSeparator1.Name = "toolStripSeparator1"
$toolStripSeparator1.Size = New-Object System.Drawing.Size(253, 6)
#
# reportsToolStripMenuItem
#
$reportsToolStripMenuItem.DropDownItems.AddRange(@(
$usersNotLoggedOnIn30DaysToolStripMenuItem,
$usersNotLoggedOnIn60DaysToolStripMenuItem,
$usersNotLoggedOnIn90DaysToolStripMenuItem,
$usersCreatedInPast30daysToolStripMenuItem))


$reportsToolStripMenuItem.Name = "reportsToolStripMenuItem"
$reportsToolStripMenuItem.Size = New-Object System.Drawing.Size(256, 22)
$reportsToolStripMenuItem.Text = "Reports"
#
# usersNotLoggedOnIn30DaysToolStripMenuItem
#
$usersNotLoggedOnIn30DaysToolStripMenuItem.Name = "usersNotLoggedOnIn30DaysToolStripMenuItem"
$usersNotLoggedOnIn30DaysToolStripMenuItem.Size = New-Object System.Drawing.Size(241, 22)
$usersNotLoggedOnIn30DaysToolStripMenuItem.Text = "Users Not Logged on in 30 Days"
$usersNotLoggedOnIn30DaysToolStripMenuItem.Add_Click({
        $Output = Get-ADUser -Filter * -Properties LastLogon | Where-Object {$_.LastLogonDate -lt (Get-Date).AddDays(-30)} | Out-GridView -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Users Logged on 30 days ago.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "Users Logged on 30 days ago exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Users Logged on 30 days ago.csv'"
        }
})
#
# usersNotLoggedOnIn60DaysToolStripMenuItem
#
$usersNotLoggedOnIn60DaysToolStripMenuItem.Name = "usersNotLoggedOnIn60DaysToolStripMenuItem"
$usersNotLoggedOnIn60DaysToolStripMenuItem.Size = New-Object System.Drawing.Size(241, 22)
$usersNotLoggedOnIn60DaysToolStripMenuItem.Text = "Users Not Logged on in 60 Days"
$usersNotLoggedOnIn60DaysToolStripMenuItem.Add_Click({
        $Output = Get-ADUser -Filter * -Properties LastLogon | Where-Object {$_.LastLogonDate -lt (Get-Date).AddDays(-60)} | Out-GridView -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Users Logged on 60 days ago.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "Users Logged on 60 days ago exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Users Logged on 60 days ago.csv'"
        }
})
#
# usersNotLoggedOnIn90DaysToolStripMenuItem
#
$usersNotLoggedOnIn90DaysToolStripMenuItem.Name = "usersNotLoggedOnIn90DaysToolStripMenuItem"
$usersNotLoggedOnIn90DaysToolStripMenuItem.Size = New-Object System.Drawing.Size(241, 22)
$usersNotLoggedOnIn90DaysToolStripMenuItem.Text = "Users Not Logged on in 90 Days"
$usersNotLoggedOnIn90DaysToolStripMenuItem.Add_Click({
        $Output = Get-ADUser -Filter * -Properties LastLogon | Where-Object {$_.LastLogonDate -lt (Get-Date).AddDays(-90)} | Out-GridView -PassThru
        if ($Output) {
            $Output | Export-csv -Path "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Users Logged on 90 days ago.csv" -NoTypeInformation -Delimiter ";"
            Write-Host "Users Logged on 90 days ago exported to 'C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_Users Logged on 90 days ago.csv'"
        }
})
#
# usersCreatedInPast1WeekToolStripMenuItem
#
$usersCreatedInPast30daysToolStripMenuItem.Name = "usersCreatedInPast30daysToolStripMenuItem"
$usersCreatedInPast30daysToolStripMenuItem.Size = New-Object System.Drawing.Size(241, 22)
$usersCreatedInPast30daysToolStripMenuItem.Text = "Users Created in past Month"
$usersCreatedInPast30daysToolStripMenuItem.Add_Click({
# Get current date and calculate date 30 days ago
                            $endDate = Get-Date
                            $startDate = $endDate.AddDays(-30)

# Get AD users created in the last 30 days
                            $users = Get-ADUser -Filter {(WhenCreated -ge $startDate) -and (WhenCreated -lt $endDate)} -Properties WhenCreated | Out-GridView -PassThru

# Check if any users were found
if ($users) {
    # Export the users to a CSV file
    $csvPath = "C:\RTAdmin\Reports Toolset Files\Reports\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_AD users created in the last 30 days.csv"
    $users | Select-Object Name, SamAccountName, WhenCreated | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "AD users created in the last 30 days exported to $csvPath"
} else {
    Write-Host "No AD users created in the last 30 days found."
}
})
                
#
# aboutToolStripMenuItem
#
$aboutToolStripMenuItem.DropDownItems.AddRange(@(
$infoToolStripMenuItem,
$contactToolStripMenuItem))
$aboutToolStripMenuItem.Name = "aboutToolStripMenuItem"
$aboutToolStripMenuItem.Size = New-Object System.Drawing.Size(52, 20)
$aboutToolStripMenuItem.Text = "About"

function OnClick_aboutToolStripMenuItem {
    # Create the About popup window
    $popupForm = New-Object System.Windows.Forms.Form
    $popupForm.Text = "About"
    $popupForm.Size = New-Object System.Drawing.Size(300, 150)
    $popupForm.BackColor = [System.Drawing.Color]::AliceBlue
    $popupForm.StartPosition = "CenterScreen"
    
    # Add header label
    $labelHeader = New-Object System.Windows.Forms.Label
    #$labelHeader.Text = "About"
    $labelHeader.AutoSize = $true
    $labelHeader.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Bold)
    $labelHeader.Location = New-Object System.Drawing.Point(20, 10)
    $popupForm.Controls.Add($labelHeader)
    
    # Add sub-description label
    $labelSubDescription = New-Object System.Windows.Forms.Label
    $labelSubDescription.Text = "Report Tool version 1.0" 
    $labelSubDescription.AutoSize = $true
    $labelSubDescription.Location = New-Object System.Drawing.Point(20, 20)
    
    # Add Label Build Version
    $labelBuild = New-Object System.Windows.Forms.Label
    $labelBuild.Location = New-Object System.Drawing.Point(20, 40)
    $labelBuild.AutoSize = $true
    $labelBuild.Text = "Build version 1.03.2024"
    
    # Add Label Designed by
    $labelDesigner = New-Object System.Windows.Forms.Label
    $labelDesigner.Location = New-Object System.Drawing.Point(20, 60)
    $labelDesigner.AutoSize = $true
    $labelDesigner.Text = "Designed by Caleb Smith"
    
    # Add labels to the About form
    $popupForm.Controls.Add($labelHeader)
    $popupForm.Controls.Add($labelSubDescription)
    $popupForm.Controls.Add($labelBuild)
    $popupForm.Controls.Add($labelDesigner)
    
    # Show the About popup window
    $popupForm.ShowDialog() | Out-Null
}
#
# infoToolStripMenuItem
#
$infoToolStripMenuItem.Name = "infoToolStripMenuItem"
$infoToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$infoToolStripMenuItem.Text = "About Report Toolset"
$infoToolStripMenuItem.Add_Click( { OnClick_aboutToolStripMenuItem } )
#
# contactToolStripMenuItem
#
$contactToolStripMenuItem.Name = "contactToolStripMenuItem"
$contactToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$contactToolStripMenuItem.Text = "Contact"

function OnClick_contactToolStripMenuItem {
    # Create the About popup window
    $contactpopupForm = New-Object System.Windows.Forms.Form
    $contactpopupForm.Text = "Contact Support"
    $contactpopupForm.Size = New-Object System.Drawing.Size(300, 150)
    $contactpopupForm.BackColor = [System.Drawing.Color]::AliceBlue
    $contactpopupForm.StartPosition = "CenterScreen"
    
    # Add header label
    $labelHeader1 = New-Object System.Windows.Forms.Label
    #$labelHeader1.Text = "Contact Support"
    $labelHeader1.AutoSize = $true
    $labelHeader1.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Bold)
    $labelHeader1.Location = New-Object System.Drawing.Point(20, 10)
    $contactpopupForm.Controls.Add($labelHeader1)
    
    # Add sub-description label
    $emaillabel = New-Object System.Windows.Forms.Label
    $emaillabel.Text = "Email: caleb.s1979@gmail.com"
    $emaillabel.AutoSize = $true
    $emaillabel.Location = New-Object System.Drawing.Point(20, 20)
    
    # Add Label Build Version
    $mobilelabel = New-Object System.Windows.Forms.Label
    $mobilelabel.Location = New-Object System.Drawing.Point(20, 40)
    $mobilelabel.AutoSize = $true
    $mobilelabel.Text = "Mobile: +27123456789"
    
    # Add Label Designed by
    $labelDesigner1 = New-Object System.Windows.Forms.Label
    $labelDesigner1.Location = New-Object System.Drawing.Point(20, 60)
    $labelDesigner1.AutoSize = $true
    $labelDesigner1.Text = "Designed by Caleb Smith"
    
    # Add labels to the About form
    $contactpopupForm.Controls.Add($labelHeader1)
    $contactpopupForm.Controls.Add($emaillabel)
    $contactpopupForm.Controls.Add($mobilelabel)
    $contactpopupForm.Controls.Add($labelDesigner1)
    
    # Show the About popup window
    $contactpopupForm.ShowDialog() | Out-Null
}
$contactToolStripMenuItem.Name = "infoToolStripMenuItem"
$contactToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$contactToolStripMenuItem.Text = "Contact"
$contactToolStripMenuItem.Add_Click( { OnClick_contactToolStripMenuItem } )
#
$LinksToolStripMenuItem.DropDownItems.AddRange(@(
$365defenderportalloginToolStripMenuItem,
$aadconnectsyncerrorsToolStripMenuItem,
$advancedhuntingToolStripMenuItem,
$authenticationmethodsToolStripMenuItem,
$azureloginportalToolStripMenuItem,
$azurelinksToolStripMenuItem,
$azureglobalhealthstatusToolStripMenuItem,
$azureresourcegraphexplorerToolStripMenuItem,
$azuresupportportalToolStripMenuItem,
$conditionalaccessToolStripMenuItem,
$defenderlinksToolStripMenuItem,
$graphexplorerToolStripMenuItem,
$intunelinksToolStripMenuItem,
$intuneloginportalToolStripMenuItem,
$office365loginportalToolStripMenuItem,
$office365linksToolStripMenuItem,
$purviewCompliancemanagerToolStripMenuItem,
$purviewCompliancemanagerToolStripMenuItem,
$purvieweDiscoveryToolStripMenuItem,
$sentinelloginportalToolStripMenuItem


))
$LinksToolStripMenuItem.Name = "linksToolStripMenuItem"
$LinksToolStripMenuItem.Size = New-Object System.Drawing.Size(52, 20)
$LinksToolStripMenuItem.Text = "Links"
#
$azurelinksToolStripMenuItem.DropDownItems.AddRange(@(
$azureloginportalToolStripMenuItem,
$azureglobalhealthstatusToolStripMenuItem,
$sentinelloginportalToolStripMenuItem,
$graphexplorerToolStripMenuItem,
$azureresourcegraphexplorerToolStripMenuItem,
$authenticationmethodsToolStripMenuItem,
$conditionalaccessToolStripMenuItem,
$aadconnectsyncerrorsToolStripMenuItem,
$azuresupportportalToolStripMenuItem))
$azurelinksToolStripMenuItem.Name = "azurelinksToolStripMenuItem"
$azurelinksToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$azurelinksToolStripMenuItem.Text = "Azure Links"
#
#
$intunelinksToolStripMenuItem.DropDownItems.AddRange(@(
$intuneloginportalToolStripMenuItem
))
$intunelinksToolStripMenuItem.Name = "intunelinksToolStripMenuItem"
$intunelinksToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$intunelinksToolStripMenuItem.Text = "Intune Links"
#
#
$office365linksToolStripMenuItem.DropDownItems.AddRange(@(
$office365loginportalToolStripMenuItem
))
$office365linksToolStripMenuItem.Name = "office365linksToolStripMenuItem"
$office365linksToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$office365linksToolStripMenuItem.Text = "Office365 Links"
#
#
$defenderlinksToolStripMenuItem.DropDownItems.AddRange(@(
$365defenderportalloginToolStripMenuItem,
$purviewCompliancemanagerToolStripMenuItem,
$purvieweDiscoveryToolStripMenuItem,
$advancedhuntingToolStripMenuItem))
$defenderlinksToolStripMenuItem.Name = "defenderlinksToolStripMenuItem"
$defenderlinksToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$defenderlinksToolStripMenuItem.Text = "Microsoft 365 Defender Links"
#
#
# azureloginportalToolStripMenuItem
#
$azureloginportalToolStripMenuItem.Name = "LinksToolStripMenuItem"
$azureloginportalToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$azureloginportalToolStripMenuItem.Text = "Azure Login Portal"
$azureloginportalToolStripMenuItem.Add_Click({Start-Process msedge "https://portal.azure.com"
})
#
# intuneloginportalToolStripMenuItem
#
$intuneloginportalToolStripMenuItem.Name = "LinksToolStripMenuItem"
$intuneloginportalToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$intuneloginportalToolStripMenuItem.Text = "Intune Login Portal"
$intuneloginportalToolStripMenuItem.Add_Click({Start-Process msedge "https://intune.microsoft.com/"
})
#
# intuneloginportalToolStripMenuItem
#
$office365loginportalToolStripMenuItem.Name = "LinksToolStripMenuItem"
$office365loginportalToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$office365loginportalToolStripMenuItem.Text = "Office365 Login Portal"
$office365loginportalToolStripMenuItem.Add_Click({Start-Process msedge "https://www.microsoft365.com/"
})
#
# azureglobalhealthstatusToolStripMenuItem
#
$azureglobalhealthstatusToolStripMenuItem.Name = "LinksToolStripMenuItem"
$azureglobalhealthstatusToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$azureglobalhealthstatusToolStripMenuItem.Text = "Azure Global Health Status"
$azureglobalhealthstatusToolStripMenuItem.Add_Click({Start-Process msedge "https://azure.status.microsoft/en-us/status"
})
#
# sentinelloginportalToolStripMenuItem
#
$sentinelloginportalToolStripMenuItem.Name = "LinksToolStripMenuItem"
$sentinelloginportalToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$sentinelloginportalToolStripMenuItem.Text = "Sentinel Login Portal"
$sentinelloginportalToolStripMenuItem.Add_Click({Start-Process msedge "https://portal.azure.com/#blade/Microsoft_Azure_Security_Insights/WorkspaceSelectorBlade?"
})
#
# 365defenderportalloginToolStripMenuItem
#
$365defenderportalloginToolStripMenuItem.Name = "LinksToolStripMenuItem"
$365defenderportalloginToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$365defenderportalloginToolStripMenuItem.Text = "365 Defender Login Portal"
$365defenderportalloginToolStripMenuItem.Add_Click({Start-Process msedge "https://security.microsoft.com/"
})
#
# graphexplorerToolStripMenuItem
#
$graphexplorerToolStripMenuItem.Name = "LinksToolStripMenuItem"
$graphexplorerToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$graphexplorerToolStripMenuItem.Text = "Graph Explorer"
$graphexplorerToolStripMenuItem.Add_Click({Start-Process msedge "https://developer.microsoft.com/en-us/graph/graph-explorer"
})
#
# azureresourcegraphexplorerToolStripMenuItem
#
$azureresourcegraphexplorerToolStripMenuItem.Name = "LinksToolStripMenuItem"
$azureresourcegraphexplorerToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$azureresourcegraphexplorerToolStripMenuItem.Text = "Azure Resource Graph Explorer"
$azureresourcegraphexplorerToolStripMenuItem.Add_Click({Start-Process msedge "https://portal.azure.com/#view/HubsExtension/ArgQueryBlade?"
})
#
# purviewCompliancemanagerToolStripMenuItem
#
$purviewCompliancemanagerToolStripMenuItem.Name = "LinksToolStripMenuItem"
$purviewCompliancemanagerToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$purviewCompliancemanagerToolStripMenuItem.Text = "Purview Compliance Manager"
$purviewCompliancemanagerToolStripMenuItem.Add_Click({Start-Process msedge "https://compliance.microsoft.com/compliancemanager"
})
#
# purvieweDiscoveryToolStripMenuItem
#
$purvieweDiscoveryToolStripMenuItem.Name = "LinksToolStripMenuItem"
$purvieweDiscoveryToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$purvieweDiscoveryToolStripMenuItem.Text = "Purview eDiscovery"
$purvieweDiscoveryToolStripMenuItem.Add_Click({Start-Process msedge "https://compliance.microsoft.com/advancedediscovery"
})
#
# advancedhuntingToolStripMenuItem
#
$advancedhuntingToolStripMenuItem.Name = "LinksToolStripMenuItem"
$advancedhuntingToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$advancedhuntingToolStripMenuItem.Text = "Advanced Hunting"
$advancedhuntingToolStripMenuItem.Add_Click({Start-Process msedge "https://security.microsoft.com/v2/advanced-hunting"
})
#
# authenticationmethodsToolStripMenuItem
#
$authenticationmethodsToolStripMenuItem.Name = "LinksToolStripMenuItem"
$authenticationmethodsToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$authenticationmethodsToolStripMenuItem.Text = "MFA Authentication Methods"
$authenticationmethodsToolStripMenuItem.Add_Click({Start-Process msedge "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AuthenticationMethodsMenuBlade/~/AdminAuthMethods"
})
#
# conditionalaccessToolStripMenuItem
#
$conditionalaccessToolStripMenuItem.Name = "LinksToolStripMenuItem"
$conditionalaccessToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$conditionalaccessToolStripMenuItem.Text = "Conditional Access"
$conditionalaccessToolStripMenuItem.Add_Click({Start-Process msedge "https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/~/Policies?"
})
#
# aadconnectsyncerrorsToolStripMenuItem
#
$aadconnectsyncerrorsToolStripMenuItem.Name = "LinksToolStripMenuItem"
$aadconnectsyncerrorsToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$aadconnectsyncerrorsToolStripMenuItem.Text = "AAD Connect Sync Errors"
$aadconnectsyncerrorsToolStripMenuItem.Add_Click({Start-Process msedge "https://entra.microsoft.com/#view/Microsoft_Azure_ADHybridHealth/AadHealthMenuBlade/~/SyncErros/Microsoft_Azure_ADHybridHealth/AadHealthMenuBlade"
})
#
# azuresupportportalToolStripMenuItem
#
$azuresupportportalToolStripMenuItem.Name = "LinksToolStripMenuItem"
$azuresupportportalToolStripMenuItem.Size = New-Object System.Drawing.Size(193, 22)
$azuresupportportalToolStripMenuItem.Text = "Azure Support Portal"
$azuresupportportalToolStripMenuItem.Add_Click({Start-Process msedge "https://entra.microsoft.com/#view/Microsoft_Azure_Support/NewSupportRequestV3Blade/callerName/ActiveDirectory/issueType/technical"
})
#
# toolStripSeparator1
#
$toolStripSeparator2.Name = "toolStripSeparator2"
$toolStripSeparator2.Size = New-Object System.Drawing.Size(253, 6)
#
# monthCalendar1
#
$monthCalendar1.BackColor = [System.Drawing.Color]::AliceBlue
$monthCalendar1.FirstDayOfWeek = [System.Windows.Forms.Day]::Monday
$monthCalendar1.Location = New-Object System.Drawing.Point(595, 55)
$monthCalendar1.MaximumSize = New-Object System.Drawing.Size(280, 280)
$monthCalendar1.Name = "monthCalendar1"
$monthCalendar1.TabIndex = 1
#
# toolStripContainer1
#
# toolStripContainer1.BottomToolStripPanel
#
$toolStripContainer1.BottomToolStripPanel.BackColor = [System.Drawing.Color]::AliceBlue
#
# toolStripContainer1.ContentPanel
#

$toolStripContainer1.ContentPanel.BackColor = [System.Drawing.Color]::AliceBlue
$toolStripContainer1.ContentPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$toolStripContainer1.ContentPanel.Controls.Add($pictureBox1)
$toolStripContainer1.ContentPanel.Controls.Add($menuStrip1)
$toolStripContainer1.ContentPanel.Controls.Add($monthCalendar1)
$toolStripContainer1.ContentPanel.Size = New-Object System.Drawing.Size(860, 226)
$toolStripContainer1.Location = New-Object System.Drawing.Point(12, 0)
$toolStripContainer1.Name = "toolStripContainer1"
$toolStripContainer1.Size = New-Object System.Drawing.Size(860, 251)
$toolStripContainer1.TabIndex = 0
$toolStripContainer1.Text = "toolStripContainer1"
#
# richTextBox1
#
$richTextBox1.Location = New-Object System.Drawing.Point(12, 267)
$richTextBox1.Name = "richTextBox1"
$richTextBox1.Size = New-Object System.Drawing.Size(858, 318)
$richTextBox1.Multiline = $true
$richTextBox1.ScrollBars = "Vertical"
$richTextBox1.TabIndex = 1
$richTextBox1.Text = ""
$richtextbox1.Font = New-Object System.Drawing.Font("Consolas", 10)
function Add-RichTextBox1{
		[CmdletBinding()]
		param ($text)
		$richTextBox1.Text += "$text;"
		$richTextBox1.Text += "`n# # # # # # # # # #`n"
	}
#
# button1
#
$exportButton.Location = New-Object System.Drawing.Point(10, 593)
$exportButton.Name = "button1"
$exportButton.Size = New-Object System.Drawing.Size(91, 34)
$exportButton.BackColor = [System.Drawing.Color]::White
$exportButton.TabIndex = 5
$exportButton.Text = "Export to csv"
$exportButton.Add_Click({
    $richtextbox1[0].Text | Out-File -FilePath "c:\temp\output.csv"
    Write-Host "c:\temp\output.csv"
})

function OnClick_button1 {
	[void][System.Windows.Forms.MessageBox]::Show("The event handler button1.Add_Click is not implemented.")
}

$exportButton.Add_Click( { OnClick_button1 } )
$Form1.Controls.Add($exportButton)
#
# clearButton
#
$clearButton.Location = New-Object System.Drawing.Point(120, 593)
$clearButton.Name = "clearButton"
$clearButton.Size = New-Object System.Drawing.Size(91, 34)
$clearButton.BackColor = [System.Drawing.Color]::White
$clearButton.TabIndex = 6
$clearButton.Text = "Clear"
$clearButton.Add_Click({
    $richtextbox1[0].Clear()
})
$Form1.Controls.Add($clearButton)
#
# button3
#
$copyButton.Location = New-Object System.Drawing.Point(229, 593)
$copyButton.Name = "copyButton"
$copyButton.Size = New-Object System.Drawing.Size(110, 34)
$copyButton.BackColor = [System.Drawing.Color]::White
$copyButton.TabIndex = 7
$copyButton.Text = "Copy to Clipboard"
$copyButton.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($richtextbox1[0].Text)
})
$Form1.Controls.Add($copyButton)
#
# button4
#
$button_formexit.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
$button_formexit.Location = New-Object System.Drawing.Point(779, 593)
$button_formexit.Name = "button_formexit"
$button_formexit.Size = New-Object System.Drawing.Size(91, 34)
$button_formexit.BackColor = [System.Drawing.Color]::White
$button_formexit.TabIndex = 8
$button_formexit.Text = "Exit"
$button_formexit.Add_Click({$confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Exit Confirmation", "YesNo", "Question")
    if ($confirm -eq "Yes") {
        $form1.Close()
    }
})
$Form1.Controls.Add($button_formexit)
#
# pictureBox1
#
$pictureBox1.Location = New-Object System.Drawing.Point(10, 45)
$imageURL = "https://25168988.fs1.hubspotusercontent-eu1.net/hubfs/25168988/Sweden/Advania-black-RGB.png"
#$imageURL = "https://25168988.fs1.hubspotusercontent-eu1.net/hubfs/25168988/Com/Press/Advania-white-RGB.png"
$image = [System.Net.WebRequest]::Create($imageUrl).GetResponse().GetResponseStream()
$pictureBox1.Image = [System.Drawing.Image]::FromStream($image)
$pictureBox1.SizeMode = "Stretch"  # Fit the image within the PictureBox
$pictureBox1.Size = New-Object System.Drawing.Size(440, 200)

# Form1
#
$Form1.BackColor = [System.Drawing.Color]::AliceBlue
$Form1.ClientSize = New-Object System.Drawing.Size(883, 637)
$Form1.StartPosition = "CenterScreen"
$Form1.Controls.Add($button4)
$Form1.Controls.Add($button3)
$Form1.Controls.Add($button2)
$Form1.Controls.Add($button1)
$Form1.Controls.Add($toolStripContainer1)
$Form1.Controls.Add($richTextBox1)
$Form1.Name = "Form1"
$Form1.Text = "Reports version 1.0"

$Form1.ShowDialog() | Out-Null
