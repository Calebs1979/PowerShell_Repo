# Function to get logon history for a specific user
function Get-UserLogonHistory {
    param (
        [string]$username,
        [int]$days
    )

    # Get current date
    $endDate = Get-Date

    # Calculate start date based on selected days
    $startDate = $endDate.AddDays(-$days)

    # Get logon events for the user within the specified time frame
    $logonEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'Security'
        ID = 4624, 4634
        StartTime = $startDate
        EndTime = $endDate
    } | Where-Object { $_.Properties[5].Value -eq $username }

    # Output logon events
    return $logonEvents
}

# Get all Active Directory users
$users = Get-ADUser -Filter *

# Iterate through each user and retrieve logon history
foreach ($user in $users) {
    $username = $user.SamAccountName
    $logonEvents = Get-UserLogonHistory -username $username -days 30

    if ($logonEvents) {
        Write-Host "Logon history for user: $username"
        $logonEvents | Format-Table -Property TimeCreated, @{Name="Computer"; Expression={$_.Properties[2].Value}} -AutoSize
    } else {
        Write-Host "No logon events found for user: $username"
    }
}
