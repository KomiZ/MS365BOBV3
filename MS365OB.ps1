# Microsoft Custom 365 User off-boarding script. Version 1.0 By: Komail Chaudhry
# The script is updated regularly to keep up to date with Microsoft APIs (Current revison April 2024)
# Ensure you're running this script in PowerShell 5.1 or newer.
# Ensure you have installed all the preq Module's needed

# Import AzureAD and Exchange Online Management V3 module
Import-Module AzureAD
Import-Module ExchangeOnlineManagement

# Connect to Azure AD and Exchange Online
$azureConnection = Connect-AzureAD
$exchangeConnection = Connect-ExchangeOnline

# Collect all initial input
$UserPrincipalNames = Read-Host "Enter the User Principal Name(s) of the user(s) to manage, separated by commas"
$UsersArray = $UserPrincipalNames -split ','
$PreserveGroupNames = Read-Host "Enter the names of groups to preserve, separated by commas"
$PreserveGroupsArray = $PreserveGroupNames -split ','

# Initialize summary data
$Summary = @()

# Function to remove a user from all groups except specified groups
function Remove-UserFromAllGroups {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserObjectId,
        [string[]]$SkipGroupNames
    )

    $Groups = Get-AzureADUserMembership -ObjectId $UserObjectId
    $RemovedGroups = @()
    $SkippedGroups = @()

    foreach ($Group in $Groups) {
        if ($Group.DisplayName -notin $SkipGroupNames) {
            try {
                Remove-AzureADGroupMember -ObjectId $Group.ObjectId -MemberId $UserObjectId
                $RemovedGroups += $Group.DisplayName
            } catch {
                Write-Output "Failed to remove from $($Group.DisplayName). Error: $($_.Exception.Message)"
            }
        } else {
            $SkippedGroups += $Group.DisplayName
        }
    }
    return @{ Removed = $RemovedGroups; Skipped = $SkippedGroups }
}

# Process each user
foreach ($UserPrincipalName in $UsersArray) {
    $UserPrincipalName = $UserPrincipalName.Trim()
    $UserSummary = @{"User" = $UserPrincipalName; "Actions" = @()}

    try {
        # Block the user from signing in
        Set-AzureADUser -ObjectId $UserPrincipalName -AccountEnabled $false
        $UserSummary["Actions"] += "Sign-in blocked."

        # Convert user mailbox to a shared mailbox
        $Mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
        if ($Mailbox) {
            Set-Mailbox -Identity $UserPrincipalName -Type Shared
            $UserSummary["Actions"] += "Mailbox converted to shared."
            Start-Sleep -Seconds 10  # Pausing to ensure mailbox conversion completes
        } else {
            $UserSummary["Actions"] += "No mailbox found."
        }

        # Unassign all licenses
        $Licenses = Get-AzureADUser -ObjectId $UserPrincipalName | Select -ExpandProperty AssignedLicenses
        $LicenseIds = $Licenses | Select -ExpandProperty SkuId
        if ($LicenseIds.Count -gt 0) {
            $LicensesToRemove = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $LicensesToRemove.AddLicenses = @()
            $LicensesToRemove.RemoveLicenses = $LicenseIds
            Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $LicensesToRemove
            $UserSummary["Actions"] += "Licenses unassigned."
        }

        # Remove user from all groups except the specified ones
        $UserObject = Get-AzureADUser -ObjectId $UserPrincipalName
        $GroupChanges = Remove-UserFromAllGroups -UserObjectId $UserObject.ObjectId -SkipGroupNames $PreserveGroupsArray
        $UserSummary["Actions"] += "Removed from groups: $($GroupChanges.Removed -join ', ')"

    } catch {
        $UserSummary["Actions"] += "Error: $_"
    }

    $Summary += $UserSummary
}

# Disconnect from services
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-AzureAD

# Display summary
$SummaryText = $Summary | ForEach-Object {
    "$($_.User):`n$($_.Actions -join "`n")`n"
}
$AdminInfo = "Admin ID: $($azureConnection.Account.Id)"
$CompletionMessage = "Script execution completed. Admin info: $AdminInfo `nSummary of actions taken:`n$SummaryText"

Write-Host $CompletionMessage -ForegroundColor Green
Read-Host -Prompt "Press Enter to exit"
