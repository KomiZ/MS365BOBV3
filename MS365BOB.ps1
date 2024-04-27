# Microsoft Custom 365 User off-boarding script. Version 1.0 By: Komail Chaudhry
# The script is updated regularly to keep up to date with Microsoft APIs (Current revision April 2024)
# Ensure you're running this script in PowerShell 5.1 or newer.
# Ensure you have installed all the preq Module's needed

# Import AzureAD and Exchange Online Management V3 module
Import-Module AzureAD
Import-Module ExchangeOnlineManagement

# Connect to Azure AD and Exchange Online
$azureConnection = Connect-AzureAD
$exchangeConnection = Connect-ExchangeOnline

# Define the path for user emails and report file based on the script's current directory
$UserFilePath = ".\user_emails.txt"
$DateTime = Get-Date -Format "yyyyMMddHHmmss"
$ReportPath = ".\OffboardingReport_$DateTime.txt"

# Check if the user emails file exists
if (-Not (Test-Path $UserFilePath)) {
    Write-Host "User emails file not found at $UserFilePath" -ForegroundColor Red
    exit
}

# Read User Principal Names from a file
$UserPrincipalNames = Get-Content -Path $UserFilePath
$UsersArray = $UserPrincipalNames -split ','

# Request names of groups to preserve
$PreserveGroupNames = Read-Host "Enter the names of groups to preserve, separated by commas"
$PreserveGroupsArray = $PreserveGroupNames -split ','

# Initialize report data
$Report = @()

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
                "Failed to remove from $($Group.DisplayName). Error: $($_.Exception.Message)" | Out-File -FilePath $ReportPath -Append
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
    $UserReport = @{"User" = $UserPrincipalName; "Actions" = @()}

    try {
        # Block the user from signing in
        Set-AzureADUser -ObjectId $UserPrincipalName -AccountEnabled $false
        $UserReport["Actions"] += "Sign-in blocked."

        # Convert user mailbox to a shared mailbox
        $Mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
        if ($Mailbox) {
            Set-Mailbox -Identity $UserPrincipalName -Type Shared
            $UserReport["Actions"] += "Mailbox converted to shared."
            Start-Sleep -Seconds 10  # Pausing to ensure mailbox conversion completes
        } else {
            $UserReport["Actions"] += "No mailbox found."
        }

        # Unassign all licenses
        $Licenses = Get-AzureADUser -ObjectId $UserPrincipalName | Select -ExpandProperty AssignedLicenses
        $LicenseIds = $Licenses | Select -ExpandProperty SkuId
        if ($LicenseIds.Count -gt 0) {
            $LicensesToRemove = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $LicensesToRemove.AddLicenses = @()
            $LicensesToRemove.RemoveLicenses = $LicenseIds
            Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $LicensesToRemove
            $UserReport["Actions"] += "Licenses unassigned."
        }

        # Remove user from all groups except the specified ones
        $UserObject = Get-AzureADUser -ObjectId $UserPrincipalName
        $GroupChanges = Remove-UserFromAllGroups -UserObjectId $UserObject.ObjectId -SkipGroupNames $PreserveGroupsArray
        $UserReport["Actions"] += "Removed from groups: $($GroupChanges.Removed -join ', ')"

    } catch {
        $UserReport["Actions"] += "Error: $_"
    }

    $Report += $UserReport
}

# Write report to a file
$Report | ForEach-Object {
    "$($_.User):`n$($_.Actions -join "`n")`n" | Out-File -FilePath $ReportPath -Append
}

# Disconnect from services
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-AzureAD

# Display final message
$TotalProcessed = $UsersArray.Count
Write-Host "Off-boarding process completed for $TotalProcessed users. Detailed report is available at $ReportPath." -ForegroundColor Green
