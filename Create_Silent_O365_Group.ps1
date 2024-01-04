#Requires -Modules AzureAD, @{ModuleName="ExchangeOnlineManagement"; RequiredVersion="3.0.0"}
# Install required modules (run as admin)
#   Install-Module ExchangeOnlineManagement
#   Install-Module AzureAD

# Set Group Owner below, Note this user MUST have a valid Exchange license.
$GroupOwner = ""

# Version History
# 1.0 - 20/3/23 - Jordan Collins

# Create Group function
function New-SilentO365Group {
    [CmdletBinding()]
    # Set parameters for group creation
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline=$True)]
        [string] $DisplayName,
        [Parameter(Mandatory = $True, ValueFromPipeline=$True)]
        [string] $EmailAddress,
        [Parameter(Mandatory = $True, ValueFromPipeline=$True)]
        [string] $Description
    )
    Write-Host "Creating Group: $DisplayName"
    # Create M365 Group without a SharePoint site
    New-UnifiedGroup -DisplayName "$DisplayName" -Notes "$Description" -AccessType Private -EmailAddresses "$EmailAddress" -Owner "$GroupOwner" | Out-Null
    # Remove Welcome Message email and Hide from Outlook    
    Set-UnifiedGroup "$DisplayName" -UnifiedGroupWelcomeMessageEnabled:$False -HiddenFromExchangeClientsEnabled:$True -AutoSubscribeNewMembers:$True
    # Remove Owner as member
    $GroupObjectId = (Get-AzureADGroup -Filter "DisplayName eq '$DisplayName'").ObjectId
    Remove-AzureADGroupMember -ObjectId "$GroupObjectId" -MemberId "$GroupOwnerId"
}
function Disconnect-Cloud {
    Disconnect-AzureAD -Confirm:$False
    Disconnect-ExchangeOnline -Confirm:$False
}

# Ask if Owner Email set in script is ok to use, prompt for new if not
$ChoiceYes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$ChoiceNo = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$ChoiceOptions = [System.Management.Automation.Host.ChoiceDescription[]]($ChoiceYes, $ChoiceNo)
$ChoiceTitle = "Owner Email"
$ChoiceMessage = "Currently configured to $GroupOwner. Ok to continue?"
$ChoiceResult = $Host.ui.PromptForChoice($ChoiceTitle, $ChoiceMessage, $ChoiceOptions, 0)
switch ($ChoiceResult) {
    0 {}
    1 {
        $GroupOwner = Read-Host "Group Owner UPN"
    }
}

# Import module and connect
Import-Module ExchangeOnlineManagement
Import-Module AzureAD
try {
    Write-Host "Please login to Azure AD in new popup for Exchange Online"
    Connect-ExchangeOnline -ShowBanner:$False
    Write-Host "Please login to Azure AD in new popup for AzureAD"
    $Null = Connect-AzureAD
} catch {
    Write-Warning "Failed to authenticate to Azure AD. please try again"
    Pause
    Exit
}

# Get Group Owner Object Id
$GroupOwnerId = (Get-AzureADUser -Filter "UserPrincipalName eq '$GroupOwner'").ObjectId

# Ask if single account or bulk from CSV
$ChoiceSingle = New-Object System.Management.Automation.Host.ChoiceDescription "&Single"
$ChoiceCSV = New-Object System.Management.Automation.Host.ChoiceDescription "&CSV"
$ChoiceOptions = [System.Management.Automation.Host.ChoiceDescription[]]($ChoiceSingle, $ChoiceCSV)
$ChoiceTitle = "Creation Type"
$ChoiceMessage = "Create single group or bulk from a CSV?"
$ChoiceResult = $Host.ui.PromptForChoice($ChoiceTitle, $ChoiceMessage, $ChoiceOptions, 0)

# Results from choice
switch ($ChoiceResult) {
    # Single group creation
    0 {
        Write-Host "Single group creation:"
        # Prompt for variables
        $DisplayName = Read-Host "Group Display Name"
        $EmailAddress = Read-Host "Group Email Address"
        $Description = Read-Host "Group Description"
        try {
            # Create Group
            New-SilentO365Group -DisplayName "$DisplayName" -EmailAddress "$EmailAddress" -Description "$Description"
        } catch {
            Write-Warning "Failed to create group. Exiting..."
            Disconnect-Cloud
            Pause
            Exit
        }
    }
    # Bulk group creation from CSV
    1 {
        Write-Host "Bulk group creation from CSV:"
        Write-Warning "CSV must be in the following formation with headers: DisplayName, EmailAddress, Description"
        Write-Host "Select CSV file in window popup."
        #Show dialog box for CSV selection
        Add-Type -AssemblyName System.Windows.Forms
        $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $FileDialog.Filter = "CSV files (*.csv)|*.csv"
        $FileDialog.Multiselect = $false
        if ($FileDialog.ShowDialog() -ne 'OK') {
            Write-Host "File not selected. Exiting..."
            Disconnect-Cloud
            Pause
            Exit
        } else {
            # Get path of and import CSV
            $CSV = $FileDialog.FileName
            $Groups = Import-CSV -Path $CSV
            # Loop through each row of the CSV
            foreach ($Group in $Groups) {
                # Set variables based off column headers
                $DisplayName = $Group.DisplayName
                $EmailAddress = $Group.EmailAddress
                $Description = $Group.Description
                try {
                    # Create Group
                    New-SilentO365Group -DisplayName "$DisplayName" -EmailAddress "$EmailAddress" -Description "$Description"
                } catch {
                    Write-Warning "Failed to create group. Exiting..."
                    Disconnect-Cloud
                    Pause
                    Exit
                }
            }
        }
    }
}
