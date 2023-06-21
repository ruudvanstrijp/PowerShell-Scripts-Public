<# 
.SYNOPSIS
 
    Get-TeamsAssignedNumbers.ps1 collects assigned phone numbers from all Microsoft Teams.
 
.DESCRIPTION
    Original Author: Andrew Morpeth
    Contact: https://ucgeek.co/
    
    Modified by: Ruud van Strijp

    This script queries Microsoft Teams for assigned numbers and displays in a formatted table with the option to export to CSV. 
    During processing LineURI's are run against a regex pattern to extract the DDI/DID and the extension to a separate column.
    
    This script collects Microsoft Teams objects including:
    Users, Meeting Rooms, Online Application Instances (Resource Accounts)

    This script does not collect objects from on-premises environments even if in hybrid, instead use this script - https://gallery.technet.microsoft.com/office/Lync-Get-All-Assigned-8c1328a0
    
    This script is provided as-is, no warrenty is provided or implied.The author is NOT responsible for any damages or data loss that may occur
    through the use of this script.  Always test before using in a production environment. This script is free to use for both personal and 
    business use, however, it may not be sold or included as part of a package that is for sale. A Service Provider may include this script 
    as part of their service offering/best practices provided they only charge for their time to implement and support.
.NOTES
    v1.0 - Initial release       
    v1.1 - Now using Microsoft Teams PowerShell module
    v1.2 - Changed login method and file save location. Added HTML export by default
    v1.3 - Added UPN. Added HTML table width for better readability
    v1.4 - Added EV columns; Calling and Voice Routing Policies
    v1.5 - Changed for module 3.0.0
    v1.6 - Get Dial Plan, fix first name, merge Resource Account and User, make table wider

.NOTES
Microsoft Teams module 4.0.0 or higher needs to be installed into PowerShell. 5.0 is heavily recommended because of it's speed
Uninstall-Module -Name MicrosoftTeams -AllVersions
Install-Module -Name MicrosoftTeams -Force -Scope AllUsers
#>

Param (
    [switch]$onlyRA
)


$debug = $true

$teamsModuleVersion = (Get-InstalledModule -Name MicrosoftTeams).Version
if ($teamsModuleVersion -lt 4.0.0) {
    Write-Host "  WARNING: Module Version older than 4.0.0 will be deprecated soon. This script might not run well" -ForegroundColor red
}
if ($teamsModuleVersion -lt 5.0.0) {
    Write-Host "  WARNING: Module Version older than 5.0.0 will run a lot slower" -ForegroundColor red
}

try {
    if ($debug -like $true) {
        Write-Host "  DEBUG: Trying to connect to existing session..." -ForegroundColor DarkGray
    }
    Get-CsTenant | Out-Null
}
Catch {
    Write-Host "  DEBUG: Could not connect to existing session, starting new session" -ForegroundColor DarkGray
    Connect-MicrosoftTeams
}

Write-Host "  Connected to tenant: " -ForegroundColor White -NoNewLine
Write-Host (Get-CsTenant).DisplayName -ForegroundColor Green

#Settings ##############################
#. "_Settings.ps1" | Out-Null
$FileName = "TeamsAssignedNumbers_" + (Get-Date -Format s).replace(":", "-") 
$FilePath = $PSScriptRoot + "\" + $FileName
$OutputType = "HTML" #OPTIONS: CSV - Outputs CSV to specified FilePath, CONSOLE - Outputs to console


##############################

$Regex1 = '^(?:tel:)?(?:\+)?(\d+)(?:;ext=(\d+))?(?:;([\w-]+))?$'
$Array1 = @()
#Get Users with LineURI
#$UsersLineURI = Get-CsOnlineUser -Filter {LineURI -ne $Null}
$UsersLineURI = Get-CsOnlineUser -Filter { EnterpriseVoiceEnabled -eq $true }
if ($UsersLineURI -ne $null) {
    foreach ($item in $UsersLineURI) {                  
        if ($onlyRA -and $Item.AccountType -ne 'ResourceAccount') {
            Continue
        }
        $myObject1 = New-Object System.Object
        
        $Matches = @()
        $Item.LineURI -match $Regex1 | out-null
        $phoneNumber = $Item.LineURI -replace "[^0-9,+]" , ''
        
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $phoneNumber
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
        $myObject1 | Add-Member -type NoteProperty -name "UPN" -Value $Item.UserPrincipalName
        $myObject1 | Add-Member -type NoteProperty -name "DisplayName" -Value $Item.DisplayName
        $myObject1 | Add-Member -type NoteProperty -name "FirstName" -Value $Item.GivenName
        $myObject1 | Add-Member -type NoteProperty -name "LastName" -Value $Item.LastName
        $myObject1 | Add-Member -type NoteProperty -name "Calling Policy" -Value $Item.TeamsCallingPolicy
        $myObject1 | Add-Member -type NoteProperty -name "Routing Policy" -Value $Item.OnlineVoiceRoutingPolicy
        $myObject1 | Add-Member -type NoteProperty -name "Dial Plan" -Value $Item.TenantDialPlan
        
        if ($Item.AccountType -eq 'ResourceAccount') {
            $applicationInstance = Get-CsOnlineApplicationInstance $Item.UserPrincipalName
            $myObject1 | Add-Member -type NoteProperty -name "Type" -Value $(if ($applicationInstance.ApplicationId -eq "ce933385-9390-45d1-9512-c8d228074e07") { "Auto Attendant Resource Account" } elseif ($applicationInstance.ApplicationId -eq "11cd3e2e-fccb-42ad-ad00-878b93575e07") { "Call Queue Resource Account" } elseif ($applicationInstance.ApplicationId -eq "01b9161a-881b-4ab0-8ee2-15e9141e95c6") { "PeterConnects Resource Account" } elseif ($applicationInstance.ApplicationId -eq "c8db29b6-8184-44fa-a6a1-086b8ae0435e") { "Roger365 Resource Account" } else { "Unknown Resource Account" })
            $myObject1 | Add-Member -type NoteProperty -name "ID" -Value $applicationInstance.ObjectId
        }
        else {
            $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "User"
            $myObject1 | Add-Member -type NoteProperty -name "ID" -Value ''
        }
        
        $Array1 += $myObject1          
    }
}

$unassignedNumbers = Get-CsTeamsUnassignedNumberTreatment
if ($unassignedNumbers -ne $null) {
    foreach ($unassignedNumber in $unassignedNumbers) {                  
        $myObject1 = New-Object System.Object
        
        $phoneNumber = $unassignedNumber.Pattern -replace "[^0-9,+]" , ''
        $user = (Get-CsOnlineUser $unassignedNumber.Target)
        
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $phoneNumber
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "UPN" -Value $user.UserPrincipalName
        $myObject1 | Add-Member -type NoteProperty -name "DisplayName" -Value $unassignedNumber.Identity
        $myObject1 | Add-Member -type NoteProperty -name "FirstName" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "LastName" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "Calling Policy" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "Routing Policy" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value 'UnassignedNumberTreatment'
        $myObject1 | Add-Member -type NoteProperty -name "ID" -Value ''
        
        $Array1 += $myObject1          
    }
}



if ($OutputType -eq "CSV") {
    $Array1 | export-csv $FilePath".csv" -NoTypeInformation
    Write-Host "ALL DONE!! Your file has been saved to $FilePath.csv"
}
elseif ($OutputType -eq "HTML") {
    $Header = @"

    <style>
    TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; width:95%;}
    TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
    TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
    </style>
"@
    
    $Array1 | Sort-Object -Property LineURI |  ConvertTo-Html -Head $Header | Out-File -FilePath $FilePath".html"
    Write-Host "ALL DONE!! Your file has been saved to $FilePath.html"
}
elseif ($OutputType -eq "CONSOLE") {
    $Array1 | FT -AutoSize -Property LineURI, DDI, Ext, DisplayName, UPN, Type
    Write-Host "ALL DONE!!"
}
else {
    $Array1 | FT -AutoSize -Property LineURI, DDI, Ext, DisplayName, UPN, Type
    Write-Host "WARNING: Valid output type not set, defaulted to console."
}