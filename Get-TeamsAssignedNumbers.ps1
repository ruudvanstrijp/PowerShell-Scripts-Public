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
    v1.7 - Added Unassigned Number Treatment
    v1.8 - Changed export to Output folder
    v1.9 - Added option to only show Resource Accounts and only show missing policies
    v2.0 - Added option to use extensions, added more resource account types, added ID column for resource accounts
    v2.1 - Fixed export to CSV

.NOTES
Microsoft Teams module 4.0.0 or higher needs to be installed into PowerShell. 5.0 is heavily recommended because of it's speed
Uninstall-Module -Name MicrosoftTeams -AllVersions
Install-Module -Name MicrosoftTeams -Force -Scope AllUsers
#>

Param (
    [switch]$onlyRA,
    [switch]$onlyMissingPolicies,
    [switch]$useExtensions,
    [ValidateSet("HTML", "CSV")]
    [string]$OutputType = "HTML"
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
    $tenantInfo = Get-CsTenant
}
Catch {
    Write-Host "  DEBUG: Could not connect to existing session, starting new session" -ForegroundColor DarkGray
    Connect-MicrosoftTeams
    $tenantInfo = Get-CsTenant
}

#Get Tenant onmicrosoft domain
$onMicrosoftDomainName = ($tenantInfo | Select-Object -ExpandProperty VerifiedDomains | Where-Object { $_.Name -like '*.onmicrosoft.com' -and $_.Name -notlike '*.mail.onmicrosoft.com' } | Select-Object -First 1).Name
$tenantName = $onMicrosoftDomainName -replace ".onmicrosoft.com", ""

Write-Host "  Connected to tenant: " -ForegroundColor White -NoNewLine
Write-Host ($tenantInfo).DisplayName -ForegroundColor Green
Write-Host "  With tenant domain: " -ForegroundColor White -NoNewLine
Write-Host $tenantName -ForegroundColor Green -NoNewLine
Write-Host ".onmicrosoft.com" -ForegroundColor White

#Settings ##############################
#. "_Settings.ps1" | Out-Null
$FileName = "TeamsAssignedNumbers_" + $tenantName + "_" + (Get-Date -Format s).replace(":", "-") 
$FolderPath = $PSScriptRoot + "\Output\"
$FilePath = $FolderPath + $FileName

#Check if FilePath Path exists, if not create it
if (!(Test-Path $FolderPath)) {
    New-Item -ItemType Directory -Force -Path $FolderPath
}


##############################

$Regex1 = '^(?:tel:)?(?:\+)?(\d+)(?:;ext=(\d+))?(?:;([\w-]+))?$'
$Array1 = @()
$userCount = $null
#Get Users with LineURI
#$UsersLineURI = Get-CsOnlineUser -Filter {LineURI -ne $Null}

if($onlyMissingPolicies){
    $UsersLineURI = Get-CsOnlineUser -Filter { EnterpriseVoiceEnabled -eq $true -and (TeamsCallingPolicy -eq $null -or OnlineVoiceRoutingPolicy -eq $null -or TenantDialPlan -eq $null)}
}else{
    $UsersLineURI = Get-CsOnlineUser -Filter { EnterpriseVoiceEnabled -eq $true }
    #Optie: Get-CsOnlineUser -Filter {FeatureTypes -contains "PhoneSystem"}
}

#$usersLineURI | Select-Object UserPrincipalName, TeamsCallingPolicy, OnlineVoiceRoutingPolicy, TenantDialPlan
$getApplications = Get-CsOnlineApplicationInstance
Write-Host "  DEBUG: Loaded user list. Processing data." -ForegroundColor DarkGray

if ($UsersLineURI) {
    foreach ($item in $UsersLineURI) {                  
        <# WIP
        Write-Host "  Querying policy information for" $item.DisplayName -ForegroundColor Green
        $UserPolicies  = Get-CsUserPolicyAssignment -Identity $item.UserPrincipalName -ErrorAction SilentlyContinue
        
        https://practical365.com/teams-policy-assignment-report/
        https://techcommunity.microsoft.com/discussions/microsoftteams/powershell-script-to-find-out-teams-policies-by-users/1210021
        
        $item.VoicePolicy = ($UserPolicies | Where-Object {$_.PolicyType -eq "VoicePolicy"}).PolicyName
        $item.VoicePolicy = (($UserPolicies | Where-Object {$_.PolicyType -eq "VoicePolicy"}).PolicySource).AssignmentType
        $item.MeetingPolicy = ($UserPolicies | Where-Object {$_.PolicyType -eq "MeetingPolicy"}).PolicyName
        $item.MeetingPolicySource = (($UserPolicies | Where-Object {$_.PolicyType -eq "MeetingPolicy"}).PolicySource).AssignmentType
        $item.TeamsMeetingPolicy = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsMeetingPolicy"}).PolicyName
        $item.TeamsMeetingPolicySource = (($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsMeetingPolicy"}).PolicySource).AssignmentType
        $item.TeamsMessagingPolicy = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsMessagingPolicy"}).PolicyName
        $item.TeamsMessagingPolicySource = (($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsMessagingPolicy"}).PolicySource).AssignmentType
        $item.TeamsAppSetupPolicy = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsAppSetupPolicy"}).PolicyName
        $item.TeamsAppSetupPolicySource = (($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsAppSetupPolicy"}).PolicySource).AssignmentType
        $item.TeamsCallingPolicy = ($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsCallingPolicy"}).PolicyName
        $item.TeamsCallingPolicySource = (($UserPolicies | Where-Object {$_.PolicyType -eq "TeamsCallingPolicy"}).PolicySource).AssignmentType
        #>

        if ($onlyRA -and $Item.AccountType -ne 'ResourceAccount') {
            Continue
        }
        if ($onlyMissingPolicies -and $Item.AccountType -eq 'ResourceAccount' -and ($Item.OnlineVoiceRoutingPolicy -or $Item.TenantDialPlan) ) {
            Continue
        }
        $myObject1 = New-Object System.Object
        
        $lineUriMatch = [regex]::Match($Item.LineURI, $Regex1)
        $phoneNumber = $Item.LineURI -replace "[^0-9,+]" , ''
        
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $phoneNumber
        if($useExtensions){
            $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $lineUriMatch.Groups[1].Value
            $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $lineUriMatch.Groups[2].Value
        }

        $myObject1 | Add-Member -type NoteProperty -name "UPN" -Value $Item.UserPrincipalName
        $myObject1 | Add-Member -type NoteProperty -name "DisplayName" -Value $Item.DisplayName
        $myObject1 | Add-Member -type NoteProperty -name "FirstName" -Value $Item.GivenName
        $myObject1 | Add-Member -type NoteProperty -name "LastName" -Value $Item.LastName
        $myObject1 | Add-Member -type NoteProperty -name "Calling Policy" -Value $Item.TeamsCallingPolicy
        $myObject1 | Add-Member -type NoteProperty -name "Routing Policy" -Value $Item.OnlineVoiceRoutingPolicy
        $myObject1 | Add-Member -type NoteProperty -name "Dial Plan" -Value $Item.TenantDialPlan
        $myObject1 | Add-Member -type NoteProperty -name "Caller ID Policy" -Value $Item.CallingLineIdentity
        
        if ($Item.AccountType -eq 'ResourceAccount') {
            #$applicationInstance = Get-CsOnlineApplicationInstance $Item.UserPrincipalName
            $applicationInstance = ($getApplications | Where-Object { $_.UserPrincipalName -eq $Item.UserPrincipalName })
            $myObject1 | Add-Member -type NoteProperty -name "Type" -Value $(if ($applicationInstance.ApplicationId -eq "ce933385-9390-45d1-9512-c8d228074e07") { "Auto Attendant Resource Account" } elseif ($applicationInstance.ApplicationId -eq "11cd3e2e-fccb-42ad-ad00-878b93575e07") { "Call Queue Resource Account" } elseif ($applicationInstance.ApplicationId -eq "01b9161a-881b-4ab0-8ee2-15e9141e95c6") { "PeterConnects Resource Account" } elseif ($applicationInstance.ApplicationId -eq "c8db29b6-8184-44fa-a6a1-086b8ae0435e") { "Roger365 Resource Account" } elseif ($applicationInstance.ApplicationId -eq "0346b13d-1bb8-4e22-9890-af279449eba9") { "Connecsy Resource Account" } else { "Unknown Resource Account" })
            $myObject1 | Add-Member -type NoteProperty -name "ID" -Value $applicationInstance.ObjectId
        }
        else {
            $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "User"
            $myObject1 | Add-Member -type NoteProperty -name "ID" -Value ''
			$userCount++
        }
        
        $Array1 += $myObject1          
    }
}

Write-Host "  Amount of Teams Voice users : " -ForegroundColor White -NoNewLine
Write-Host $userCount -ForegroundColor Green

$unassignedNumbers = Get-CsTeamsUnassignedNumberTreatment
if ($unassignedNumbers -and !$onlyMissingPolicies) {
    foreach ($unassignedNumber in $unassignedNumbers) {                  
        $myObject1 = New-Object System.Object
        
        $phoneNumber = $unassignedNumber.Pattern -replace "[^0-9,+]" , ''
        $user = (Get-CsOnlineUser $unassignedNumber.Target)
        
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $phoneNumber
        if($useExtensions){
            $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $unassignedNumber.Identity
            $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value ''
        }
        $myObject1 | Add-Member -type NoteProperty -name "UPN" -Value $user.UserPrincipalName
        $myObject1 | Add-Member -type NoteProperty -name "DisplayName" -Value $unassignedNumber.Description
        $myObject1 | Add-Member -type NoteProperty -name "FirstName" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "LastName" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "Calling Policy" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "Routing Policy" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "Dial Plan" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "Caller ID Policy" -Value ''
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "UnassignedNumberTreatment ($($unassignedNumber.TreatmentPriority))"
        $myObject1 | Add-Member -type NoteProperty -name "ID" -Value $unassignedNumber.Target
        
        $Array1 += $myObject1          
    }
}



if ($OutputType -eq "CSV") {
    $Array1 | Sort-Object -Property LineURI | Export-Csv -Path ($FilePath + ".csv") -NoTypeInformation
    Write-Host "ALL DONE!! Your file has been saved to $FilePath.csv"
}
elseif ($OutputType -eq "HTML") {
    $Header = '<style>
    body {
        background-color: white;
        font-family:      "Calibri";
    }

    table {
        border-width:     1px;
        border-style:     solid;
        border-color:     black;
        border-collapse:  collapse;
        width:            100%;
    }

    th {
        border-width:     1px;
        padding:          5px;
        border-style:     solid;
        border-color:     black;
        background-color: #98C6F3;
    }

    td {
        border-width:     1px;
        padding:          5px;
        border-style:     solid;
        border-color:     black;
        background-color: White;
    }

    tr {
        text-align:       left;
    }
    </style>'
    
    $Array1 | Sort-Object -Property LineURI |  ConvertTo-Html -Head $Header | Out-File -FilePath $FilePath".html"
    Write-Host "ALL DONE!! Your file has been saved to $FilePath.html"
}
else {
    throw "Unsupported OutputType '$OutputType'. Use HTML or CSV."
}