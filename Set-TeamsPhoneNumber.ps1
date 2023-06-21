<#
.SYNOPSIS
PowerShell script to assign Direct Route phone numbers and Voice Policies to users
By Ruud van Strijp - Axians
ruud.vanstrijp@axians.com

.NOTES
Microsoft Teams module 4.0.0 or higher is highly recommended, things might break if you run an older version
Method 1 (Recommended)
Update-Module MicrosoftTeams

Method 2 (Alternative)
Uninstall-Module -Name MicrosoftTeams -AllVersions
Install-Module -Name MicrosoftTeams -Force -Scope AllUsers

.EXAMPLE
.\Set-TeamsPhoneNumber.ps1 -identity <upn> -phoneNumber <phoneNumber> -voiceRoutingPolicy <voiceRoutingPolicy> (all optional)
If upn and phoneNumber are left empty, they will be requested
If voiceRoutingPolicy is left empty, all existing policies will be queried and a selection can be made
If teamsCallingPolicy is left empty, a check will be done to see if StandardUser exists. If not, all existing policies will be queried and a selection can be made
If upn does not contain an @, the script assumes the user's Display Name is entered and it will try to look up the corresponding upn

.EXAMPLE
.\Set-TeamsPhoneNumber
.\Set-TeamsPhoneNumber -identity firstname.lastname@domain.com -phoneNumber +31123456789
.\Set-TeamsPhoneNumber -identity firstname.lastname@domain.com -phoneNumber +31123456789 -voiceRoutingPolicy NL-International
.\Set-TeamsPhoneNumber -identity "fistname lastname" -phoneNumber +31123456789 -voiceRoutingPolicy NL-International
.\Set-TeamsPhoneNumber -identity firstname.lastname@domain.com -phoneNumber +31123456789 -teamsCallingPolicy AllowCalling

#>

Param (
[Parameter (Mandatory = $true)][string]$identity,
[Parameter (Mandatory = $true)][AllowEmptyString()][string]$phoneNumber,
[Parameter (Mandatory = $false)][string]$voiceRoutingPolicy,
[Parameter (Mandatory = $false)][string]$teamsDialPlan,
[Parameter (Mandatory = $false)][string]$teamsCallingPolicy,
[Parameter (Mandatory = $false)][string]$voicemailLanguage
)

$ProgressPreference='SilentlyContinue'

$debug = $true
 
$teamsModuleVersion = (Get-InstalledModule -Name MicrosoftTeams).Version
if($teamsModuleVersion -lt 4.0.0){
    Write-Host "  WARNING: Module Version older than 4.0.0 will be deprecated soon. This script might not run well" -ForegroundColor red
}

try {
    $null = Get-CsTenant
} catch { 
    Write-Host "  Currently not connected to Teams, connecting" -ForegroundColor yellow
    Connect-MicrosoftTeams 
}

Write-Host "  Connected to tenant: " -ForegroundColor White -NoNewLine
Write-Host (Get-CsTenant).DisplayName -ForegroundColor Green

#Correct User
if($identity -notmatch "\@"){
    Write-Host "  Not a UPN: "-ForegroundColor yellow -NoNewline
    Write-Host "$($identity)" -ForegroundColor green -NoNewline
    Write-Host ", trying to look up UPN" -ForegroundColor yellow
    
    $foundUsers = Get-CsOnlineUser | Where-Object DisplayName -like "*$identity*" | Select-Object DisplayName,UserPrincipalName
    
    if($foundUsers.count -eq 0 ){
        Write-Host "  ERROR: Name not found" -ForegroundColor Red
        exit
    }else{

        Write-Host "================ Please select the User ================"
        $i=0
        foreach ($foundUser in $foundUsers) {
            $i++
            Write-Host "$i : Press $i for $($foundUser.DisplayName)" -NoNewline
            Write-Host " ($($foundUser.UserPrincipalName))" -ForegroundColor green
        }
    
        $choice = Read-Host "Make a choice"
    
        if ($choice -gt 0 -and $choice -le $foundUsers.count) {
                $foundUser = $foundUsers[$choice-1]
                $upn = $foundUser.UserPrincipalName
            }
        else {
            Write-Host "Invalid selection" -ForegroundColor red
            exit
        }
        
    }

    
    Write-Host "  Found UPN: "-ForegroundColor yellow -NoNewline
    Write-Host "$($upn)" -ForegroundColor green
    Write-Host ""
}else{
    $upn = $identity
}

if($phoneNumber -eq $null -or $phoneNumber -eq ""){
    $user = Get-CsOnlineUser $upn | Select-Object DisplayName,LineURI,OnlineVoiceRoutingPolicy,TeamsCallingPolicy,DialPlan

    Write-Host ""
    Write-Host "Phone Number variable was left empty" -ForegroundColor red

    Write-Host " Selected user: " -ForegroundColor White -NoNewLine
    Write-Host "$($upn)" -ForegroundColor Green
    Write-Host " With Phone Number: " -ForegroundColor White -NoNewLine
    Write-Host "$($user.LineURI)" -ForegroundColor Green
    Write-Host " Voice Routing Policy: " -ForegroundColor White -NoNewLine
    Write-Host "$($user.OnlineVoiceRoutingPolicy)" -ForegroundColor Green
    Write-Host " Dial Plan: " -ForegroundColor White -NoNewLine
    Write-Host "$($user.DialPlan)" -ForegroundColor Green
    Write-Host " and Calling Policy: " -ForegroundColor White -NoNewLine
    Write-Host "$($user.TeamsCallingPolicy)" -ForegroundColor Green

    Write-Host "Do you want to remove the Phone Number and above policies from the user?" -ForegroundColor red
    
    $title    = ''
    $question = ''
    $choices  = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $null
        Grant-CsTenantDialPlan -Identity $upn -PolicyName $null
        Grant-CsTeamsCallingPolicy -Identity $upn -PolicyName $null
        Remove-CsPhoneNumberAssignment -Identity $upn -RemoveAll

        Write-Host "Result:" -ForegroundColor white
        Get-CsOnlineUser $upn | Select-Object DisplayName,LineURI,OnlineVoiceRoutingPolicy,EnterpriseVoiceEnabled,TeamsCallingPolicy,teamsupgrade*
        exit
    } else {
        Write-Host 'Cancelled'
        exit
    }
}

Write-Host ""

$voiceRoutingPolicies = Get-CsOnlineVoiceRoutingPolicy  | ForEach-Object {($_.Identity -replace "Tag:")}
if($voiceRoutingPolicy -eq $null -or $voiceRoutingPolicy -eq ""){
    Write-Host "================ Please select the Voice Routing Policy ================"

    $i=0
    foreach ($voiceRoutingPolicy in $voiceRoutingPolicies) {
        $i++
        Write-Host "$i : Press $i for $voiceRoutingPolicy"
    }

    $choice = Read-Host "Make a choice"

    if ($choice -gt 0 -and $choice -le $voiceRoutingPolicies.count) {
            $voiceRoutingPolicy = $voiceRoutingPolicies[$choice-1]
            #Write-Host "  Chosen Voice Routing Policy is: " -ForegroundColor White -NoNewline
            #Write-Host "$($voiceRoutingPolicy)" -ForegroundColor Green
        }
    else {
        Write-Host "Invalid selection" -ForegroundColor red
        exit
    }

}
elseif($voiceRoutingPolicy -notin $voiceRoutingPolicies){
    Write-Host "Specified Voice Routing Policy does not exist" -ForegroundColor red
    exit
}

Write-Host ""
$teamsCallingPolicies = Get-CsTeamsCallingPolicy  | ForEach-Object {($_.Identity -replace "Tag:")}
if($teamsCallingPolicy -eq $null -or $teamsCallingPolicy -eq ""){
    Write-Host "================ " -NoNewLine
    Write-Host "Please select the Teams Calling Policy ================"

    $i=0
    foreach ($teamsCallingPolicy in $teamsCallingPolicies) {
        $i++
        Write-Host "$i : Press $i for $teamsCallingPolicy"
    }

    $choice = Read-Host "Make a choice"

    if ($choice -gt 0 -and $choice -le $teamsCallingPolicies.count) {
            $teamsCallingPolicy = $teamsCallingPolicies[$choice-1]
            #Write-Host "  Chosen Voice Routing Policy is: " -ForegroundColor White -NoNewline
            #Write-Host "$($teamsCallingPolicy)" -ForegroundColor Green
        }
    else {
        Write-Host "Invalid selection" -ForegroundColor red
        exit
    }

}
elseif($teamsCallingPolicy -notin $teamsCallingPolicies){
    Write-Host "Specified Teams Calling Policy does not exist" -ForegroundColor red
    exit
}

Write-Host ""
$teamsDialPlans = Get-CsTenantDialPlan  | ForEach-Object {($_.Identity -replace "Tag:")}
if($teamsDialPlan -eq $null -or $teamsDialPlan -eq ""){
    Write-Host "================ " -NoNewLine
    Write-Host "Please select the Teams Dial Plan ================"

    $i=0
    foreach ($teamsDialPlan in $teamsDialPlans) {
        $i++
        Write-Host "$i : Press $i for $teamsDialPlan"
    }

    $choice = Read-Host "Make a choice"

    if ($choice -gt 0 -and $choice -le $teamsDialPlans.count) {
            $teamsDialPlan = $teamsDialPlans[$choice-1]
        }
    else {
        Write-Host "Invalid selection" -ForegroundColor red
        exit
    }

}
elseif($teamsDialPlan -notin $teamsDialPlans){
    Write-Host "Specified Teams Dial Plan does not exist" -ForegroundColor red
    exit
}

$checkUserExists = Get-CsOnlineUser $upn  -ErrorAction SilentlyContinue
if($null -eq $checkUserExists){
    Write-Host "  ERROR: user with UPN " -ForegroundColor Red -NoNewLine
    Write-Host "$($upn)" -ForegroundColor green -NoNewline
    Write-Host " not found" -ForegroundColor Red
    exit
}

#Check if the number is already assigned to another user
$filterString = 'LineURI -like "{0}"' -f $phoneNumber
$getLineUri = Get-CsOnlineUser -Filter $filterString | Select-Object DisplayName,UserPrincipalName

if($getLineUri -and $getLineUri.UserPrincipalName -ne $upn){
    Write-Host "  ERROR: Number already assigned to user: " -ForegroundColor Red -NoNewLine
    Write-Host "$($getLineUri.DisplayName)" -ForegroundColor Green -NoNewline
    Write-Host " with UPN " -ForegroundColor Red -NoNewLine
    Write-Host "$($getLineUri.UserPrincipalName)" -ForegroundColor Green
    exit
}

if($phoneNumber -like "tel:*"){
    $phoneNumber = $phoneNumber -replace "tel:"
    Write-Host "  DEBUG: Tel: is no longer required. Removed tel:" -ForegroundColor DarkGray
}

if($phoneNumber -like "+*"){
    $phoneNumber = $phoneNumber
}
else{
    $phoneNumber = "+"+$phoneNumber
}

Write-Host ""
Write-Host "Updating user: " -ForegroundColor White -NoNewLine
Write-Host "$($upn)" -ForegroundColor Green -NoNewLine
Write-Host " with " -ForegroundColor White -NoNewLine
Write-Host "$($phoneNumber)" -ForegroundColor Green -NoNewline
Write-Host " Voice Routing Policy " -ForegroundColor White -NoNewLine
Write-Host "$($voiceRoutingPolicy)" -ForegroundColor Green -NoNewLine
Write-Host " Dial Plan " -ForegroundColor White -NoNewLine
Write-Host "$($teamsDialPlan)" -ForegroundColor Green -NoNewLine
Write-Host " and Calling Policy " -ForegroundColor White -NoNewLine
Write-Host "$($teamsCallingPolicy)" -ForegroundColor Green

$title    = ''
$question = 'Do you want to continue?'
$choices  = '&Yes', '&No'

$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
if ($decision -eq 0) {
    
} else {
    Write-Host 'Cancelled'
    exit
}

#Enable user and assign phone number
if($debug -like $true){
    Write-Host "  DEBUG: Attempting to set Teams settings: Enabling Telephony Features and Configure Phone Number" -ForegroundColor DarkGray
}
try{
    #Set-CsUser -Identity $upn -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -LineURI $telLineURI
    Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $phoneNumber -PhoneNumberType DirectRouting
}
Catch{
    $errOutput = [PSCustomObject]@{
        status = "failed"
        error = $_.Exception.Message
        step = "SetCsPhoneNumberAssignment"
        cmdlet = "Set-CsPhoneNumberAssignment"
    }
    Write-Output ( $errOutput | ConvertTo-Json)
    exit
}


#Switch User to Teams Only
if($debug -like $true){
    Write-Host "  DEBUG: Attempting to switch user to Teams Only" -ForegroundColor DarkGray
}
try{
    Grant-CsTeamsUpgradePolicy $upn -PolicyName UpgradeToTeams
}
Catch{
    $errOutput = [PSCustomObject]@{
        status = "failed"
        error = $_.Exception.Message
        step = "TeamsUpgradePolicy"
        cmdlet = "Grant-CsTeamsUpgradePolicy"
    }
    Write-Output ( $errOutput | ConvertTo-Json)
    exit
}

if($debug -like $true){
    Write-Host "  DEBUG: Attempting to grant Teams Calling Policy: " -ForegroundColor DarkGray -NoNewLine
    Write-Host "$($teamsCallingPolicy)" -ForegroundColor Green
}
try{
    Grant-CsTeamsCallingPolicy -PolicyName $teamsCallingPolicy -Identity $upn
}
Catch{
    $errOutput = [PSCustomObject]@{
        status = "failed"
        error = $_.Exception.Message
        step = "TeamsCallingPolicy"
        cmdlet = "Grant-CsTeamsCallingPolicy"
    }
    Write-Output ( $errOutput | ConvertTo-Json)
    exit
}

#Assign Voice Routing Policy
if($debug -like $true){
    Write-Host "  DEBUG: Attempting to grant Online Voice Routing Policy: " -ForegroundColor DarkGray -NoNewLine
    Write-Host "$($voiceRoutingPolicy)" -ForegroundColor Green
}

if($voiceRoutingPolicy -eq "Global"){
    $voiceRoutingPolicy = $null
}

try{
    Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $voiceRoutingPolicy
}
Catch{
    $errOutput = [PSCustomObject]@{
        status = "failed"
        error = $_.Exception.Message
        step = "VoiceRoutingPolicy"
        cmdlet = "Grant-CsOnlineVoiceRoutingPolicy"
    }
    Write-Output ( $errOutput | ConvertTo-Json)
    exit
}

#Assign Dial Plan
if($debug -like $true){
    Write-Host "  DEBUG: Attempting to grant Tenant Dial Plan: " -ForegroundColor DarkGray -NoNewLine
    Write-Host "$($teamsDialPlan)" -ForegroundColor Green
}

try{
    Grant-CsTenantDialPlan -Identity $upn -PolicyName $teamsDialPlan
}
Catch{
    $errOutput = [PSCustomObject]@{
        status = "failed"
        error = $_.Exception.Message
        step = "TenantDialPlan"
        cmdlet = "Grant-CsTenantDialPlan"
    }
    Write-Output ( $errOutput | ConvertTo-Json)
    exit
}

Write-Host "Result:" -ForegroundColor white
Get-CsOnlineUser $upn | Select-Object DisplayName,LineURI,OnlineVoiceRoutingPolicy,EnterpriseVoiceEnabled,TenantDialPlan,TeamsCallingPolicy,teamsupgrade*
Write-Host "Warning: Voice Routing Policy might take some time to update" -ForegroundColor yellow