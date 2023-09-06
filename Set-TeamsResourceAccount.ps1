<#
.SYNOPSIS
PowerShell script to assign Direct Routing phone numbers and Voice Policies to Resource Accounts
This script also supplies an easy way to switch resource account type

By Ruud van Strijp - Axians
ruud.vanstrijp@axians.com

#>

Param (
[Parameter (Mandatory = $false)][string]$upn,
[Parameter (Mandatory = $false)][string]$phoneNumber,
[Parameter (Mandatory = $false)][string]$voiceRoutingPolicy,
[Parameter (Mandatory = $false)][ValidateSet("AA","CQ","Roger365")]$type
)
 
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

$resourceAccounts = Get-CsOnlineApplicationInstance
if($upn -eq $null -or $upn -eq ""){
    Write-Host "================ Please select the Resource Account ================"

    $i=0
    foreach ($resourceAccount in $resourceAccounts) {
        $i++
        Write-Host "$i : Press $i for" $resourceAccount.DisplayName "(" -NoNewline
        Write-Host $resourceAccount.UserPrincipalName  -ForegroundColor green -NoNewline
        Write-Host ")"
    }

    $choice = Read-Host "Make a choice"

    $choice = [int]$choice

    if ($choice -gt 0 -and $choice -le $resourceAccounts.count) {
            $upn = $resourceAccounts[$choice-1].UserPrincipalName
        }
    else {
        Write-Host "Invalid selection" -ForegroundColor red
        exit
    }
}

Write-Host "Selected user: " -ForegroundColor White -NoNewLine
Write-Host "$($upn)" -ForegroundColor Green


#Correct User
if($upn -notmatch "\@"){
    Write-Host "  WARNING: Not a UPN: "-ForegroundColor yellow -NoNewline
    Write-Host "$($upn)" -ForegroundColor green -NoNewline
    exit
}

if($phoneNumber -eq $null -or $phoneNumber -eq ""){
    $title = ''
    $question = 'Do you want to assign a phone number to the user?'
    $choices = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        $phoneNumber = Read-Host -Prompt 'Input phone number'
        
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

        Write-Host "Updating user: " -ForegroundColor White -NoNewLine
        Write-Host "$($upn)" -ForegroundColor Green -NoNewLine
        Write-Host " with " -ForegroundColor White -NoNewLine
        Write-Host "$($phoneNumber)" -ForegroundColor Green

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
    }
    else {
        Write-Host '  Skipping phone number assignment' -ForegroundColor Yellow
    }
}
else{
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
}

if($voiceRoutingPolicy -eq $null -or $voiceRoutingPolicy -eq ""){
    Write-Host
    $title = ''
    $question = 'Do you want to assign a Voice Routing Policy to this Resource Account?'
    $choices = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
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

            #Assign Voice Routing Policy
        if($debug -like $true){
            Write-Host "  DEBUG: Attempting to grant Teams settings: Assign the Online Voice Routing Policy" -ForegroundColor DarkGray
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
    }
    else {
        Write-Host '  Skipping policy assignment' -ForegroundColor Yellow
    }
}
else{
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
}

if(!$type){
    Write-Host
    $title = ''
    $question = 'Do you want to change the Resource Account Type?'
    $choices = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        
        write-host "" 
        write-host "----------------------------------------------------------------------------------"
        write-host "Select the required Resource Account Type"
        write-host "----------------------------------------------------------------------------------"
        write-host ""
        Write-Host "(1) Auto Attendant"
        Write-Host "(2) Call Queue"
        Write-host "(3) Roger365"
        Write-host ""
        Write-Host "Press 'q' to quit."

        $input1 = Read-Host "Enter your choice"
        switch ($input1) {
            1 { 
                Set-CsOnlineApplicationInstance -ApplicationId "ce933385-9390-45d1-9512-c8d228074e07" -Identity $upn
                Write-host "Setting type to Auto Attendant" -ForegroundColor White
            } 
            2 { 
                Set-CsOnlineApplicationInstance -ApplicationId "11cd3e2e-fccb-42ad-ad00-878b93575e07" -Identity $upn
                Write-host "Setting type to Call Queue" -ForegroundColor White
            }
            3 { 
                Set-CsOnlineApplicationInstance -ApplicationId "c8db29b6-8184-44fa-a6a1-086b8ae0435e" -Identity $upn
                Write-host "Setting type to Roger365" -ForegroundColor White
            }
            q {
                Write-Host "Pressed quit" -ForegroundColor Red
                #break script
            }
            default {
                Write-Host "No valid choice made" -ForegroundColor Red
                #break script
            }
        }

    }
    else {
        Write-Host '  Skipping type change' -ForegroundColor Yellow
    }
}
elseif($type -eq 'AA'){
    Set-CsOnlineApplicationInstance -ApplicationId "ce933385-9390-45d1-9512-c8d228074e07" -Identity $upn
    Write-host "Setting type to Auto Attendant" -ForegroundColor White
}
elseif($type -eq 'CQ'){
    Set-CsOnlineApplicationInstance -ApplicationId "11cd3e2e-fccb-42ad-ad00-878b93575e07" -Identity $upn
    Write-host "Setting type to Call Queue" -ForegroundColor White
}
elseif($type -eq 'Roger365'){
    Set-CsOnlineApplicationInstance -ApplicationId "c8db29b6-8184-44fa-a6a1-086b8ae0435e" -Identity $upn
    Write-host "Setting type to Roger365" -ForegroundColor White
}

$objectID = (Get-CsOnlineApplicationInstance -Identity $upn).ObjectID
Sync-CsOnlineApplicationInstance -ObjectID $objectID

Write-Host "Resource Account ObjectID: " -ForegroundColor White -NoNewLine
Write-Host "$($objectID)" -ForegroundColor Green