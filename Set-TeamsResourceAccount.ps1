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
    [Parameter (Mandatory = $false)][string]$teamsDialPlan,
    [Parameter (Mandatory = $false)][ValidateSet("AA", "CQ", "Roger365")]$type
)
 
$debug = $true

$newAccountCreated = $false
 
$teamsModuleVersion = (Get-InstalledModule -Name MicrosoftTeams).Version
if ($teamsModuleVersion -lt 4.0.0) {
    Write-Host "  WARNING: Module Version older than 4.0.0 will be deprecated soon. This script might not run well" -ForegroundColor red
}

try {
    $null = Get-CsTenant
}
catch { 
    Write-Host "  Currently not connected to Teams, connecting" -ForegroundColor yellow
    Connect-MicrosoftTeams 
}

Write-Host "  Connected to tenant: " -ForegroundColor White -NoNewLine
Write-Host (Get-CsTenant).DisplayName -ForegroundColor Green

Write-Host
Write-Host 'Create new Resource Account, Modify existing Resource Account or End script?' -ForegroundColor yellow
$Create = New-Object System.Management.Automation.Host.ChoiceDescription '&Create', 'Create new resource account.'
$Modify = New-Object System.Management.Automation.Host.ChoiceDescription '&Modify', 'Modify existing resource account'
$End = New-Object System.Management.Automation.Host.ChoiceDescription '&End', 'End script; do not create other policies'
$choices = [System.Management.Automation.Host.ChoiceDescription[]]($Create, $Modify, $End)
$message = ''
$RASelect = $Host.UI.PromptForChoice($caption, $message, $choices, -1)
Write-Host
If ($RASelect -eq 1) {
    #continue
}
ElseIf ($RASelect -eq 0) {
    $upn = Read-Host -Prompt 'Enter UPN for the new Resource Account'
    $DisplayName = Read-Host -Prompt 'Enter Display Name for the new Resource Account'
    
    Write-Host
    Write-Host "Select the Resource Account Type:" -ForegroundColor Yellow
    Write-Host "(1) Auto Attendant"
    Write-Host "(2) Call Queue"
    Write-Host "(3) Roger365"
    Write-Host "(4) Connecsy"
    
    $typeChoice = Read-Host "Enter your choice (default is 1 for Auto Attendant)"
    
    # Default Application ID is for Auto Attendant
    $applicationId = "ce933385-9390-45d1-9512-c8d228074e07" # Auto Attendant
    
    switch ($typeChoice) {
        1 { 
            Write-Host "Creating Resource Account of Auto Attendant Type" -ForegroundColor White
            $applicationId = "ce933385-9390-45d1-9512-c8d228074e07" # Auto Attendant
        }
        2 { 
            Write-Host "Creating Resource Account of Call Queue Type" -ForegroundColor White
            $applicationId = "11cd3e2e-fccb-42ad-ad00-878b93575e07" # Call Queue
        }
        3 { 
            Write-Host "Creating Resource Account of Roger365 Type" -ForegroundColor White
            $applicationId = "c8db29b6-8184-44fa-a6a1-086b8ae0435e" # Roger365
        }
        4 { 
            Write-Host "Creating Resource Account of Connecsy Type" -ForegroundColor White
            $applicationId = "0346b13d-1bb8-4e22-9890-af279449eba9" # Connecsy
        }
        "" { 
            Write-Host "No selection made, defaulting to Auto Attendant Type" -ForegroundColor Yellow
            $applicationId = "ce933385-9390-45d1-9512-c8d228074e07" # Auto Attendant
        }
        default {
            Write-Host "Invalid selection, defaulting to Auto Attendant Type" -ForegroundColor Yellow
            $applicationId = "ce933385-9390-45d1-9512-c8d228074e07" # Auto Attendant
        }
    }
    
    # Create the resource account with the selected type
    New-CsOnlineApplicationInstance -UserPrincipalName $upn -DisplayName $DisplayName -ApplicationId $applicationId
    
    # Markeer dat we een nieuw account hebben aangemaakt
    $newAccountCreated = $true
    
    # Pause script for 1 minute
    Write-Host 'Pausing script for 1 minute to allow license to be assigned by dynamic group membership'
    Start-Sleep -s 60
}
ElseIf ($RASelect -eq 2) {
    Write-Host 'Ending script'
    Break
}
Else {
    Write-Host 'No valid choice was made. Ending script'
    Break
}

$resourceAccounts = Get-CsOnlineApplicationInstance
if ($null -eq $upn -or $upn -eq "") {
    Write-Host "================ Please select the Resource Account ================"

    $i = 0
    foreach ($resourceAccount in $resourceAccounts) {
        $i++
        Write-Host "$i : Press $i for" $resourceAccount.DisplayName "(" -NoNewline
        Write-Host $resourceAccount.UserPrincipalName  -ForegroundColor green -NoNewline
        Write-Host ")"
    }

    $choice = Read-Host "Make a choice"

    $choice = [int]$choice

    if ($choice -gt 0 -and $choice -le $resourceAccounts.count) {
        $upn = $resourceAccounts[$choice - 1].UserPrincipalName
    }
    else {
        Write-Host "Invalid selection" -ForegroundColor red
        exit
    }
}

Write-Host
Write-Host "Selected user: " -ForegroundColor White -NoNewLine
Write-Host "$($upn)" -ForegroundColor Green
Write-Host


#Correct User
if ($upn -notmatch "\@") {
    Write-Host "  WARNING: Not a UPN: "-ForegroundColor yellow -NoNewline
    Write-Host "$($upn)" -ForegroundColor green -NoNewline
    exit
}

$objectID = (Get-CsOnlineApplicationInstance -Identity $upn).ObjectID

if ($null -eq $phoneNumber -or $phoneNumber -eq "") {
    $title = ''
    $question = 'Assign a Phone Number, Remove a Phone Number or Skip Phone Number assignment?'
    $choices = '&Assign', '&Remove', '&Skip'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 2)
    if ($decision -eq 0) {
        $phoneNumber = Read-Host -Prompt 'Input phone number'
        
        #Check if the number is already assigned to another user
        $filterString = 'LineURI -like "{0}"' -f $phoneNumber
        $getLineUri = Get-CsOnlineUser -Filter $filterString | Select-Object DisplayName, UserPrincipalName

        if ($getLineUri -and $getLineUri.UserPrincipalName -ne $upn) {
            Write-Host "  ERROR: Number already assigned to user: " -ForegroundColor Red -NoNewLine
            Write-Host "$($getLineUri.DisplayName)" -ForegroundColor Green -NoNewline
            Write-Host " with UPN " -ForegroundColor Red -NoNewLine
            Write-Host "$($getLineUri.UserPrincipalName)" -ForegroundColor Green
            exit
        }

        if ($phoneNumber -like "tel:*") {
            $phoneNumber = $phoneNumber -replace "tel:"
            Write-Host "  DEBUG: Tel: is no longer required. Removed tel:" -ForegroundColor DarkGray
        }

        if ($phoneNumber -like "+*") {
            $phoneNumber = $phoneNumber
        }
        else {
            $phoneNumber = "+" + $phoneNumber
        }

        Write-Host "Updating user: " -ForegroundColor White -NoNewLine
        Write-Host "$($upn)" -ForegroundColor Green -NoNewLine
        Write-Host " with " -ForegroundColor White -NoNewLine
        Write-Host "$($phoneNumber)" -ForegroundColor Green

        try {
            #Set-CsUser -Identity $upn -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -LineURI $telLineURI
            Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $phoneNumber -PhoneNumberType DirectRouting
        }
        Catch {
            $errOutput = [PSCustomObject]@{
                status = "failed"
                error  = $_.Exception.Message
                step   = "SetCsPhoneNumberAssignment"
                cmdlet = "Set-CsPhoneNumberAssignment"
            }
            Write-Output ( $errOutput | ConvertTo-Json)
            exit
        }
    }if($decision -eq 1) {
        Write-Host '  Removing phone number assignment' -ForegroundColor Yellow

        try {  
            Remove-CsPhoneNumberAssignment -Identity $upn -RemoveAll
        }
        Catch {
            $errOutput = [PSCustomObject]@{
                status = "failed"
                error  = $_.Exception.Message
                step   = "SetCsPhoneNumberAssignment"
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
else {
    try {
        #Set-CsUser -Identity $upn -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -LineURI $telLineURI
        Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $phoneNumber -PhoneNumberType DirectRouting
    }
    Catch {
        $errOutput = [PSCustomObject]@{
            status = "failed"
            error  = $_.Exception.Message
            step   = "SetCsPhoneNumberAssignment"
            cmdlet = "Set-CsPhoneNumberAssignment"
        }
        Write-Output ( $errOutput | ConvertTo-Json)
        exit
    }
}

if ($null -eq $voiceRoutingPolicy -or $voiceRoutingPolicy -eq "") {
    Write-Host
    $title = ''
    $question = 'Do you want to assign a Voice Routing Policy to this Resource Account?'
    $choices = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        $voiceRoutingPolicies = Get-CsOnlineVoiceRoutingPolicy  | ForEach-Object { ($_.Identity -replace "Tag:") }
        if ($null -eq $voiceRoutingPolicy -or $voiceRoutingPolicy -eq "") {
            Write-Host "================ Please select the Voice Routing Policy ================"

            $i = 0
            foreach ($voiceRoutingPolicy in $voiceRoutingPolicies) {
                $i++
                Write-Host "$i : Press $i for $voiceRoutingPolicy"
            }

            $choice = Read-Host "Make a choice"
            $choice = [int]$choice

            if ($choice -gt 0 -and $choice -le $voiceRoutingPolicies.count) {
                $voiceRoutingPolicy = $voiceRoutingPolicies[$choice - 1]
                #Write-Host "  Chosen Voice Routing Policy is: " -ForegroundColor White -NoNewline
                #Write-Host "$($voiceRoutingPolicy)" -ForegroundColor Green
            }
            else {
                Write-Host "Invalid selection" -ForegroundColor red
                exit
            }

        }
        elseif ($voiceRoutingPolicy -notin $voiceRoutingPolicies) {
            Write-Host "Specified Voice Routing Policy does not exist" -ForegroundColor red
            exit
        }

        #Assign Voice Routing Policy
        if ($debug -like $true) {
            Write-Host "  DEBUG: Attempting to grant Teams settings: Assign the Online Voice Routing Policy" -ForegroundColor DarkGray
        }

        if ($voiceRoutingPolicy -eq "Global") {
            $voiceRoutingPolicy = $null
        }

        try {
            Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $voiceRoutingPolicy
        }
        Catch {
            $errOutput = [PSCustomObject]@{
                status = "failed"
                error  = $_.Exception.Message
                step   = "VoiceRoutingPolicy"
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
else {
    try {
        Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $voiceRoutingPolicy
    }
    Catch {
        $errOutput = [PSCustomObject]@{
            status = "failed"
            error  = $_.Exception.Message
            step   = "VoiceRoutingPolicy"
            cmdlet = "Grant-CsOnlineVoiceRoutingPolicy"
        }
        Write-Output ( $errOutput | ConvertTo-Json)
        exit
    }
}

$teamsDialPlans = Get-CsTenantDialPlan  | ForEach-Object { ($_.Identity -replace "Tag:") }
if ($null -eq $teamsDialPlan -or $teamsDialPlan -eq "") {
    Write-Host
    $title = ''
    $question = 'Do you want to assign a Dial Plan to this Resource Account?'
    $choices = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {    
            
        Write-Host "================ " -NoNewLine
        Write-Host "Please select the Teams Dial Plan ================"

        $i = 0
        foreach ($teamsDialPlan in $teamsDialPlans) {
            $i++
            Write-Host "$i : Press $i for $teamsDialPlan"
        }

        $choice = Read-Host "Make a choice"
        $choice = [int]$choice

        if ($choice -gt 0 -and $choice -le $teamsDialPlans.count) {
            $teamsDialPlan = $teamsDialPlans[$choice - 1]
        }
        else {
            Write-Host "Invalid selection" -ForegroundColor red
            exit
        }

        #Assign Dial Plan
        if ($debug -like $true) {
            Write-Host "  DEBUG: Attempting to grant Tenant Dial Plan: " -ForegroundColor DarkGray -NoNewLine
            Write-Host "$($teamsDialPlan)" -ForegroundColor Green
        } 

        try {
            Grant-CsTenantDialPlan -Identity $upn -PolicyName $teamsDialPlan
        }
        Catch {
            $errOutput = [PSCustomObject]@{
                status = "failed"
                error  = $_.Exception.Message
                step   = "VoiceRoutingPolicy"
                cmdlet = "Grant-CsTenantDialPlan"
            }
            Write-Output ( $errOutput | ConvertTo-Json)
            exit
        }
    }
    else {
        Write-Host '  Skipping policy assignment' -ForegroundColor Yellow
    }
}
elseif ($teamsDialPlan -notin $teamsDialPlans) {
    Write-Host "Specified Teams Dial Plan does not exist" -ForegroundColor red
    exit
}
else {
    try {
        Grant-CsTenantDialPlan -Identity $upn -PolicyName $teamsDialPlan
    }
    Catch {
        $errOutput = [PSCustomObject]@{
            status = "failed"
            error  = $_.Exception.Message
            step   = "TenantDialPlan"
            cmdlet = "Grant-CsTenantDialPlan"
        }
        Write-Output ( $errOutput | ConvertTo-Json)
        exit
    }
}


if (!$type -and !$newAccountCreated) {
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
        Write-host "(4) Connecsy"
        Write-host ""
        Write-Host "Press 'q' to quit."

        $input1 = Read-Host "Enter your choice"
        switch ($input1) {
            1 { 
                Write-host "Setting type to Auto Attendant" -ForegroundColor White
                Set-CsOnlineApplicationInstance -ApplicationId "ce933385-9390-45d1-9512-c8d228074e07" -Identity $upn
                Write-host "Syncing Application Instance" -ForegroundColor White
                Sync-CsOnlineApplicationInstance -ApplicationId "ce933385-9390-45d1-9512-c8d228074e07" -ObjectID $objectID
            } 
            2 { 
                Write-host "Setting type to Call Queue" -ForegroundColor White
                Set-CsOnlineApplicationInstance -ApplicationId "11cd3e2e-fccb-42ad-ad00-878b93575e07" -Identity $upn
                Write-host "Syncing Application Instance" -ForegroundColor White
                Sync-CsOnlineApplicationInstance -ApplicationId "11cd3e2e-fccb-42ad-ad00-878b93575e07" -ObjectID $objectID
            }
            3 { 
                Write-host "Setting type to Roger365" -ForegroundColor White
                Set-CsOnlineApplicationInstance -ApplicationId "c8db29b6-8184-44fa-a6a1-086b8ae0435e" -Identity $upn
                Write-host "Syncing Application Instance" -ForegroundColor White
                Sync-CsOnlineApplicationInstance -ApplicationId "c8db29b6-8184-44fa-a6a1-086b8ae0435e" -ObjectID $objectID
            }
            4 { 
                Write-host "Setting type to Connecsy" -ForegroundColor White
                Set-CsOnlineApplicationInstance -ApplicationId "0346b13d-1bb8-4e22-9890-af279449eba9" -Identity $upn
                Write-host "Syncing Application Instance" -ForegroundColor White
                Sync-CsOnlineApplicationInstance -ApplicationId "0346b13d-1bb8-4e22-9890-af279449eba9" -ObjectID $objectID
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
elseif (!$type -and $newAccountCreated) {
    Write-Host "  Skipping type change because the account was just created with the selected type" -ForegroundColor Yellow
}
elseif ($type -eq 'AA') {
    Write-host "Setting type to Auto Attendant" -ForegroundColor White
    Set-CsOnlineApplicationInstance -ApplicationId "ce933385-9390-45d1-9512-c8d228074e07" -Identity $upn
    Write-host "Syncing Application Instance" -ForegroundColor White
    Sync-CsOnlineApplicationInstance -ApplicationId "ce933385-9390-45d1-9512-c8d228074e07" -ObjectID $objectID
    
}
elseif ($type -eq 'CQ') {
    Write-host "Setting type to Call Queue" -ForegroundColor White
    Set-CsOnlineApplicationInstance -ApplicationId "11cd3e2e-fccb-42ad-ad00-878b93575e07" -Identity $upn
    Write-host "Syncing Application Instance" -ForegroundColor White
    Sync-CsOnlineApplicationInstance -ApplicationId "11cd3e2e-fccb-42ad-ad00-878b93575e07" -ObjectID $objectID
}
elseif ($type -eq 'Roger365') {
    Write-host "Setting type to Roger365" -ForegroundColor White
    Set-CsOnlineApplicationInstance -ApplicationId "c8db29b6-8184-44fa-a6a1-086b8ae0435e" -Identity $upn
    Write-host "Syncing Application Instance" -ForegroundColor White
    Sync-CsOnlineApplicationInstance -ApplicationId "c8db29b6-8184-44fa-a6a1-086b8ae0435e" -ObjectID $objectID
}
elseif ($type -eq 'Connecsy') {
    Write-host "Setting type to Connecsy" -ForegroundColor White
    Set-CsOnlineApplicationInstance -ApplicationId "0346b13d-1bb8-4e22-9890-af279449eba9" -Identity $upn
    Write-host "Syncing Application Instance" -ForegroundColor White
    Sync-CsOnlineApplicationInstance -ApplicationId "0346b13d-1bb8-4e22-9890-af279449eba9" -ObjectID $objectID
}



Write-Host "Resource Account ObjectID: " -ForegroundColor White -NoNewLine
Write-Host "$($objectID)" -ForegroundColor Green
Write-Host ""