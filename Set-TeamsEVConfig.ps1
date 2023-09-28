<#
.SYNOPSIS
Based on UCDialPlans.com
https://www.ucdialplans.com

Stripped and modified by Ruud van Strijp - Axians
Added Gateway creation

.NOTES
Below settings are created:

Routing Policies:

Policy: COUNTRY-National
COUNTRY-Service
COUNTRY-National
COUNTRY-Mobile
COUNTRY-Free

Policy: COUNTRY-International
COUNTRY-Service
COUNTRY-National
COUNTRY-Mobile
COUNTRY-Free
COUNTRY-International
COUNTRY-Premium

#>

# $ErrorActionPreference can be set to SilentlyContinue, Continue, Stop, or Inquire for troubleshooting purposes
$Error.Clear()
$ErrorActionPreference = 'SilentlyContinue'

$sleepTimer = 15

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


#Create dialplans

function CreateDialPlan {
    param($DPname)
    
    write-host "" 
    write-host "----------------------------------------------------------------------------------"
    write-host "Select the Dial Plan Country"
    write-host "----------------------------------------------------------------------------------"
    write-host ""
    Write-Host "(1) Netherlands"
    Write-Host "(2) Belgium"
    Write-host "(3) Germany"
    Write-host "(4) France"
    Write-host "(5) Italy"
    Write-host "(6) Greece"
    Write-host "(7) Hungary"
    Write-host ""
    Write-Host "Press 'q' to quit."

    $input1 = Read-Host "Enter your choice"
    switch ($input1) {
        1 { 
            $country = "NL"
            
            Write-Host "Creating normalization rules for country " -NoNewline
            Write-Host "Netherlands" -ForegroundColor green
            $NR = @()
            $NR += New-CsVoiceNormalizationRule -Name "$country-Free" -Parent $DPname -Pattern '^0(800\d{4,7})\d*$' -Translation '+31$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Premium" -Parent $DPname -Pattern '^0(90\d{5,8}|8[47]\d{7})$' -Translation '+31$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Mobile" -Parent $DPname -Pattern '^0((6\d{8}))$' -Translation '+31$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-National" -Parent $DPname -Pattern '^0(([1-57]\d{4,8}|8[58]\d{7}))\d*(\D+\d+)?$' -Translation '+31$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Service" -Parent $DPname -Pattern '^(112|144|140\d{2,3}|116\d{3}|18\d{2})$' -Translation '$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-International" -Parent $DPname -Pattern '^(?:\+|00)(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{6,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory 
        } 
        2 { 
            $country = "BE"
            
            Write-Host "Creating normalization rules for country " -NoNewline
            Write-Host "Belgium" -ForegroundColor green
            $NR = @()
            $NR += New-CsVoiceNormalizationRule -Name "$country-Free" -Parent $DPname -Pattern '^0(800\d{6})\d*$' -Translation '+32$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Premium" -Parent $DPname -Pattern '^0(90\d{7})$' -Translation '+32$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Mobile" -Parent $DPname -Pattern '^0((4[1-9]\d{7}))$' -Translation '+32$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-National" -Parent $DPname -Pattern '^0(([1-8]\d{7}|9[1-9]\d{6}))\d*(\D+\d+)?$' -Translation '+32$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Service" -Parent $DPname -Pattern '^(1\d{2,3}|116\d{3})$' -Translation '$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-International" -Parent $DPname -Pattern '^(?:\+|00)(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{5,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory
        }
        3 { 
            $country = "DE"
            
            Write-Host "Creating normalization rules for country " -NoNewline
            Write-Host "Germany" -ForegroundColor green
            $NR = @()
            $NR += New-CsVoiceNormalizationRule -Name "$country-Free" -Parent $DPname -Pattern '^0(800\d{7,12})\d*$' -Translation '+49$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Premium" -Parent $DPname -Pattern '^0((900\d{7,8}|137\d{7}|138\d{4}))$' -Translation '+49$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Mobile" -Parent $DPname -Pattern '^0((1[567]\d{8,11}))$' -Translation '+49$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-National" -Parent $DPname -Pattern '^0((180[1-7]\d{6,7}|181\d{4,11}|18[2-9]\d{8}|700\d{8,11}|([2-7]\d{1,2}|80[2-9]|90[6-9]|[89][1-9]\d)\d{4,}))\d*(\D+\d+)?$' -Translation '+49$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Service" -Parent $DPname -Pattern -Pattern '^(11([025]|6\d{3}|8\d{2,3}))$' -Translation '$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-International" -Parent $DPname -Pattern '^(?:\+|00)(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{6,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory 
        }
        4 { 
            $country = "FR"
            
            Write-Host "Creating normalization rules for country " -NoNewline
            Write-Host "France" -ForegroundColor green
            $NR = @()
            $NR += New-CsVoiceNormalizationRule -Name "$country-Free" -Parent $DPname -Pattern '^0?(80\d{7})\d*$' -Translation '+33$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Premium" -Parent $DPname -Pattern '^0(8[129]\d{7})$' -Translation '+33$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Mobile" -Parent $DPname -Pattern '^0?(([67]\d{8}))$' -Translation '+33$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-National" -Parent $DPname -Pattern '^0([1-59]\d{8})\d*(\D+\d+)?$' -Translation '+33$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Service" -Parent $DPname -Pattern '^(1\d{1,2}|11[68]\d{3}|10\d{2}|3\d{3})$' -Translation '$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-International" -Parent $DPname -Pattern '^(?:\+|00)(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{6,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory 
        }
        5 { 
            $country = "IT"
            
            Write-Host "Creating normalization rules for country " -NoNewline
            Write-Host "Italy" -ForegroundColor green
            $NR = @()
            $NR += New-CsVoiceNormalizationRule -Name "$country-Free" -Parent $DPname -Pattern '^(80\d{5,7}|40\d{0,12})\d*$' -Translation '+39$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Premium" -Parent $DPname -Pattern '^((1(44|66|99)|8[49]\d)\d{4,7})$' -Translation '+39$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Mobile" -Parent $DPname -Pattern '^0((3\d{8,9}))$' -Translation '+39$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-National" -Parent $DPname -Pattern '^((0[1-9]\d{4,9}|55\d{8}))\d*(\D+\d+)?$' -Translation '+39$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Service" -Parent $DPname -Pattern '^(11\d|15\d\d|116\d{3})$' -Translation '$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-International" -Parent $DPname -Pattern '^(?:\+|00)(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{6,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory 
        }
        6 { 
            $country = "GR"
            
            Write-Host "Creating normalization rules for country " -NoNewline
            Write-Host "Greece" -ForegroundColor green
            $NR = @()
            $NR += New-CsVoiceNormalizationRule -Name "$country-Free" -Parent $DPname -Pattern '^(80[01]\d{7})\d*$' -Translation '+30$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Premium" -Parent $DPname -Pattern '^(90[19]\d{7})$' -Translation '+30$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Mobile" -Parent $DPname -Pattern '^((69\d{8}))$' -Translation '+30$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-National" -Parent $DPname -Pattern '^(((2[1-9]|70)\d{8}))\d*(\D+\d+)?$' -Translation '+30$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Service" -Parent $DPname -Pattern '^(1\d{2,4})$' -Translation '$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-International" -Parent $DPname -Pattern '^(?:\+|00)(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{6,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory 
        }
        7 { 
            $country = "HU"
            
            Write-Host "Creating normalization rules for country " -NoNewline
            Write-Host "Hungary" -ForegroundColor green
            $NR = @()
            $NR += New-CsVoiceNormalizationRule -Name "$country-Free" -Parent $DPname -Pattern '^06(80\d{6})\d*$' -Translation '+36$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Premium" -Parent $DPname -Pattern '^06(6(81|9[01])\d{6})$' -Translation '+36$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Mobile" -Parent $DPname -Pattern '^06((([237]0|31)\d{7}))$' -Translation '+36$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-National" -Parent $DPname -Pattern '^06(([1-7]\d|[89][2-9])\d{6})\d*(\D+\d+)?$' -Translation '+36$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-Service" -Parent $DPname -Pattern '^(10[4,5,7]|112)$' -Translation '$1' -InMemory
            $NR += New-CsVoiceNormalizationRule -Name "$country-International" -Parent $DPname -Pattern '^(?:\+|00)(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{6,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory 
        }
        q {
            Write-Host "Pressed quit" -ForegroundColor Red
            break script
        }
        default {
            Write-Host "No valid choice made" -ForegroundColor Red
            break script
        }
    }


    #$dialplans = Get-CsTenantDialPlan | Select-Object Identity,NormalizationRules
    #if ($dialplans.NormalizationRules.Name -notcontains 'NL-Free'){
    #old place for commands
    #}

    Write-Host "Creating tenant dial plan with name " -NoNewline
    Write-Host "$DPname" -ForegroundColor green
    Set-CsTenantDialPlan -Identity $DPname -NormalizationRules @{add = $NR }

    askCreateDiskPlan
}

function askCreateDiskPlan {
    Write-Host
    Write-Host 'Create global dial plan, user-level dial plan, skip creation or end script?' -ForegroundColor yellow
    $Global = New-Object System.Management.Automation.Host.ChoiceDescription '&Global', 'Create global dial plan.'
    $User = New-Object System.Management.Automation.Host.ChoiceDescription '&User', 'Create user-level dial plan'
    $Skip = New-Object System.Management.Automation.Host.ChoiceDescription '&Skip', 'Skip/Finish dial plan creation'
    $End = New-Object System.Management.Automation.Host.ChoiceDescription '&End', 'End script; do not create other policies'
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($Global, $User, $Skip, $End)
    $message = ''
    $DialPlanSelect = $Host.UI.PromptForChoice($caption, $message, $choices, -1)
    Write-Host
    If ($DialPlanSelect -eq 1) {
        $DPname = Read-Host -Prompt 'Please input Dial Plan name'
        New-CsTenantDialPlan $DPname | Out-Null
        CreateDialPlan -DPname $DPname
    }
    ElseIf ($DialPlanSelect -eq 0) {
        CreateDialPlan -DPname Global
    }
    ElseIf ($DialPlanSelect -eq 3) {
        Write-Host 'Ending script'
        Break
    }
    Else {
        Write-Host 'Continuing'
    }
}

askCreateDiskPlan


# Need to create gateway first?
Write-Host "Current gateways:" -ForegroundColor Green
Get-CsOnlinePSTNGateway | Format-List Identity

function CreateGateway {

    $fqdn = Read-Host -Prompt 'Gateway FQDN'

    $port = Read-Host -Prompt ‘Gateway port [5061]’

    if ([string]::IsNullOrWhiteSpace($port)) {
        $port = ‘5061’
    }

    Write-Host "Creating gateway with FQDN '$fqdn' and port '$port'" 
    
    New-CsOnlinePSTNGateway -Identity $fqdn -SipSignalingPort $port -ForwardCallHistory $True -ForwardPai $True -MediaBypass $False -Enabled $True | Out-Null

    Write-Host "Current gateways:" -ForegroundColor Green
    Get-CsOnlinePSTNGateway | Format-List Identity

    $confirmation = Read-Host "Do you want to create another gateway? [y/n]"
    if ($confirmation -eq 'y') {
        CreateGateway
    }
    else {
        Write-Host "Skipping gateway"
    }
}

$confirmation = Read-Host "Do you want to create a gateway now [y/n]"
if ($confirmation -eq 'y') {
    CreateGateway
}
else {
    Write-Host "Skipping gateway"
}

function CreateVoiceRoutingPolicies{
    param(
        [Parameter (Mandatory = $true)] [String]$country,
        [Parameter (Mandatory = $true)] [String]$patternMobile,
        [Parameter (Mandatory = $true)] [String]$patternFree,
        [Parameter (Mandatory = $true)] [String]$patternPremium,
        [Parameter (Mandatory = $true)] [String]$patternNational,
        [Parameter (Mandatory = $true)] [String]$patternService,
        [Parameter (Mandatory = $true)] [String]$patternInternational
        )

    Write-Host "Using country code " -NoNewline
    Write-Host "$country" -ForegroundColor green

    #### Voice Routing Policies
    Write-Host 'Creating Voice Routing Policies' -ForegroundColor yellow
    New-CsOnlineVoiceRoutingPolicy "$country-National" -WarningAction:SilentlyContinue | Out-Null
    New-CsOnlineVoiceRoutingPolicy "$country-International" -WarningAction:SilentlyContinue | Out-Null
    Start-Sleep -s $sleepTimer

    #### PSTN Usages
    Write-Host 'Creating PSTN Usages' -ForegroundColor yellow
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add = "$country-Service" } -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add = "$country-National" } -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add = "$country-Free" } -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add = "$country-Mobile" } -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add = "$country-Premium" } -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add = "$country-International" } -WarningAction:SilentlyContinue | Out-Null
    Start-Sleep -s $sleepTimer

    #### Add PSTN Usages to Routing Policies
    Write-Host "Adding PSTN usages to voice routing policies" -ForegroundColor yellow
    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-National" -OnlinePstnUsages @{Add = "$country-Service" } | Out-Null
    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-National" -OnlinePstnUsages @{Add = "$country-National" } | Out-Null
    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-National" -OnlinePstnUsages @{Add = "$country-Free" } | Out-Null
    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-National" -OnlinePstnUsages @{Add = "$country-Mobile" } | Out-Null

    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-International" -OnlinePstnUsages @{Add = "$country-Service" } | Out-Null
    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-International" -OnlinePstnUsages @{Add = "$country-National" } | Out-Null
    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-International" -OnlinePstnUsages @{Add = "$country-Free" } | Out-Null
    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-International" -OnlinePstnUsages @{Add = "$country-Mobile" } | Out-Null
    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-International" -OnlinePstnUsages @{Add = "$country-Premium" } | Out-Null
    Set-CsOnlineVoiceRoutingPolicy -Identity "$country-International" -OnlinePstnUsages @{Add = "$country-International" } | Out-Null

    Write-Host "Creating Voice Routes" -ForegroundColor yellow
    New-CsOnlineVoiceRoute -Name "$country-Mobile" -OnlinePstnUsages "$country-Mobile" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern $patternMobile | Out-Null
    New-CsOnlineVoiceRoute -Name "$country-Free" -OnlinePstnUsages "$country-Free" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern $patternFree | Out-Null
    New-CsOnlineVoiceRoute -Name "$country-Premium" -OnlinePstnUsages "$country-Premium" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern $patternPremium | Out-Null
    New-CsOnlineVoiceRoute -Name "$country-National" -OnlinePstnUsages "$country-National" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern $patternNational | Out-Null
    New-CsOnlineVoiceRoute -Name "$country-Service" -OnlinePstnUsages "$country-Service" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern $patternService | Out-Null
    New-CsOnlineVoiceRoute -Name "$country-International" -OnlinePstnUsages "$country-International" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern $patternInternational | Out-Null
    Start-Sleep -s $sleepTimer

    <#
    $i = 0
    $total = 10
    while ($i -lt $total) {
        $i++
        Write-Progress -Activity "Creating Voice Routes" -Status "Please wait..." -PercentComplete ($i / $total * 100)
        Start-Sleep -s 1
    }
    #>
}


function CreateVoiceRoutes {
    param($DPname)
    

    # Check for existence of PSTN gateways and prompt to add PSTN usages/routes

    $PSTNGW = Get-CsOnlinePSTNGateway
    If (($PSTNGW.Identity -eq $null) -and ($PSTNGW.Count -eq 0)) {
        Write-Host
        Write-Host 'No PSTN gateway found. If you want to configure Direct Routing, you have to define at least one PSTN gateway Using the New-CsOnlinePSTNGateway command.' -ForegroundColor Yellow

        Exit  
    }

    If ($PSTNGW.Count -gt 1) {
        $PSTNGWList = @()
        Write-Host
        Write-Host "ID    PSTN Gateway"
        Write-Host "==    ============"
        For ($i = 0; $i -lt $PSTNGW.Count; $i++) {
            $a = $i + 1
            Write-Host ($a, $PSTNGW[$i].Identity) -Separator "     "
        }

        $Range = '(1-' + $PSTNGW.Count + ')'
        Write-Host
        $Select = Read-Host "Select a primary PSTN gateway to apply routes" $Range

        If (($Select -gt $PSTNGW.Count) -or ($Select -lt 1)) {
            Write-Host 'Invalid selection' -ForegroundColor Red
            Exit
        }
        Else {
            $PSTNGWList += $PSTNGW[$Select - 1]
        }

        $Select = Read-Host "OPTIONAL - Select a secondary PSTN gateway to apply routes (or 0 to skip)" $Range

        If (($Select -gt $PSTNGW.Count) -or ($Select -lt 0)) {
            Write-Host 'Invalid selection' -ForegroundColor Red
            Exit
        }
        ElseIf ($Select -gt 0) {
            $PSTNGWList += $PSTNGW[$Select - 1]
        }
    }
    Else {
        # There is only one PSTN gateway
        $PSTNGWList = Get-CSOnlinePSTNGateway
    }

    write-host "" 
    write-host "----------------------------------------------------------------------------------"
    write-host "Select the Voice Routing Policies and PSTN Usages Country"
    write-host "----------------------------------------------------------------------------------"
    write-host ""
    Write-Host "(1) Netherlands"
    Write-Host "(2) Belgium"
    Write-host "(3) Germany"
    Write-host "(4) France"
    Write-host "(5) Italy"
    Write-host "(6) Greece"
    Write-host "(7) Hungary"
    Write-host ""
    Write-Host "Press 'q' to quit."

    $input1 = Read-Host "Enter your choice"
    switch ($input1) {
        1 { 
            $country = "NL"

            Write-Host
            Write-Host "Creating Routing settings for country " -NoNewline
            Write-Host "Netherlands" -ForegroundColor green
            

            $functionParams = @{
                country = "$country"
                patternMobile = '^\+31(6\d{8})$'
                patternFree = '^\+31800\d{4,7}$'
                patternPremium = '^\+3190\d{5,8}|8[47]\d{7}$'
                patternNational = '^\+310?([1-57]\d{4,8}|8[58]\d{7})'
                patternService = '^\+?(112|144|140\d{2,3}|116\d{3}|18\d{2}|319008844)$'
                patternInternational = '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(31))([2-9]\d{6,14})))'
            }

            CreateVoiceRoutingPolicies @functionParams

            #$priority = ((Get-CsOnlineVoiceRoute).Priority | Measure-Object -maximum).maximum

        } 
        2 { 
            $country = "BE"

            Write-Host
            Write-Host "Creating Routing settings for country " -NoNewline
            Write-Host "Belgium" -ForegroundColor green

            $functionParams = @{
                country = "$country"
                patternMobile = '^\+32(4[1-9]\d{7})$'
                patternFree = '^\+32800\d{6}$'
                patternPremium = '^\+3290\d{7}$'
                patternNational = '^\+32([1-8]\d{7}|9[1-9]\d{6})'
                patternService = '^\+?(1\d{2,3}|116\d{3})$'
                patternInternational = '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(32))([2-9]\d{6,14})))'
            }

            CreateVoiceRoutingPolicies @functionParams

        }
        3 { 
            $country = "DE"
            
            Write-Host
            Write-Host "Creating Routing settings for country " -NoNewline
            Write-Host "Germany" -ForegroundColor green

            $functionParams = @{
                country = "$country"
                patternMobile = '^\+49(1[567]\d{8,11})$'
                patternFree = '^\+49800\d{7,12}$'
                patternPremium = '^\+49(900\d{7,8}|137\d{7}|138\d{4})$'
                patternNational = '^\+490?(180[1-7]\d{6,7}|181\d{4,11}|18[2-9]\d{8}|700\d{8,11}|([2-7]\d{1,2}|80[2-9]|90[6-9]|[89][1-9]\d)\d{4,})'
                patternService = '^\+?(11([025]|6\d{3}|8\d{2,3}))$'
                patternInternational = '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(49))([2-9]\d{6,14})))'
            }

            CreateVoiceRoutingPolicies @functionParams

        }
        4 { 
            $country = "FR"
            
            Write-Host
            Write-Host "Creating Routing settings for country " -NoNewline
            Write-Host "France" -ForegroundColor green

            $functionParams = @{
                country = "$country"
                patternMobile = '^\+33([67]\d{8})$'
                patternFree = '^\+3380\d{7}$'
                patternPremium = '^\+338[129]\d{7}$'
                patternNational = '^\+330?[1-59]\d{8}'
                patternService = '^\+?(1\d{1,2}|11[68]\d{3}|10\d{2}|3\d{3})$'
                patternInternational = '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(33))([2-9]\d{6,14})))'
            }

            CreateVoiceRoutingPolicies @functionParams

        }
        5 { 
            $country = "IT"
            
            Write-Host
            Write-Host "Creating Routing settings for country " -NoNewline
            Write-Host "Italy" -ForegroundColor green

            $functionParams = @{
                country = "$country"
                patternMobile = '^\+39(3\d{8,9})$'
                patternFree = '^\+3980\d{5,7}|40\d{0,12}$'
                patternPremium = '^\+39(1(44|66|99)|8[49]\d)\d{4,7}$'
                patternNational = '^\+39(0[1-9]\d{4,9}|55\d{8})'
                patternService = '^\+?(11\d|15\d\d|116\d{3})$'
                patternInternational = '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(39))([2-9]\d{6,14})))'
            }

            CreateVoiceRoutingPolicies @functionParams

        }
        6 { 
            $country = "GR"
            
            Write-Host
            Write-Host "Creating Routing settings for country " -NoNewline
            Write-Host "Greece" -ForegroundColor green

            $functionParams = @{
                country = "$country"
                patternMobile = '^\+30(69\d{8})$'
                patternFree = '^\+3080[01]\d{7}$'
                patternPremium = '^\+3090[19]\d{7}$'
                patternNational = '^\+30((2[1-9]|70)\d{8})'
                patternService = '^\+?(1\d{2,4})$'
                patternInternational = '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(30))([2-9]\d{6,14})))'
            }

            CreateVoiceRoutingPolicies @functionParams

        }
        7 { 
            $country = "HU"
            
            Write-Host
            Write-Host "Creating Routing settings for country " -NoNewline
            Write-Host "Hungary" -ForegroundColor green
            
            $functionParams = @{
                country = "$country"
                patternMobile = '^\+36(([237]0|31)\d{7})$'
                patternFree = '^\+3680\d{6}$'
                patternPremium = '^\+366(81|9[01])\d{6}$'
                patternNational = '^\+36([1-7]\d|[89][2-9])\d{6}'
                patternService = '^\+?(10[4,5,7]|112)$'
                patternInternational = '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(32))([2-9]\d{6,14})))'
            }

            CreateVoiceRoutingPolicies @functionParams

        }
        q {
            Write-Host "Pressed quit" -ForegroundColor Red
            break script
        }
        default {
            Write-Host "No valid choice made" -ForegroundColor Red
            break script
        }
    }

    askCreateVoiceRoutes
}

function askCreateVoiceRoutes {
    Write-Host
    Write-Host 'Create Voice Routing Policies and PSTN Usages, skip creation or end script?' -ForegroundColor yellow
    Write-Host 'This is usually only required once per country' -ForegroundColor Green
    $Create = New-Object System.Management.Automation.Host.ChoiceDescription '&Create', 'Create Routing'
    $Skip = New-Object System.Management.Automation.Host.ChoiceDescription '&Skip', 'Skip/Finish Routing creation'
    $End = New-Object System.Management.Automation.Host.ChoiceDescription '&End', 'End script; do not create other policies'
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($Create, $Skip, $End)
    $message = ''
    $VoiceRouteSelect = $Host.UI.PromptForChoice($caption, $message, $choices, -1)
    Write-Host
    If ($VoiceRouteSelect -eq 0) {
        CreateVoiceRoutes
    }
    ElseIf ($VoiceRouteSelect -eq 2) {
        Write-Host 'Ending script'
        Break
    }
    Else {
        Write-Host 'Continuing'
    }
}

askCreateVoiceRoutes





<#
Write-Host 'Creating outbound translation rules' -ForegroundColor yellow
$OutboundTeamsNumberTranslations = New-Object 'System.Collections.Generic.List[string]'
New-CsTeamsTranslationRule -Identity "NL-TeamsTranslationRule" -Pattern '^\+(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{6,14})(;ext=\d+)?$' -Translation '+$1$2' | Out-Null
$OutboundTeamsNumberTranslations.Add("NL-TeamsTranslationRule")
Start-Sleep -s $sleepTimer

Write-Host 'Adding translation rules to PSTN gateways' -ForegroundColor yellow
ForEach ($PSTNGW in $PSTNGWList) {
	Set-CsOnlinePSTNGateway -Identity $PSTNGW.Identity -OutboundTeamsNumberTranslationRules $OutboundTeamsNumberTranslations -ErrorAction SilentlyContinue
}
#>

Write-Host
Write-Host "Do you want to create Standard User Calling Policy?" -ForegroundColor yellow
    
$title = ''
$question = ''
$choices = '&Yes', '&No'

$createCAP = $Host.UI.PromptForChoice($title, $question, $choices, 1)
if ($createCAP -eq 0) {
    Write-Host 'Creating Calling Policy StandardUser' -ForegroundColor yellow
    New-CsTeamsCallingPolicy -Identity StandardUser | Out-Null
}
else {
    Write-Host 'Not creating Standard User configuration'  -ForegroundColor yellow
}

Write-Host
Write-Host "Do you want to create Common Area Phone Calling Policy and User Policy?" -ForegroundColor yellow
    
$title = ''
$question = ''
$choices = '&Yes', '&No'

$createCAP = $Host.UI.PromptForChoice($title, $question, $choices, 1)
if ($createCAP -eq 0) {
    Write-Host 'Creating CAP configuration'  -ForegroundColor yellow
    New-CsTeamsIPPhonePolicy -Identity CommonAreaPhone -Description "Common Area Phone User Policy" -SignInMode CommonAreaPhoneSignIn -SearchOnCommonAreaPhoneMode Enabled -AllowHomeScreen Disabled -AllowBetterTogether Disabled -AllowHotDesking $FALSE
    New-CsTeamsCallingPolicy -Identity CommonAreaPhone  -Description "Common Area Phone Policy" -AllowCallForwardingToPhone $false -AllowCallForwardingToUser $false -AllowCallGroups $false -AllowVoicemail AlwaysDisabled -AllowWebPSTNCalling $false -AllowPrivateCalling $true -AllowDelegation $false -LiveCaptionsEnabledTypeForCalling Disabled
}
else {
    Write-Host 'Not creating CAP configuration'  -ForegroundColor yellow
}

Write-Host 'Configuration complete' -ForegroundColor Yellow