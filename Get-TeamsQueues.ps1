<#
.SYNOPSIS
PowerShell script to bulk export Teams Resource Accounts, Auto Attendants and Call Queues

.DESCRIPTION
By Ruud van Strijp - Axians
ruud.vanstrijp@axians.com

.NOTES


.EXAMPLE
.\Get-TeamsQueues.ps1

#>

$debug = $true

<#
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

#>

#Settings ##############################
#. "_Settings.ps1" | Out-Null
$FileName = "TeamsQueues_" + (Get-Date -Format s).replace(":", "-") 
$FilePath = $PSScriptRoot + "\Output\" + $FileName
$OutputType = "HTML" #OPTIONS: CSV - Outputs CSV to specified FilePath, CONSOLE - Outputs to console


##############################

$Regex1 = '^(?:tel:)?(?:\+)?(\d+)(?:;ext=(\d+))?(?:;([\w-]+))?$'
$CallFlows = @()

$allUsers = Get-CsOnlineUser -Filter {LineURI -ne $null -and AccountEnabled -eq $True -and AccountType -eq 'User'}
$allResourceAccounts = Get-CsOnlineApplicationInstance
$VoiceAppAas = Get-CsAutoAttendant
$VoiceAppCqs = Get-CsCallQueue -WarningAction SilentlyContinue


if ($VoiceAppCqs -ne $null) {
    foreach ($VoiceAppCq in $VoiceAppCqs) {                  
        $VoiceAppCqCallFlow = New-Object System.Object
        

        #Resource Accounts Phone Numbers and UPN's
        $ApplicationInstanceAssociationCounter = 0
        $ResourceAccountPhoneNumbers = ""
        $ResourceAccountUPNs = ""

        foreach ($ResourceAccount in $VoiceAppCq.ApplicationInstances) {
            $ResourceAccountPhoneNumber = ($allResourceAccounts | Where-Object { $_.ObjectId -eq $ResourceAccount }).PhoneNumber
            if ($ResourceAccountPhoneNumber) {
                $ResourceAccountPhoneNumber = $ResourceAccountPhoneNumber.Replace("tel:", "")
                # Add leading + if PS fails to read it from online application
                if ($ResourceAccountPhoneNumber -notmatch "\+") {
                    $ResourceAccountPhoneNumber = "+$ResourceAccountPhoneNumber"
                }
                $ResourceAccountPhoneNumbers += "$ResourceAccountPhoneNumber, "
            }

            $ResourceAccountUPN = ($allResourceAccounts | Where-Object { $_.ObjectId -eq $ResourceAccount }).UserPrincipalName
            if ($ResourceAccountUPN) {
                $ResourceAccountUPNs += "$ResourceAccountUPN, "
                $ApplicationInstanceAssociationCounter ++
            }
        }

        if ($ApplicationInstanceAssociationCounter -ge 2) {
            $ResourceAccountPhoneNumbers = $ResourceAccountPhoneNumbers.Replace(",", "::")
            $ResourceAccountUPNs = $ResourceAccountUPNs.Replace(",", "::")
        }
        else {
            $ResourceAccountPhoneNumbers = $ResourceAccountPhoneNumbers.TrimEnd(", ")
            $ResourceAccountUPNs = $ResourceAccountUPNs.TrimEnd(", ")
        }

        $teamName = ""
        #Team or directly associated
        if ($VoiceAppCq.DistributionLists){
            $teamName = (Get-Team -GroupId $VoiceAppCq.DistributionLists.Guid).DisplayName
        }

        #Agents
        $AgentCounter = 0
        $AgentUPNs = ""

        foreach ($Agent in $VoiceAppCq.Agents) {
            $AgentUPN = ($allUsers | Where-Object { $_.Identity -eq $Agent.ObjectId }).UserPrincipalName
            if ($AgentUPN) {
                $AgentUPNs += "$AgentUPN, "
                $AgentCounter ++
            }
        }

        if ($AgentCounter -ge 2) {
            $AgentUPNs = $AgentUPNs.Replace(",", "::")
        }
        else {
            $AgentUPNs = $AgentUPNs.TrimEnd(", ")
        }


        #Routing settings
        $routing = ""
        $routing += "AllowOptOut: $($VoiceAppCq.AllowOptOut)::"
        $routing += "PresenceBasedRouting: $($VoiceAppCq.PresenceBasedRouting)::"
        $routing += "ConferenceMode: $($VoiceAppCq.ConferenceMode)::"

        #Exception Handling
        $exceptions = ""
        #$exceptions += "Overflow: $($VoiceAppCq.OverflowAction) ($($VoiceAppCq.OverflowActionTarget.Id))::"
        #$exceptions += "Timeout: $($VoiceAppCq.TimeoutAction) ($($VoiceAppCq.TimeoutActionTarget.Id))::"
        #$exceptions += "NoAgent: $($VoiceAppCq.NoAgentAction) ($($VoiceAppCq.NoAgentActionTarget.Id))::"

        $exceptions += "Overflow: $($VoiceAppCq.OverflowAction) ($($VoiceAppCq.OverflowThreshold)s)::"
        $exceptions += "Timeout: $($VoiceAppCq.TimeoutAction) ($($VoiceAppCq.TimeoutThreshold)s)::"
        $exceptions += "NoAgents: $($VoiceAppCq.NoAgentAction)::"

                
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "Name" -Value $VoiceAppCq.Name
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "PhoneNumber" -Value $ResourceAccountPhoneNumbers
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "UPN" -Value $ResourceAccountUPNs
        #$VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "Language" -Value $VoiceAppCq.LanguageId
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "Welcome Prompt" -Value $VoiceAppCq.WelcomeTextToSpeechPrompt
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "Routing Settings" -Value $routing
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "Routing Method" -Value $VoiceAppCq.RoutingMethod
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "AgentAlertTime" -Value $VoiceAppCq.AgentAlertTime
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "Team Name" -Value $teamName
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "Agent UPN" -Value $AgentUPNs
        $VoiceAppCqCallFlow | Add-Member -type NoteProperty -name "Exceptions" -Value $exceptions
              
        $CallFlows += $VoiceAppCqCallFlow          
    }
}

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

$html = $CallFlows | Sort-Object -Property Name |  ConvertTo-Html -Head $Header
$html.Replace("::","<br/>") | Out-File -FilePath $FilePath".html"
Write-Host "ALL DONE!! Your file has been saved to $FilePath.html"
