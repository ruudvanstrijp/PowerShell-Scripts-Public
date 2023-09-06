<#
.SYNOPSIS
PowerShell script to bulk export Teams Resource Accounts, Auto Attendants and Call Queues

.DESCRIPTION
By Ruud van Strijp - Axians
ruud.vanstrijp@axians.com

.NOTES


.EXAMPLE
.\Get-TeamsAttendants.ps1

#>
Param (
    [switch]$detailed
)
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
$FileName = "TeamsAttendants_" + (Get-Date -Format s).replace(":", "-") 
$FilePath = $PSScriptRoot + "\Output\" + $FileName
$OutputType = "HTML" #OPTIONS: CSV - Outputs CSV to specified FilePath, CONSOLE - Outputs to console


##############################

$Regex1 = '^(?:tel:)?(?:\+)?(\d+)(?:;ext=(\d+))?(?:;([\w-]+))?$'
$CallFlows = @()

#$allUsers = Get-CsOnlineUser -Filter {LineURI -ne $null -and AccountEnabled -eq $True -and AccountType -eq 'User'}
$allResourceAccounts = Get-CsOnlineApplicationInstance
$VoiceAppAas = Get-CsAutoAttendant
#$VoiceAppCqs = Get-CsCallQueue -WarningAction SilentlyContinue


if ($VoiceAppAas -ne $null) {
    foreach ($VoiceAppAa in $VoiceAppAas) {                  
        $VoiceAppAaCallFlow = New-Object System.Object

        #Resource Accounts Phone Numbers and UPN's
        $ApplicationInstanceAssociationCounter = 0
        $ResourceAccountPhoneNumbers = ""
        $ResourceAccountUPNs = ""

        foreach ($ResourceAccount in $VoiceAppAa.ApplicationInstances) {
            
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

        #Menu Options
        $MenuOptions = $VoiceAppAa.DefaultCallFlow.Menu.MenuOptions
        $aaMenuOptions = ""

        if ($MenuOptions.Count -ge 2) {
            $MenuOptionCounter = 0
            

            foreach ($MenuOption in $VoiceAppAa.DefaultCallFlow.Menu.MenuOptions) {
                $MenuKey = $MenuOption.DtmfResponse -Replace ("Tone", "")
                switch ($MenuOption.CallTarget.Type) {
                    ApplicationEndpoint { 
                        $aaMenuOptions += "Option: $MenuKey Target: $(($allResourceAccounts | Where-Object { $_.ObjectId -eq $MenuOption.CallTarget.Id }).UserPrincipalName)::"
                    }
                    SharedVoicemail { 
                        $aaMenuOptions += "Option: $MenuKey Target: SharedVoicemail::"
                    }
                    Voicemail { 
                        $aaMenuOptions += "Option: $MenuKey Target: Voicemail::"
                    }
                    ExternalPstn { 
                        $aaMenuOptions += "Option: $MenuKey Target: $($MenuOption.CallTarget.Id)::"
                    }
                    Announcement { 
                        $aaMenuOptions += "Option: $MenuKey $($MenuOption.Prompt)::"
                    }
                    Default {}
                }
            }       
        }
        else {
            
            switch ($MenuOptions[0].CallTarget.Type) {
                ApplicationEndpoint { 
                    $aaMenuOptions += "Direct transfer to: $(($allResourceAccounts | Where-Object { $_.ObjectId -eq $MenuOptions[0].CallTarget.Id }).UserPrincipalName)::"
                }
                SharedVoicemail { 
                    $aaMenuOptions += "Direct transfer to: SharedVoicemail::"
                }
                Voicemail { 
                    $aaMenuOptions += "Direct transfer to: Voicemail::"
                }
                ExternalPstn { 
                    $aaMenuOptions += "Direct transfer to: $($MenuOptions[0].CallTarget.Id)::"
                }
                Announcement { 
                    $aaMenuOptions += "$($MenuOptions[0].Prompt)::"
                }
                Default {}
            }
            
            
            #$MenuOptions[0].CallTarget.Id
            #$aaMenuOptions += "Direct transfer to: $(($allResourceAccounts | Where-Object { $_.ObjectId -eq $MenuOptions[0].CallTarget.Id }).UserPrincipalName)::"
        }

        
        $VoiceAppAaCallFlow | Add-Member -type NoteProperty -name "Name" -Value $VoiceAppAa.Name
        $VoiceAppAaCallFlow | Add-Member -type NoteProperty -name "PhoneNumber" -Value $ResourceAccountPhoneNumbers
        $VoiceAppAaCallFlow | Add-Member -type NoteProperty -name "UPN" -Value $ResourceAccountUPNs
        $VoiceAppAaCallFlow | Add-Member -type NoteProperty -name "Menu Options" -Value $aaMenuOptions

        if ($detailed) {

            $VoiceAppAaCallFlow | Add-Member -type NoteProperty -name "Language" -Value $VoiceAppAa.LanguageId

            #Get prompts
            $prompts = ""
            if($VoiceAppAa.DefaultCallFlow.Greetings.TextToSpeechPrompt){
                $prompts += "Welcome: $($VoiceAppAa.DefaultCallFlow.Greetings.TextToSpeechPrompt)::"
            }
            if($VoiceAppAa.DefaultCallFlow.Menu.Prompts.TextToSpeechPrompt){
                $prompts += "Menu: $($VoiceAppAa.DefaultCallFlow.Menu.Prompts.TextToSpeechPrompt)::"
            }

            foreach ($CallFlow in $VoiceAppAa.CallFlows) {
                
                if ($CallFlow.Menu.Name -eq "After hours call flow") {
                    $prompts += "Gesloten: $($CallFlow.Greetings.TextToSpeechPrompt)::"
                }
                else{
                    $prompts += "$($CallFlow.Name): $($CallFlow.Greetings.TextToSpeechPrompt)::"
                }
            }
            
            
            $VoiceAppAaCallFlow | Add-Member -type NoteProperty -name "Prompts" -Value $prompts

            #Resource Accounts Phone Numbers and UPN's
            $params = @{
                Identity = $VoiceAppAa.Id
                Name = $VoiceAppAa.Name
                Type = $VoiceAppAa.Type
                AssociatedConfigurationId = $VoiceAppAa.AssociatedConfigurationId
            }

            $schedules = ""
            $holidaySchedules = ""

            foreach ($Schedule in $VoiceAppAa.Schedules) {
                
                if ($Schedule.Type -eq [Microsoft.Rtc.Management.Hosted.Online.Models.ScheduleType]::Fixed) {
                    
                    $DateTimeRanges = $Schedule.FixedSchedule.DateTimeRanges
                    $dateTimeRangeStandardFormat = 'yyyy-MM-ddTHH:mm:ss';
                    $fixedScheduleDateTimeRanges = @()
                    foreach ($dateTimeRange in $DateTimeRanges) {
                        $fixedScheduleDateTimeRanges += @{
                            Start = $dateTimeRange.Start.ToString($dateTimeRangeStandardFormat)
                            End = $dateTimeRange.End.ToString($dateTimeRangeStandardFormat)
                        }
                        #Holidays
                        $holidaySchedules += "$($Schedule.Name)::"                   
                        $holidaySchedules += "Start: $($dateTimeRange.Start.ToString($dateTimeRangeStandardFormat))::"
                        $holidaySchedules += "End: $($dateTimeRange.End.ToString($dateTimeRangeStandardFormat))::"
                    }
                }
                


                if ($Schedule.Type -eq [Microsoft.Rtc.Management.Hosted.Online.Models.ScheduleType]::WeeklyRecurrence) {
                                        
                    # Define an array of day names
                    $dayNames = @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")

                    # Loop through each day and process its hours
                    foreach ($dayName in $dayNames) {
                        $dayHours = $Schedule.WeeklyRecurrentSchedule."$($dayName)Hours"
                        
                        if ($dayHours -ne $null -and $dayHours.Length -gt 0) {
                            $params["WeeklyRecurrentSchedule${dayName}Hour"] = @()
                            
                            foreach ($hour in $dayHours) {
                                $params["WeeklyRecurrentSchedule${dayName}Hour"] += @{
                                    Start = $hour.Start
                                    End = $hour.End
                                }
                                
                                if ($hour.End.TotalHours -eq 24) {
                                    $schedules += "$($dayName): Open All Day::"
                                } else {
                                    $schedules += "$($dayName): $($hour.Start) - $($hour.End)::"
                                }
                            }
                        }
                        else{
                            $schedules += "$($dayName): Closed All Day::"
                        }
                    }
     
                    $params['WeeklyRecurrentScheduleIsComplemented'] = $Schedule.WeeklyRecurrentSchedule.ComplementEnabled
                    
                    if ($Schedule.WeeklyRecurrentSchedule.RecurrenceRange -ne $null) {
                        if ($Schedule.WeeklyRecurrentSchedule.RecurrenceRange.Start -ne $null) { $params['RecurrenceRangeStart'] = $Schedule.WeeklyRecurrentSchedule.RecurrenceRange.Start }
                        if ($Schedule.WeeklyRecurrentSchedule.RecurrenceRange.End -ne $null) { $params['RecurrenceRangeEnd'] = $Schedule.WeeklyRecurrentSchedule.RecurrenceRange.End }
                        if ($Schedule.WeeklyRecurrentSchedule.RecurrenceRange.Type -ne $null) { $params['RecurrenceRangeType'] = $Schedule.WeeklyRecurrentSchedule.RecurrenceRange.Type }
                    }
                }

            }
        
            Write-Host "---------- PROCESSED AA $($VoiceAppAa.Name) --------"


            $VoiceAppAaCallFlow | Add-Member -type NoteProperty -name "Opening Hours" -Value $schedules
            $VoiceAppAaCallFlow | Add-Member -type NoteProperty -name "Holidays" -Value $holidaySchedules
        }

        $CallFlows += $VoiceAppAaCallFlow 
               
    }
}

if ($detailed) {
    $width = '150%'
}
else {
    $width = '90%'
}


$Header = "<style>
body {
    background-color: white;
    font-family:      Calibri;
}

table {
    border-width:     1px;
    border-style:     solid;
    border-color:     black;
    border-collapse:  collapse;
    width:            $($width);
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
</style>"

$html = $CallFlows | Sort-Object -Property Name |  ConvertTo-Html -Head $Header
$html.Replace("::", "<br/>") | Out-File -FilePath $FilePath".html"
Write-Host "ALL DONE!! Your file has been saved to $FilePath.html"
