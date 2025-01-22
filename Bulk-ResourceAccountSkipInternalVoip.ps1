Write-Host "Current state:" -ForegroundColor White
# Get Phone Number Assignments
$allPhonenumbers = Get-CsPhoneNumberAssignment -Top ([int]::MaxValue)

# Filter Phone number Assignments to only show those assigned to resource accounts
$raPhoneNumbers = $allPhonenumbers | Where-Object {$_.PstnAssignmentStatus -eq "VoiceApplicationAssigned"}

# Show all phone numbers
#$allPhonenumbers | Select-Object TelephoneNumber,NumberType,PstnAssignmentStatus | ft

# Show phone numbers assigned to resource accounts
$raPhonenumbers | Select-Object TelephoneNumber,PstnAssignmentStatus,ReverseNumberLookup | ft

# Get all Teams users (used to show the UPN of the account associated with the phonenumber)
$allUsers = Get-CsOnlineUser -Filter { EnterpriseVoiceEnabled -eq $true }

foreach ($raPhonenumber in $raPhonenumbers) {
	
	$raUPN = ($allUsers | Where-Object { $_.Identity -eq $raPhonenumber.AssignedPstnTargetId } | Select-Object UserPrincipalName).UserPrincipalName

	Write-Host "Editing phone Number: " -ForegroundColor White -NoNewLine
	Write-Host "$($raPhonenumber.TelephoneNumber)" -ForegroundColor Green -NoNewLine

	Write-Host " that belongs to account: " -ForegroundColor White -NoNewLine
	Write-Host "$($raUPN)" -ForegroundColor Green

	Set-CsPhoneNumberAssignment -Identity $raUPN -PhoneNumber $raPhonenumber.TelephoneNumber -PhoneNumberType DirectRouting -ReverseNumberLookup 'SkipInternalVoip'
	#Set-CsPhoneNumberAssignment -Identity $raUPN -PhoneNumber $raPhonenumber.TelephoneNumber -PhoneNumberType DirectRouting
}
Write-Host ""
Write-Host "Result:" -ForegroundColor White
$allPhonenumbers = Get-CsPhoneNumberAssignment
$raPhoneNumbers = $allPhonenumbers | Where-Object {$_.PstnAssignmentStatus -eq "VoiceApplicationAssigned"}
$raPhonenumbers | Select-Object TelephoneNumber,PstnAssignmentStatus,ReverseNumberLookup | ft