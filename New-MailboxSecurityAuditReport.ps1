<#
.SYNOPSIS
	This script collects data on Exchange Online mailboxes rules and sign-in for review by Nexigen security team members. 

.DESCRIPTION
	This script depends on WMF 5+ and the AzureAD and ImportExcel modules, and will attempt to install the modules if they are not found. 
	This script uses the IPStack API for geolocation data; if you don't have a key already you can sign up for one here: https://ipstack.com/signup/free
	The script downloads data from Exchange Online for Exchange delivery rules and Outlook inbox rules and exports to an Excel snapshot in a per-tenant output directory. 
	If multiple snapshot entries exist in the tenant output directory, it compares the snapshots and highlights changes, as well as color codes keywords in an Excel workbook. 
	The script downloads sign-in data from AzureAD, looks up geolocation data for sign-ins, and exports to an Excel snapshot. 
	If multiple snapshot entries exist in the tenant output directory, it compares the snapshots and highlights changes, as well as color codes keywords in an Excel workbook. 
	The script stores credentials in encrypted format in the output directory; this directory can be supplied as a parameter for scheduled execution. 
	
.EXAMPLE
	.\New-MailboxSecurityAuditReport.ps1 -CredentialPath c:\users\username\Desktop\mailsecurity\contoso.com\credentials

.NOTES
    Author:             Douglas Hammond (douglas@douglashammond.com)
						Portions relating to AzureAD sign-in location written by Nexigen Security and Nexigen Service team members TJB and MC. 
	License: 			This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
						As long as you retain this notice you can do whatever you want with this stuff.
						If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.
#>
param ($CredentialPath)

#install the necessary modules if they aren't installed already
If (!(Get-Module -ListAvailable -Name ExchangePowershell)) { Install-Module ExchangePowershell -scope CurrentUser -Force } 
If (!(Get-Module -ListAvailable -Name AzureAD)) { Install-Module AzureAD -scope CurrentUser -Force } 
If (!(Get-Module -ListAvailable -Name ImportExcel)) { Install-Module ImportExcel -scope CurrentUser -Force } 
import-module ImportExcel

#credential file names
$storedkey = "key.txt"
$storeduser = "user.txt"
$storedpass = "pass.txt"

#use supplied credentials, if any
If (($null -ne $CredentialPath) -and ((Test-Path -Path $CredentialPath\$storedkey) -and (Test-Path -Path $CredentialPath\$storeduser) -and (Test-Path -Path $CredentialPath\$storedpass))) {
	$key = Get-Content $CredentialPath\$storedkey
	$username = Get-Content $CredentialPath\$storeduser
	$password = Get-Content $CredentialPath\$storedpass | ConvertTo-SecureString -Key $key
	$globaladmincreds = New-Object System.Management.Automation.PSCredential ($username, $password)
}
else {
	$globaladmincreds = Get-Credential
	if (!$globaladmincreds) { exit }
}

# old connection method, used for old get-mailbox cmdlet
#Function Connect-ExchangeOnline($credentials) {
#	$Session = New-PSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
#        -ConfigurationName Microsoft.Exchange -Credential $credentials `
#        -Authentication Basic -AllowRedirection
#	Import-PSSession $Session -AllowClobber
#}
#Connect-ExchangeOnline($globaladmincreds)

Connect-ExchangeOnline -Credential $globaladmincreds
Clear-Host

#define paths
$datestring = ((get-date).tostring("yyyy-MM-dd"))
$domains = try { Get-AcceptedDomain -ErrorAction Stop } catch { write-error "Not connected to Exchange Online. This may indicate the credentials have changed."; exit }
$tenant = (Get-AcceptedDomain | Where-Object { $_.Default }).name
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$tenantpath = "$DesktopPath\MailSecurityReview\$tenant"
$snapshotpath = "$tenantpath\snapshot"
$reportspath = "$tenantpath\reports"
$CredentialPath = "$tenantpath\credentials"
$XLSreport = "$reportspath\$tenant-report-$datestring.xlsx"

$IPStackAPIKeyPath = "$DesktopPath\MailSecurityReview\IPStackAPIKey.txt"
If (!(Test-Path -path $IPStackAPIKeyPath)) { Write-Error "IPStack API key not found. Make sure the key file exists at $IPStackAPIKeyPath. "; Read-Host -Prompt "Press Enter to exit"; exit }
$IPStackAPIKey = get-content -path $IPStackAPIKeyPath

#create output paths
If (!(Test-Path -path $snapshotpath)) { New-Item -ItemType directory -Path $snapshotpath -Force | Out-Null }
If (!(Test-Path -path $reportspath)) { New-Item -ItemType directory -Path $reportspath -Force | Out-Null }
If (!(Test-Path -path $CredentialPath)) { New-Item -ItemType directory -Path $CredentialPath -Force | Out-Null }

#create encrypted credential files for scheduled execution
If (!((Test-Path -path $CredentialPath\$storeduser) -and (Test-Path -path $CredentialPath\$storeduser))) {
	$Key = New-Object Byte[] 32
	[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
	$Key | out-file $CredentialPath\$storedkey
	$globaladmincreds.UserName | set-content "$CredentialPath\$storeduser"	
	$globaladmincreds.Password | ConvertFrom-SecureString -Key $Key | set-content "$CredentialPath\$storedpass"	
}

#remove a previous file if one already exists with today's date inside this tenant
$snapshot = "$snapshotpath\$tenant-snapshot-$datestring.xlsx"
If (Test-Path -path $snapshot) { Remove-Item -Path $snapshot -Force }

#pull mailbox rules from EXO and evaluate
$MailboxRuleResultObject = @()
$DeliveryRuleResultObject = @()

#used with old remote powershell method
#$mailboxes = Get-Mailbox -ResultSize Unlimited
$mailboxes = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties PrimarySmtpAddress, DisplayName, ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxandForward

$stepcounter = 0

foreach ($mailbox in $mailboxes) {
	$stepcounter = $stepcounter + 1
	Write-Progress -ID 0 -Activity "Processing $tenant mailbox data." -CurrentOperation "Collecting data for mailbox $stepcounter of $($mailboxes.count) - $($mailbox.primarysmtpaddress)." -PercentComplete (($stepcounter / $(($mailboxes).count)) * 100)

	$deliveryRuleHash = $null
	$deliveryRuleHash = [ordered]@{
		PrimarySmtpAddress         = $mailbox.PrimarySmtpAddress
		DisplayName                = $mailbox.DisplayName
		ForwardingAddress          = $mailbox.ForwardingAddress
		ForwardingSMTPAddress      = $mailbox.ForwardingSMTPAddress
		DeliverToMailboxandForward = $mailbox.DeliverToMailboxandForward
	}
		
	#add the forwardings to an object for later export to excel
	$deliveryRuleObject = New-Object PSObject -Property $deliveryRuleHash
	$DeliveryRuleResultObject += $deliveryruleObject 

	$rules = get-inboxrule -IncludeHidden -Mailbox $mailbox.primarysmtpaddress 3>&1
	$rulescounter = 0
	foreach ($rule in $rules) {
		$rulescounter = $rulescounter + 1
		if ($rules.count -ge 1) {
			Write-Progress -Id 1 -Activity "Processing mailbox rules." -CurrentOperation "Processing rule $rulescounter of $($rules.count)." -PercentComplete (($rulescounter / $(($rules).count)) * 100)
		}
		else {
			Write-Progress -Id 1 -Activity "Processing mailbox rules." -CurrentOperation "Processing rules for $($mailbox.primarysmtpaddress)."
		}
			
		$recipients = @()
		$recipients = $rule.ForwardTo | Where-Object { $_ -match "SMTP" }
		$recipients += $rule.ForwardAsAttachmentTo | Where-Object { $_ -match "SMTP" }
		$externalRecipients = @()
		$internalRecipients = @()
		$extRecString = $null
		$intRecString = $null
		$redirectString = $null
		#the output from this next bit is really ugly, and maybe used for internal addresses only
		# might be possible to convert it to something readable with a -replace "\[.*?\]",""
		$forwardings = @()
		$forwardings = $rule.ForwardTo
		$forwardings += $rule.ForwardAsAttachmentTo
		$forwardingsString = $null
		
		#check whether forwarding recipients are external or internal
		#this check is maybe useless, since internal forwardings don't seem to use SMTP addresses
		foreach ($recipient in $recipients) {
			$email = ($recipient -split "SMTP:")[1].Trim("]")
			$domain = ($email -split "@")[1]
 
			if ($domains.DomainName -notcontains $domain) {	$externalRecipients += $email }
			else { $internalRecipients += $email }
		}
	
		#convert the objects to strings we can put into a cell
		if ($externalRecipients) { $extRecString = $externalRecipients -join ", " }
		if ($internalRecipients) { $intRecString = $internalRecipients -join ", " }
		if ($rule.RedirectTo) { $redirectString = $rule.RedirectTo -join ", " }
		if ($forwardings) { $forwardingsString = $forwardings -join ", " }

		$ruleHash = $null
		$ruleHash = [ordered]@{
			PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
			DisplayName        = $mailbox.DisplayName
			RuleId             = $rule.Identity
			RuleName           = $rule.Name
			RuleDescription    = $rule.Description
			Enabled            = $rule.Enabled
			RedirectTo         = $redirectString
			MoveToFolder       = $rule.MoveToFolder
			Forwardings        = $forwardingsString
			ExternalRecipients = $extRecString
			InternalRecipients = $intRecString
		}
			
		#add the rule to an object for later export to excel
		$ruleObject = New-Object PSObject -Property $ruleHash
		$MailboxRuleResultObject += $ruleObject 
	}
	Write-Progress -Id 1 -Activity "Processing mailbox rules." -Completed
}

#done evaluating mailboxes, export the delivery rule object for comparison
Write-Progress -ID 0 -Activity "Processing $tenant mailbox data." -CurrentOperation "Exporting Delivery Rules to file."
$DeliveryRuleResultObject | Export-Excel `
	-KillExcel `
	-Path $snapshot `
	-WorkSheetname "Delivery Rules" `
	-ClearSheet `
	-BoldTopRow `
	-Autosize `
	-FreezePane 2 `
	-Autofilter `

#done evaluating rules, export the rule object for comparison
Write-Progress -ID 0 -Activity "Processing $tenant mailbox data." -CurrentOperation "Exporting Mailbox Rules to file."
$MailboxRuleResultObject | Export-Excel `
	-KillExcel `
	-Path $snapshot `
	-WorkSheetname "Mailbox Rules" `
	-ClearSheet `
	-BoldTopRow `
	-Autosize `
	-FreezePane 2 `
	-Autofilter `


#determine the paths of the current and previous (as determined by file write time)
$currentsnapshot = Get-ChildItem $snapshotpath -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty fullname 
$previoussnapshot = Get-ChildItem $snapshotpath -File | Sort-Object LastWriteTime -Descending | Select-Object -Skip 1 -First 1 -ExpandProperty fullname 

#make sure current and past rule sets exist
If ((Test-Path -Path $currentsnapshot) -and ($previoussnapshot)) {


	#compare the mailbox rules
	Write-Progress -ID 0 -Activity "Processing $tenant mailbox data." -CurrentOperation "Comparing mailbox rules."
	$currentruleset = $currentsnapshot | Import-Excel -WorkSheetname "Mailbox Rules"
	$previousruleset = $previoussnapshot | Import-Excel -WorkSheetname "Mailbox Rules"
	$MailboxRuleReport = Compare-Object $currentruleset $previousruleset -Property RuleId -IncludeEqual -PassThru 
	#translate the powershell sideindicator to a string that excel doesn't misinterpret as a formula
	ForEach ($RuleId in $MailboxRuleReport) {
		#	write-host $RuleId.SideIndicator
		switch ($RuleId.SideIndicator) {
			'==' { $RuleId.SideIndicator = 'rule existed' }
			'<=' { $RuleId.SideIndicator = 'rule added' }
			'=>' { $RuleId.SideIndicator = 'rule removed' }
		}
	}


	#output the mailbox rule comparisons to XLSX sheet and apply highlighting on significant text matches
	Write-Progress -ID 0 -Activity "Processing $tenant mailbox data." -CurrentOperation "Writing Mailbox Rule report to file."
	$MailboxRuleReport | Export-Excel `
		-KillExcel `
		-Path $XLSreport `
		-WorkSheetname "Mailbox Rule Changes" `
		-ClearSheet `
		-BoldTopRow `
		-Autosize `
		-FreezePane 2 `
		-Autofilter `
		-ConditionalText $(
		New-ConditionalText "rule added" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
		New-ConditionalText "rule removed" -ConditionalTextColor DarkGreen -BackgroundColor LightGreen
		New-ConditionalText "delete the message" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
		New-ConditionalText "junk" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
		New-ConditionalText "rss" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
		New-ConditionalText "archive" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
		New-ConditionalText "note" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
		New-ConditionalText "read" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
	)`
		#	-Show


	#compare the delivery rules
	Write-Progress -ID 0 -Activity "Processing $tenant mailbox data." -CurrentOperation "Comparing delivery rules."
	$currentruleset = $null
	$previousruleset = $null
	$currentruleset = $currentsnapshot | Import-Excel -WorkSheetname "Delivery Rules"
	$previousruleset = $previoussnapshot | Import-Excel -WorkSheetname "Delivery Rules"

	$deliveryRuleReport = Compare-Object $currentruleset $previousruleset -Property PrimarySmtpAddress, ForwardingSMTPAddress, DeliverToMailboxandForward -IncludeEqual -PassThru 
	#translate the powershell sideindicator to a string that excel doesn't misinterpret as a formula
	ForEach ($Rule in $DeliveryRuleReport) {
		switch ($Rule.SideIndicator) {
			'==' { $Rule.SideIndicator = 'rule existed' }
			'<=' { $Rule.SideIndicator = 'rule added' }
			'=>' { $Rule.SideIndicator = 'rule removed' }
		}
	}


	#output the delivery rules to XLSX and apply highlighting on significant text matches
	Write-Progress -ID 0 -Activity "Processing $tenant mailbox data." -CurrentOperation "Writing Delivery Rule report to file."
	$DeliveryRuleReport | Sort-Object PrimarySmtpAddress | Export-Excel `
		-KillExcel `
		-Path $XLSreport `
		-WorkSheetname "Delivery Rule Changes" `
		-ClearSheet `
		-BoldTopRow `
		-Autosize `
		-FreezePane 2 `
		-Autofilter `
		-ConditionalText $(
		New-ConditionalText "rule added" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
		New-ConditionalText "rule removed" -ConditionalTextColor DarkGreen -BackgroundColor LightGreen
	)`

}
else {
	Write-Warning "Skipped comparison step, a comparison file is missing or has not been created yet."
	Write-Warning "This message is NORMAL if this is the first time the report has run for this customer. "
}

#the code following here is mostly copied directly from the old individual scripts, with minor edits for output to excel instead of CSV

Clear-Host
Write-Progress -ID 0 -Activity "Processing $tenant signin log data." -CurrentOperation "Connecting to AzureAD."
Connect-AzureAD -Credential $globaladmincreds

# get logs from office 365 on logins for the last 90 days
$startDate = (Get-Date).AddDays(-90)
$endDate = (Get-Date)
$Logs = @()
Write-Progress -ID 0 -Activity "Processing $tenant signin log data." -CurrentOperation "Collecting logs."
do {
	$logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn
}while ($Logs.count % 3000 -eq 0 -and $logs.count -ne 0)


#get each user id that are found in the logs 
$userIds = $logs.userIds | Sort-Object -Unique

$SignInLocationResultObject = @()
$stepcounter = 0
#loops through each user id found 
foreach ($userId in $userIds) {
	$stepcounter = $stepcounter + 1
	Write-Progress -ID 0 -Activity "Processing $tenant signin log data." -CurrentOperation "Evaluating log for ID $stepcounter of $(($userIds).count) - $userId." -PercentComplete (($stepcounter / $(($userIds).count)) * 100)
 
	# start of the garbage man work --- to fix data sent from logs
	# this will exclude bad guid's, convert guid to usable email address and set the skipUserId var
	# SkipUserId var is used to keep any bad guid from getting into the report
	$skipUserID = "false" 
	$convertName = "no-input"
	if ($userId -match '\w{8}\-\w{4}\-\w{4}\-\w{4}\-\w{12}\|\|') {
		# testing for the guid with the || in it
		$skipUserID = "true"
	}
	elseif ($userId -match '^[0]{5}') {
  #testing for the all zero guid
		$skipUserID = "true"
	}
	elseif ($userId -match '\w{8}\-\w{4}\-\w{4}\-\w{4}\-\w{12}') { 
		$convertName = Get-AzureADUser -ObjectId $userId
		$convertName = $($convertName.Mail)
		if ($null -eq $convertName) {
			# make sure there is a email address attached to the guid
			$skipUserID = "true"
		}
		
	}
	elseif ($userId -match '[0-9A-Za-z]\@[0-9A-Za-z]') {
		# if userID is a email address instead of a guid it is handled here
		$convertName = $userId
	}
	else {
		write-warning 'error on $userId'
		$skipUserID = "true"
	} # --- end of the garbage man 
	
	$userLoginName = $convertName # set the userLoginName var to use the converted name from the garbage man
   
	If ($skipUserID -eq "false" ) { 
		$ips = @()
		#		Write-Host "Getting logon IPs for $userId"
		$searchResult = ($logs | Where-Object { $_.userIds -contains $userId -or $_.userIds -contains $userLoginName }).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue -ErrorVariable errorVariable
		#		Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
		$ips = $searchResult.clientip | Sort-Object -Unique
		#		Write-Host "Found $($ips.count) unique IP addresses for $userId"
		foreach ($ip in $ips) {
			#			Write-Host "Checking $ip" -ForegroundColor Yellow
			$mergedObject = @{}
			$singleResult = $searchResult | Where-Object { $_.clientip -contains $ip } | Select-Object -First 1
			Start-sleep -m 400
			# get the geo location data from the Ip addres and save to ip results 
			$ipresult = Invoke-restmethod -method get -uri "http://api.ipstack.com/$($ip)?access_key=$($IPStackAPIKey)&output=json"
			# get user agent string 
			$UserAgent = $singleResult.extendedproperties.value[0]
			#			Write-Host "Country: $($ipresult.country_name) UserAgent: $UserAgent"
			$singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
			foreach ($property in $singleResultProperties) {
				if ($property.Definition -match "object") {
					$string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
					$mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
				}
				else { $mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty }          
			}
			# using the connection to the azure ad checking if the email address exists in their system 
			$userEmailExists = try { Get-AzureADUser -ObjectId $userLoginName } catch { write-warning "$userLoginName is not an account in this tenant"; $userEmailExists = $Null }
			if ($Null -ne $userEmailExists) {
				$answer = "true"
				$mergedObject | Add-Member -Name "userExists" -Value $answer -MemberType NoteProperty
			}
			else {
				$answer = "false"
				$mergedObject | Add-Member -Name "userExists" -Value $answer -MemberType NoteProperty
			}
			$mergedObject | Add-Member -Name "userLoginName" -Value $userLoginName -MemberType NoteProperty
			$property = $null
			$ipProperties = $ipresult | get-member -MemberType NoteProperty
 
			foreach ($property in $ipProperties) {
				$mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
			}
			$SignInLocationResultObject += $mergedObject | Select-Object userLoginName, Operation, CreationTime, @{Name = "UserAgent"; Expression = { $UserAgent } }, ip, City, region_name, country_name, userExists
		}
	}
}

Write-Progress -ID 0 -Activity "Processing $tenant signin log data." -CurrentOperation "Exporting log to file."

$SignInLocationResultObject | Export-Excel `
	-KillExcel `
	-Path $snapshot `
	-WorkSheetname "Sign-in Locations" `
	-ClearSheet `
	-BoldTopRow `
	-Autosize `
	-FreezePane 2 `
	-Autofilter `
	

#make sure current and past rule sets exist
If ((Test-Path -Path $currentsnapshot) -and ($previoussnapshot)) {

	$currentsigninset = $currentsnapshot | Import-Excel -WorkSheetname "Sign-in Locations"
	$previoussigninset = $previoussnapshot | Import-Excel -WorkSheetname "Sign-in Locations"

	function LoginOccurredPriorWeek ($csv, $userID) {
		# Initialize empty array
		$arrayOfUserIDs = @()

		# Loop through csv and add each user ID to array
		$csv | ForEach-Object { $arrayOfUserIDs += $_.userLoginName }

		# Check if array contains the passed-in user ID. Returns True or False
		$arrayOfUserIDs.Contains($userID)
	}

	function UserLoggedInDifferentIPsSameWeek ($csv, $userID) {
		$arrayOfIPs = @()
    
		foreach ($object in $csv) {
			if ($object.userLoginName -eq $userID) {
				$arrayOfIPs += $object.ip
			}
		}

		if ($arrayOfIPs.Count -gt 1) {
			return $true
		}
		else {
			return $false
		}
	}

	function IsKnownIPFromPriorWeek ($csv, $userID, $loginIP) {
		$arrayOfIPs = @()
    
		foreach ($object in $csv) {
			if ($object.userLoginName -eq $userID) {
				$arrayOfIPs += $object.ip
			}
		}
    
		if ($arrayOfIPs -contains $loginIP) {
			return $true
		}  
		else {
			return $false
		}  
	}


	$signincomparisonresultobject = @()
	$stepcounter = 0
	foreach ($object in $currentsigninset) {
		$userID = $object.UserLoginName
		$loginIP = $object.ip
		$userExists = $object.userExists

		$stepcounter = $stepcounter + 1
		Write-Progress -ID 0 -Activity "Processing $tenant signin log data." -CurrentOperation "Comparing signins for ID $stepcounter of $(($currentsigninset).count) - $userId." -PercentComplete (($stepcounter / $(($currentsigninset).count)) * 100)

		if ( ($userExists -eq 'False') -and ($userID -ne 'Unknown') ) {
			$flag = 'ExternalMailbox'
		}
		elseif ($userID -eq 'Unknown') {
			$flag = 'UnknownUser'
		}
		elseif ( !(LoginOccurredPriorWeek $previoussigninset $userID) ) {     
			$flag = 'NoLoginsLastScan'
		}
		elseif (UserLoggedInDifferentIPsSameWeek $currentsigninset $userID) {
			if ((IsKnownIPFromPriorWeek $previoussigninset $userID $loginIP)) {
				$flag = 'Normal'    
			}
			else {
				$flag = 'DifferentIPThanLastScan1'
			}
		}
		elseif (!(IsKnownIPFromPriorWeek $previoussigninset $userID $loginIP)) {
			$flag = 'DifferentIPThanLastScan2'    
		}
		else {
			$flag = 'Normal'
		}


		$accessHash = $null
		$accessHash = [ordered]@{
			userID       = $object.userLoginName
			operation    = $object.Operation
			creationTime = $object.CreationTime
			userAgent    = $object.UserAgent
			loginIP      = $object.ip
			loginCity    = $object.city
			loginRegion  = $object.region_name
			loginCountry = $object.country_name
			userExists   = $object.userExists
			flag         = $flag
		}
		
		#add the ordered hash to an object for later export to excel
		$accessObject = New-Object PSObject -Property $accessHash
		$signincomparisonresultobject += $accessObject 	
	
	}

	Write-Progress -ID 0 -Activity "Processing $tenant signin log data." -CurrentOperation "Exporting comparisons to file."
	$signincomparisonresultobject | Export-Excel `
		-KillExcel `
		-Path $XLSreport `
		-WorkSheetname "Sign-in Locations" `
		-ClearSheet `
		-BoldTopRow `
		-Autosize `
		-FreezePane 2 `
		-Autofilter `
		-ConditionalText $(
		New-ConditionalText "DifferentIPThanLastScan" -ConditionalTextColor DarkBlue -BackgroundColor LightBlue
		New-ConditionalText "NoLoginsLastScan" -ConditionalTextColor Black -BackgroundColor Yellow
		New-ConditionalText "ExternalMailbox" -ConditionalTextColor White -BackgroundColor Orange
		New-ConditionalText "UnknownUser" -ConditionalTextColor DarkRed -BackgroundColor LightPink
	)
	#-Show #comment this line out if you don't want the report to auto-launch when it's finished, or if you're running multiple instances simultaneously


}
else {
	Write-Warning "Skipped comparison step, a comparison file is missing or has not been created yet."
	Write-Warning "This message is NORMAL if this is the first time the report has run for this customer. "
}

Write-Progress -ID 0 -Activity "Processing $tenant signin log data." -CurrentOperation "Done." -Completed

Clear-Host
Write-Host "TIP: Run this script unattended with $PSCommandPath -CredentialPath $($CredentialPath) ."
Read-Host -Prompt "Press Enter to exit"