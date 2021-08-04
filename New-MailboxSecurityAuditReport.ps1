Function Connect-EXOnline {
	$credentials = Get-Credential
	if (!$credentials){exit}
	$Session = New-PSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
        -ConfigurationName Microsoft.Exchange -Credential $credentials `
        -Authentication Basic -AllowRedirection
	Import-PSSession $Session
}

#install the ImportExcel module if it's not installed already
If (!(Get-Module -ListAvailable -Name ImportExcel)) {Install-Module ImportExcel -scope CurrentUser -Force} 
import-module importexcel

Connect-EXOnline

#define output paths
$datestring = ((get-date).tostring("yyyy-MM-dd"))
$domains = Get-AcceptedDomain
$tenant = (Get-AcceptedDomain | Where-Object {$_.Default}).name
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$tenantpath = "$DesktopPath\MailSecurityReview\$tenant\"
$mailboxrulespath = "$tenantpath\mailbox-rules"
$loginlocationspath = "$tenantpath\loginlocations"
$reportspath = "$tenantpath\reports"
$XLSreport = "$reportspath\$tenant-report-$datestring.xlsx"

#create output paths if necessary
If (!(Test-Path -path $mailboxrulespath)){New-Item -ItemType directory -Path $mailboxrulespath -Force | Out-Null}
If (!(Test-Path -path $loginlocationspath)){New-Item -ItemType directory -Path $loginlocationspath -Force | Out-Null}
If (!(Test-Path -path $reportspath)){New-Item -ItemType directory -Path $reportspath -Force | Out-Null}


#remove a previous file if one already exists with today's date inside this tenant
$currentmailboxrules = "$mailboxrulespath\$tenant-mailboxrules-$datestring.xlsx"
If (Test-Path -path $currentmailboxrules){Remove-Item -Path $currentmailboxrules -Force}


#pull mailbox rules from EXO and evaluate
$DeliveryRuleArray = @()
$MailboxRuleArray = @()
$mailboxes = Get-Mailbox -ResultSize Unlimited

foreach ($mailbox in $mailboxes) {
	Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)" -foregroundColor Green

		$deliveryRuleHash = $null
		$deliveryRuleHash = [ordered]@{
			PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
			DisplayName        = $mailbox.DisplayName
			ForwardingSMTPAddress        = $mailbox.ForwardingSMTPAddress
			DeliverToMailboxandForward        = $mailbox.DeliverToMailboxandForward
		}
		
		#add the forwardings to an object for later export to excel
		$deliveryRuleObject = New-Object PSObject -Property $deliveryRuleHash
		$deliveryRuleArray += $deliveryruleObject 

	$rules = get-inboxrule -Mailbox $mailbox.primarysmtpaddress
	foreach ($rule in $rules) {
		$recipients = @()
		$recipients = $rule.ForwardTo | Where-Object {$_ -match "SMTP"}
		$recipients += $rule.ForwardAsAttachmentTo | Where-Object {$_ -match "SMTP"}
	 	$externalRecipients = @()
		$internalRecipients = @()
		
		#check whether forwarding recipients are external or internal
		foreach ($recipient in $recipients) {
			$email = ($recipient -split "SMTP:")[1].Trim("]")
			$domain = ($email -split "@")[1]
 
			if ($domains.DomainName -notcontains $domain) {	$externalRecipients += $email }
			else {$internalRecipients += $email }
		}
		if ($externalRecipients) {$extRecString = $externalRecipients -join ", "}
		if ($internalRecipients) {$intRecString = $internalRecipients -join ", "}
		if ($rule.RedirectTo) {$redirectString = $rule.RedirectTo -join ", "}

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
			ExternalRecipients = $extRecString
			InternalRecipients = $intRecString
		}
			
		#add the rule to an object for later export to excel
		$ruleObject = New-Object PSObject -Property $ruleHash
		$MailboxRuleArray += $ruleObject 
	}
}

#done evaluating mailboxes, export the delivery rule object for comparison
$DeliveryRuleArray | Export-Excel `
	-KillExcel `
	-Path $currentmailboxrules `
	-WorkSheetname "Delivery Rules" `
	-ClearSheet `
	-BoldTopRow `
	-Autosize `
	-FreezePane 2 `
	-Autofilter `

#done evaluating rules, export the rule object for comparison
$MailboxRuleArray | Export-Excel `
	-KillExcel `
	-Path $currentmailboxrules `
	-WorkSheetname "Mailbox Rules" `
	-ClearSheet `
	-BoldTopRow `
	-Autosize `
	-FreezePane 2 `
	-Autofilter `


#determine the paths of the current and previous (as determined by file write time)
$currentrulesetpath = Get-ChildItem $mailboxrulespath -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty fullname 
$previousrulesetpath = Get-ChildItem $mailboxrulespath -File | Sort-Object LastWriteTime -Descending | Select-Object -Skip 1 -First 1 -ExpandProperty fullname 

#make sure current and past rule sets exist
If ((Test-Path -Path $currentrulesetpath) -and ($previousrulesetpath)){
	
#compare the mailbox rules
$currentruleset = $currentrulesetpath | Import-Excel -WorkSheetname "Mailbox Rules"
$previousruleset = $previousrulesetpath | Import-Excel -WorkSheetname "Mailbox Rules"
$MailboxRuleReport = Compare-Object $currentruleset $previousruleset -Property RuleId -IncludeEqual -PassThru 
#translate the powershell sideindicator to a string that excel doesn't misinterpret as a formula
ForEach ($RuleId in $MailboxRuleReport){
#	write-host $RuleId.SideIndicator
	switch ($RuleId.SideIndicator){
		'==' { $RuleId.SideIndicator = 'rule existed' }
		'<=' { $RuleId.SideIndicator = 'rule added' }
		'=>' { $RuleId.SideIndicator = 'rule removed' }
	}
}


#output the mailbox rule comparisons to XLSX
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
$currentruleset = $null
$previousruleset = $null
$currentruleset = $currentrulesetpath | Import-Excel -WorkSheetname "Delivery Rules"
$previousruleset = $previousrulesetpath | Import-Excel -WorkSheetname "Delivery Rules"

$deliveryRuleReport = Compare-Object $currentruleset $previousruleset -Property PrimarySmtpAddress, ForwardingSMTPAddress, DeliverToMailboxandForward -IncludeEqual -PassThru 
ForEach ($Rule in $DeliveryRuleReport){
	switch ($Rule.SideIndicator){
		'==' { $Rule.SideIndicator = 'rule existed' }
		'<=' { $Rule.SideIndicator = 'rule added' }
		'=>' { $Rule.SideIndicator = 'rule removed' }
	}
}


#output the delivery rules to XLSX
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

} else {
	Write-Warning "Skipped comparison step, a comparison file is missing or has not been created yet."
}
exit-pssession