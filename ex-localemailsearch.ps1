## Search Local Emails
## Set COMObject and run SQL search over local .ost
## Oscar
## Some time in Oct 2020

# Standard Prompts module, to replace with native PowerShell cmdlet params [TODO]
if ($args[0] -eq $null) {
	$Scope = read-host -prompt "Search location? Inbox, Sent Items, etc"
	$Term = read-host -prompt "Value being searched"
} else {
	$Scope = $args[0]
	$Term = $args[1]
}

# Open an outlook ComObject. Should probably assess and close outlook beforehand to prevent any interruption
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application

# Run search on the local email box
$Emails = $Outlook.AdvancedSearch($Scope, "urn:schemas:httpmail:subject LIKE '%$Term%'", $True)

# Print results, should set up a way to export via pipe/csv/etc
$Emails.Results | Select-Object -Property subject,senton