#Archive OD4B
#Oscar
#Sanitised version 20161025, search for 'domain.tld', 'domain_tld' and 'o365domain' for relevant locations to change
#Fuck you I like tabs not spaces

if ($args[2] -eq $null) {
	$UsrFN = read-host -prompt "User's first name"
	$UsrSN = read-host -prompt "User's surname"
	$CLoc = read-host -prompt "Copy Location"
} else {
	$UsrFN=$args[0]
	$UsrSN=$args[1]
	$CLoc=$args[2]
}

#Check to see if SPO module is installed
if (Get-Module -ListAvailable -Name Microsoft.Online.Sharepoint.Powershell) {
    Write "Sharepoint Online Module Installed"
} else {
    Write "Install Sharepoint Online Management Shell"
	exit 1
}

#Sign into SPO
$UserCredential = Get-Credential -Message "Admin Credentials" -UserName $env:username@domain.tld
Connect-SPOService -url https://o365domain-admin.sharepoint.com -Credential $UserCredential

#Add Registry key
$registryPath = "hkcu:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\365domain-my.sharepoint.com"

#Intranet
$value1 = "1"
#Trusted Zone
$value2 = "2"
#Wildcard value
$Name2 = "*"

#Create path if it's not present yet
if (!(Test-Path $registryPath))  {
	New-Item -Path $registryPath -Force | Out-Null
	write "Reg Path created"
}

#Add Dword values for Intranet and Trusted sites
New-ItemProperty -path $RegistryPath -name $name2 -value $value2 -PropertyType DWORD -Force | Out-Null
write "Wildcard reg key added trusted"
New-ItemProperty -Path $registryPath -Name $name2 -Value $value1 -PropertyType DWORD -Force | Out-Null
write "Wildcard reg key added local"


#Need to open SPO in IE for the creds to be authenticated
write "Signing into SPO"
$IE=new-object -com internetexplorer.application
$IE.navigate2("o365domain-my.sharepoint.com")
$IE.visible=$true

#Literally does nothing but hold up the script and wait for authentication. You can write whatever you want
$go = read-host -prompt "Continue?"


$Url = "https://o365domain-my.sharepoint.com/personal/$($UsrFN)_$($UsrSN)_domain_tld"

write "Connecting to $($Url)"
$site = Get-SPOSite -Identity $Url
Set-SPOUser -Site $site.Url -LoginName $env:username@domain.tld -IsSiteCollectionAdmin $true

$Drv = ls function:[u-z]: -n | ?{ !(test-path $_) } | select -first 1
write "Mapping to $($Drv)"


$Rt = "\\o365domain-my.sharepoint.com@ssl\davwwwroot\personal\$($UsrFN)_$($UsrSN)_domain_tld\documents"

net use $drv $rt

robocopy $Drv "$($Cloc)\OD4B" * /move /r:6 /w:10 /e /np /tee /log+:"$($Cloc)\OD4B\robocopy.log"

net use $drv /delete
disconnect-sposervice
