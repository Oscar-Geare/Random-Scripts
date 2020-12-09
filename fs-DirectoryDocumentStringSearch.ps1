## Search list of documents for certain strings
## 20/01/2020
## Oscar & Loverton

#Get all relevant items. Probably should interact with paramaters or argumets [TODO]
Get-ChildItem r:\ -recurse -include *Command*.docx,*Initial*.docx | Where-Object { $_.CreationTime -gt (Get-Date).AddDays(-90) } | Select-Object Fullname | Export-CSV "C:\Temp\documentlist.csv" -NoTypeInformation

#Stop that pesky word application to stop it from interfering with our work. I hope you didn't have any unsaved documents. Probably should have a warning [TODO]
Stop-Process -Name *WINWORD* -Force

#Create the things for word to be able to do whatever word does
$application = New-Object -comobject word.application
$application.visible = $false

#Loverton made these and I think they're kinda neat so they stay. Again, interaction with params or args [TODO]
[string[]]$docs = Get-Content -path "C:\temp\documentlist.csv"
[string[]]$ipList = Get-Content -path "c:\temp\watchlistValues.txt"

#This is to count things
$counter = 1

#Create the hash table and objects to build a CSV at the end. Probably a neater way to make it work [TODO]
$results = @()
$matches = @{
	"File" = ""
	"IP" = ""
}
#Ensuring the variable is blank
$document = ""

#Run through all the documents
foreach ($doc in $docs) {
	#Neato progress bar
	Write-Progress -Activity "Processing files" -status "Processing $($doc)" -PercentComplete ($counter /$docs.Count * 100) 
	#Open the document in the application we created the at the start
	#Some documetns error about here, saying it's null. Should probably work out why [TODO]
	$document = $application.documents.open($doc) 
	#Get all the content of the document. We have to do it this way because we use stupid tables to make the formatting look nice, and that doesn't play well with powershell
	$range = $document.Paragraphs | ForEach-Object { $_.Range.Text }
	
	#Run through all the IPs
	foreach ($ip in $ipList) {
		#Look to see if you can find an IP in the document
		$wordFound = $range | Select-String -pattern $ip
		#If we do find the iP in the document, we add the name and the IP into a CSV
		if($wordFound) { 
			$matches."File" = $doc
			$matches."IP" = $ip
			#Debug writing
			##Write-Host "$($ip) in $($doc)"
			#Add the matches to the results table
			$results += New-Object -TypeName PsCustomObject -Property $matches
		} #end if $wordFound
    } #end foreach $ip
	#Close the document
	$document.close()
	#Make the countyboi count
	$counter++
	#go countyboi go
} #end foreach $doc

#close the word application
$application.quit()
#Place where the results go. Params args etc [TODO]
$output = "C:\Temp\Emotet Incidents.csv"

#If there are results export it to the place listed above
if ($results) {
	$results | Export-CSV $output -NoTypeInformation
}

#clean up stuff. Should probably kill some of the other variables as well
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($application) | Out-Null