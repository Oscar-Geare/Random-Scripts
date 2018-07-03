## Mass edit email templates
## 20180702
## Oscar

#Standard prompts module
if ($args[0] -eq $null) {
    $path = read-host -prompt "Path"
    $field = read-host -prompt "Field to be replaced"
    $replacement = read-host -prompt "Replacement text"
} else {
    $path = $args[0]
    $field = $args[1]
    $replacement = $args[2]
}

#Gets just the path to the file. If you try and do this with just gci $path or name or fullname it skitzes out and wont work
$fnfiles = (gci $path).fullname

#Begin Loop
foreach ($file in $fnfiles) {
    #Creates an outlook object
    $outlook= New-Object -ComObject outlook.application
    #Opens the template in the object
    $msg= $outlook.createitemfromtemplate($file)
    #Gets the old text that you're replacing
    $oldtext = $msg.$field
    #replaces oldtext with new text
    $msg.$field = $msg.$field -replace $oldtext, $replacement
    #saves the file over the old file
    $msg.saveas($file)
}
