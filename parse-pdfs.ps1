#set working directories
$maindir = 'C:\Temp\AD-Auto-Account-Creation\2875'
$workingdir = 'C:\Temp\AD-Auto-Account-Creation\2875\working2875'
$storagedir = 'C:\Temp\AD-Auto-Account-Creation\2875\processed2875s'
$initialdirectory = '$env:USERPROFILE\Downloads'

<# block temporarily commented out when testing script
#do initial check in downloads folder to see if any pdfs exist that have been added in the last 3 days, if not, quit script
$downloadedforms = @(Get-ChildItem -Path $env:USERPROFILE\Downloads -Filter *2875*.pdf | where {
$_.LastWriteTime -gt (Get-Date).AddDays(-3)}).Count
    $popup = New-Object -ComObject wscript.Shell
        if ($downloadedforms -eq 0) {
            $popup.Popup("Can't find any more 2875 forms to process!",0,"All Done!",0x10)
            Exit
            }
#>

#cleanup from last run
Remove-Item -Path $workingdir\*.txt
Move-Item -Path $workingdir\*.pdf -Destination $storagedir -Force

#select which PDF you'd like to process  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialdirectory
        $OpenFileDialog.filter = "PDF (*.PDF)|*.PDF"
        $OpenFileDialog.ShowDialog() | Out-Null
        Move-Item -Path $OpenFileDialog.filename -Destination $workingdir #move item into working directory
 
#capture 2875's original name and store it in a variable
$item = Get-Childitem -Path $workingdir -Filter *.pdf
    $oldname = $item.Name

#rename 2875 so pdftotext batch script can read it
Get-Childitem -Path $workingdir -Filter *.pdf | Rename-Item -Newname {"DD2875-input" + $_.extension} -Force
Get-AuthenticodeSignature -FilePath 'C:\TEMP\AD-Auto-Account-Creation\2875\working2875\DD2875-input.pdf'

#run pdftotext batch script
Start-Process -FilePath $maindir\pdftopngbatch.bat -Wait 

#rename 2875 back to original name
Get-Childitem -Path $workingdir -Filter *.pdf | select-object -first 1 | Rename-Item -NewName $oldname

#rename output document to recognizable name
Get-Childitem -Path $workingdir -Filter *input.txt | Rename-Item -Newname {"DD2875-output" + $_.extension} -Force

#parse 2875; import and format text into an array
$text = get-content -path $workingdir\DD2875-output.txt 

#assign pieces of imported text to variables
$name = '1. NAME \(Last, First, Middle Initial\)' #search for name heading on form
    $namelinenumber = $text | select-string $name -CaseSensitive
    $namelinenumber = $namelinenumber.linenumber #get line number
        $name = $text[$namelinenumber] #select text directly after line number
        $name = $name -replace ',','' #remove commas
        $name = $name -replace '\.','' #remove periods
        $name = $name -split ' '  #split text line by spaces to separate names
       
        $first = $name | select-object  -index 1 #select second line and assign to first name variable
        $MI = $name | select-object  -index 2
        $last = $name | select-object  -index 0
        $username = "$first.$MI.$last" #create username from existing name variables
            $username = $username.ToLower() #convert username to lowercase

$email = '5. OFFICIAL E-MAIL ADDRESS'
    $emaillinenumber = $text | select-string $email -CaseSensitive
    $emaillinenumber = $emaillinenumber.linenumber
        $email = $text[$emaillinenumber]

$phone = '4. PHONE \(DSN or Commercial\)' 
    $phonelinenumber = $text | select-string $phone -CaseSensitive
    $phonelinenumber = $phonelinenumber.linenumber 
        $phone = $text[$phonelinenumber] 
        $phone = $phone -replace '-','' 
        $phone = $phone -replace '\(','' 
        $phone = $phone -replace '\)','' 
        $phone = $phone -replace ' ','' 
            $phone = "{0:###\.###\.####}" -f [int64]$phone #format number

        if ($phone -match '^\.\d\d\d\.\d\d\d\d') { #add a '443' to beginning of phone number if it's missing
            $phone = '443'+$phone
            }

$EDIPI = $text -match "\.\d\d\d\d\d\d\d\d\d\d" #search for 10-digit strings following periods
    $EDIPI = $EDIPI -split " " #split at spaces
    $EDIPI = $EDIPI -split "\." #split at period to separate numbers from name 
        $EDIPI = $EDIPI -Match '^\d\d*\d\d' | select -first 1 #filter out EDIPI-like strings and select the first match

$justificationheading = '13. JUSTIFICATION FOR ACCESS' #find beginning of the justification paragraph
    $justificationlinenumber1 = $text | select-string $justificationheading -CaseSensitive
    $justificationlinenumber1 = $justificationlinenumber1.linenumber 
         $justificationend = '14. TYPE OF ACCESS REQUIRED:' #find end of the justification paragraph
             $justificationlinenumber2 = $text | select-string $justificationend -CaseSensitive
             $justificationlinenumber2 = $justificationlinenumber2.linenumber
                $justification = $text[$justificationlinenumber1..($justificationlinenumber2-2)] #capture all lines between beginning and end of justification paragraph
                    $justification = $justification | where-object {$_.trim() -ne ""} #change double-spaced lines to single-spaced

$building = '7. OFFICIAL MAILING ADDRESS'
    $buildinglinenumber = $text | select-string $building -CaseSensitive
    $buildinglinenumber = $buildinglinenumber.linenumber 
        $building = $text[$buildinglinenumber] 
        $building = $building -split {$_ -eq " "} #split text line by spaces in order to pull out number
            $building = $building -match '60+' #filter lines by building number

$office = '3. OFFICE SYMBOL\/DEPARTMENT'
    $officelinenumber = $text | select-string $office -CaseSensitive
    $officelinenumber = $officelinenumber.linenumber 
        $office = $text[$officelinenumber] 
            $office = $office -replace ' ','-' #replace spaces with hyphens

#to capture user's signature, import everything from section 11. to PART II and filter out unnecessary characters
$user1 = "11. USER SIGNATURE"
    $userlinenumber1 = $text | select-string $user1 -CaseSensitive
    $userlinenumber1 = $userlinenumber1.linenumber 
$user2 = "PART II - ENDORSEMENT OF ACCESS BY INFORMATION OWNER, USER SUPERVISOR OR GOVERNMENT SPONSOR \(If individual is a contractor - provide company name, contract number, and date of contract expiration in Block 16\.\)" 
    $userlinenumber2 = $text | select-string $user2 -CaseSensitive
    $userlinenumber2 = $userlinenumber2.linenumber 
        $user = $text[($userlinenumber1-1)..($userlinenumber2-2)] 
            $user = $user -replace '11\.',''
            $user = $user -replace '12\.',''
            $user = $user -replace '13\.',''
            $user = $user -replace '14\.',''
            $user = $user -replace '15\.',''
            $user = $user -replace '16\.',''
            $user = $user -replace '17\.',''
            $user = $user -replace '18\.',''
            $user = $user -replace '19\.',''
            $user = $user -replace '20\.',''
            $user = $user -replace '21\.',''
            $user = $user -replace '21.\.',''
            $user = $user -replace '22\.',''
            $user = $user -replace '23\.',''
            $user = $user -replace '24\.',''
            $user = $user -replace '25\.',''
            $user = $user -replace '26\.',''
            $user = $user -replace '27\.',''
            $user = $user -replace '28\.',''
            $user = $user -replace '29\.',''
            $user = $user -replace '30\.',''
            $user = $user -replace '31\.',''
            $user = $user -replace '32\.',''
            $user = $user -replace '33\.',''
            $user = $user -replace 'supervisor',''
            $user = $user -replace "iao",''
            $user = $user -replace 'user ',''
            $user = $user -replace 'security',''
            $user = $user -replace 'signed',''
            $user = $user -replace 'signature',''
            $user = $user -replace 'digitally ',''
            $user = $user -replace ' by ',''
            $user = $user -replace "'s ",''
            $user = $user -replace 'organization\/',''
            $user = $user -replace 'CECOM',''
            $user = $user -replace 'SEC',''
            $user = $user -replace 'SERVICES',''
            $user = $user -replace 'PREVIOUS',''
            $user = $user -replace 'EDITION',''
            $user = $user -replace 'IS OBSOLETE',''
            $user = $user -replace 'PHONE NUMBER',''
            $user = $user -replace 'department',''
            $user = $user -replace 'part ',''
            $user = $user -replace 'II ',''
            $user = $user -replace 'III ',''
            $user = $user -replace 'IV ',''
            $user = $user -replace ' - ',''
            $user = $user -replace 'T\.R\.',''
            $user = $user -replace 'ENDORSEMENT OF ACCESS BY INFORMATION OWNER',''
            $user = $user -replace 'COMPLETIONAUTHORIZED STAFF PREPARING ACCOUNT INFORMATION',''
            $user = $user -replace 'USER user OR GOVERNMENT SPONSOR',''
            $user = $user -replace 'DATE \(YYYYMMDD\)',''
            $user = $user -replace '2018\d\d\d\d',''
            $user = $user -replace "-\d\d'\d\d'",''
            $user = $user -replace " \d\d:\d\d:\d\d ",''
            $user = $user -replace "Date: 20\d\d\.",''
            $user = $user -replace 'SIGN HERE ',''
            $user = $user -replace ' DD FORM 2875, ',''
            $user = $user -replace "DN: ",''
            $user = $user -replace "US, ",''
            $user = $user -replace "U\.S\. ",''
            $user = $user -replace "DOD, ",''
            $user = $user -replace "=CONTRACTOR",''
            $user = $user -replace "PKI, ",''
            $user = $user -replace "Government, ",''
            $user = $user -replace "c=",''
            $user = $user -replace "o=",''
            $user = $user -replace "ou=",''
            $user = $user -replace "cn=",''
            $user = $user -replace ' \s ',''
            $user = $user -replace ' \S ',''
            $user = $user -replace ',',''
            $user = $user -replace '  *'," "
            $user = $user -match '\.[ -~]*.\d\d\d\d\d\d\d\d\d*\d*' | select -first 1 #final filter to pull out the signature string
                
#retrieve signatures and timestamps of user's certifying officials
$supervisor1 = "18. SUPERVISOR'S SIGNATURE" #begin capturing line beginning after "18. SUPERVISOR'S SIGNATURE"
    $supervisorlinenumber1 = $text | select-string $supervisor1 -CaseSensitive
    $supervisorlinenumber1 = $supervisorlinenumber1.linenumber 
$supervisor2 = "21. SIGNATURE OF INFORMATION OWNER\/OPR" #stop capturing lines at "19. DATE (YYYYMMDD)"
    $supervisorlinenumber2 = $text | select-string $supervisor2 -CaseSensitive
    $supervisorlinenumber2 = $supervisorlinenumber2.linenumber 
        $supervisor = $text[($supervisorlinenumber1-1)..($supervisorlinenumber2-1)]
	        $supervisor = $supervisor -replace '11\.',''
            $supervisor = $supervisor -replace '12\.',''
            $supervisor = $supervisor -replace '13\.',''
            $supervisor = $supervisor -replace '14\.',''
            $supervisor = $supervisor -replace '15\.',''
            $supervisor = $supervisor -replace '16\.',''
            $supervisor = $supervisor -replace '17\.',''
            $supervisor = $supervisor -replace '18\.',''
            $supervisor = $supervisor -replace '19\.',''
            $supervisor = $supervisor -replace '20\.',''
            $supervisor = $supervisor -replace '21\.',''
            $supervisor = $supervisor -replace '21.\.',''
            $supervisor = $supervisor -replace '22\.',''
            $supervisor = $supervisor -replace '23\.',''
            $supervisor = $supervisor -replace '24\.',''
            $supervisor = $supervisor -replace '25\.',''
            $supervisor = $supervisor -replace '26\.',''
            $supervisor = $supervisor -replace '27\.',''
            $supervisor = $supervisor -replace '28\.',''
            $supervisor = $supervisor -replace '29\.',''
            $supervisor = $supervisor -replace '30\.',''
            $supervisor = $supervisor -replace '31\.',''
            $supervisor = $supervisor -replace '32\.',''
            $supervisor = $supervisor -replace '33\.',''
            $supervisor = $supervisor -replace 'user',''
            $supervisor = $supervisor -replace "iao",''
            $supervisor = $supervisor -replace 'supervisor',''
            $supervisor = $supervisor -replace 'security',''
            $supervisor = $supervisor -replace 'digitally ',''
            $supervisor = $supervisor -replace 'signed',''
            $supervisor = $supervisor -replace 'signature',''
            $supervisor = $supervisor -replace ' by ',''
            $supervisor = $supervisor -replace "'s ",''
            $supervisor = $supervisor -replace 'organization\/',''
            $supervisor = $supervisor -replace 'CECOM',''
            $supervisor = $supervisor -replace 'SEC',''
            $supervisor = $supervisor -replace 'SERVICES',''
            $supervisor = $supervisor -replace 'PREVIOUS',''
            $supervisor = $supervisor -replace 'EDITION',''
            $supervisor = $supervisor -replace 'IS OBSOLETE',''
            $supervisor = $supervisor -replace 'PHONE NUMBER',''
            $supervisor = $supervisor -replace 'department',''
            $supervisor = $supervisor -replace 'part ',''
            $supervisor = $supervisor -replace 'II ',''
            $supervisor = $supervisor -replace 'III ',''
            $supervisor = $supervisor -replace 'IV ',''
            $supervisor = $supervisor -replace ' - ',''
            $supervisor = $supervisor -replace 'T\.R\.',''
            $supervisor = $supervisor -replace 'ENDORSEMENT OF ACCESS BY INFORMATION OWNER',''
            $supervisor = $supervisor -replace 'COMPLETIONAUTHORIZED STAFF PREPARING ACCOUNT INFORMATION',''
            $supervisor = $supervisor -replace 'supervisor supervisor OR GOVERNMENT SPONSOR',''
            $supervisor = $supervisor -replace 'DATE \(YYYYMMDD\)',''
            $supervisor = $supervisor -replace '2018\d\d\d\d',''
            $supervisor = $supervisor -replace "-\d\d'\d\d'",''
            $supervisor = $supervisor -replace " \d\d:\d\d:\d\d ",''
            $supervisor = $supervisor -replace "Date: 20\d\d\.",''
            $supervisor = $supervisor -replace 'SIGN HERE ',''
            $supervisor = $supervisor -replace ' DD FORM 2875, ',''
            $supervisor = $supervisor -replace "DN: ",''
            $supervisor = $supervisor -replace "US, ",''
            $supervisor = $supervisor -replace "U\.S\. ",''
            $supervisor = $supervisor -replace "DOD, ",''
            $supervisor = $supervisor -replace "=CONTRACTOR",''
            $supervisor = $supervisor -replace "PKI, ",''
            $supervisor = $supervisor -replace "Government, ",''
            $supervisor = $supervisor -replace "c=",''
            $supervisor = $supervisor -replace "o=",''
            $supervisor = $supervisor -replace "ou=",''
            $supervisor = $supervisor -replace "cn=",''
            $supervisor = $supervisor -replace ',',''
            $supervisor = $supervisor -replace ' \S ',''
            $supervisor = $supervisor -replace ' \s ',''
            $supervisor = $supervisor -replace '  *'," "
            $supervisor = $supervisor -match '\.[ -~]*.\d\d\d\d\d\d\d\d\d*\d*' | select -first 1
			#alternate method of parsing signature, currently not being used
            <# 
            [regex]$supervisorsignature = '\w*\s*\S*\w+\.\w+\.\w+\S*\.\d{10}' #pull out only signature and date from lines
            $supervisorsignature.Matches($supervisor) | foreach-object {$_.Value} | Select-object -First 1 | Out-String -OutVariable supsig
            [regex]$supervisortimestamp = 'Date: 20\d\d.\d\d.\d\d \d\d:\d\d:\d\d'
            $supervisortimestamp.Matches($supervisor) | foreach-object {$_.Value} | Select-object -First 1 | Out-String -OutVariable supdate
                $supervisor = "$supsig$supdate"
                #>           
$IAO1 = "22. SIGNATURE OF IAO OR APPOINTEE"
    $IAOlinenumber1 = $text | select-string $IAO1 -CaseSensitive
    $IAOlinenumber1 = $IAOlinenumber1.linenumber 
$IAO2 = "24. PHONE NUMBER"
    $IAOlinenumber2 = $text | select-string $IAO2 -CaseSensitive
    $IAOlinenumber2 = $IAOlinenumber2.linenumber 
        $IAO = $text[($IAOlinenumber1-1)..($IAOlinenumber2-1)]
            $iao = $iao -replace '11\.',''
            $iao = $iao -replace '12\.',''
            $iao = $iao -replace '13\.',''
            $iao = $iao -replace '14\.',''
            $iao = $iao -replace '15\.',''
            $iao = $iao -replace '16\.',''
            $iao = $iao -replace '17\.',''
            $iao = $iao -replace '18\.',''
            $iao = $iao -replace '19\.',''
            $iao = $iao -replace '20\.',''
            $iao = $iao -replace '21\.',''
            $iao = $iao -replace '21.\.',''
            $iao = $iao -replace '22\.',''
            $iao = $iao -replace '23\.',''
            $iao = $iao -replace '24\.',''
            $iao = $iao -replace '25\.',''
            $iao = $iao -replace '26\.',''
            $iao = $iao -replace '27\.',''
            $iao = $iao -replace '28\.',''
            $iao = $iao -replace '29\.',''
            $iao = $iao -replace '30\.',''
            $iao = $iao -replace '31\.',''
            $iao = $iao -replace '32\.',''
            $iao = $iao -replace '33\.',''
            $iao = $iao -replace 'user',''
            $iao = $iao -replace "supervisor",''
            $iao = $iao -replace 'iao ',''
            $iao = $iao -replace 'security',''
            $iao = $iao -replace 'digitally ',''
            $iao = $iao -replace 'signed',''
            $iao = $iao -replace 'signature',''
            $iao = $iao -replace ' by ',''
            $iao = $iao -replace "'s ",''
            $iao = $iao -replace 'organization\/',''
            $iao = $iao -replace 'CECOM',''
            $iao = $iao -replace 'SEC',''
            $iao = $iao -replace 'SERVICES',''
            $iao = $iao -replace 'PREVIOUS',''
            $iao = $iao -replace 'EDITION',''
            $iao = $iao -replace 'IS OBSOLETE',''
            $iao = $iao -replace 'PHONE NUMBER',''
            $iao = $iao -replace 'department',''
            $iao = $iao -replace 'part ',''
            $iao = $iao -replace 'II ',''
            $iao = $iao -replace 'III ',''
            $iao = $iao -replace 'IV ',''
            $iao = $iao -replace ' - ',''
            $iao = $iao -replace 'T\.R\.',''
            $iao = $iao -replace 'ENDORSEMENT OF ACCESS BY INFORMATION OWNER',''
            $iao = $iao -replace 'DD FORM 2875',''
            $iao = $iao -replace 'OF OR APPOINTEE',''
            $iao = $iao -replace 'DATE \(YYYYMMDD\)',''
            $iao = $iao -replace '2018\d\d\d\d',''
            $iao = $iao -replace "-\d\d'\d\d'",''
            $iao = $iao -replace " \d\d:\d\d:\d\d ",''
            $iao = $iao -replace "Date: 20\d\d\.",''
            $iao = $iao -replace 'SIGN HERE ',''
            $iao = $iao -replace ' DD FORM 2875, ',''
            $iao = $iao -replace "DN: ",''
            $iao = $iao -replace "US, ",''
            $iao = $iao -replace "U\.S\. ",''
            $iao = $iao -replace "DOD, ",''
            $iao = $iao -replace "=CONTRACTOR",''
            $iao = $iao -replace "PKI, ",''
            $iao = $iao -replace "Government, ",''
            $iao = $iao -replace "c=",''
            $iao = $iao -replace "o=",''
            $iao = $iao -replace "ou=",''
            $iao = $iao -replace "cn=",''
            $iao = $iao -replace ' \s ',''
            $iao = $iao -replace ' \S ',''
            $iao = $iao -replace ',',''
            $iao = $iao -replace '  *'," "
            $iao = $iao -match '\.[ -~]*.\d\d\d\d\d\d\d\d\d*\d*' | select -first 1
            <#
            [regex]$IAOsignature = '\w*\s*\S*\w+\.\w+\.\w+\S*\.\d{10}'
            $IAOsignature.Matches($IAO) | foreach-object {$_.Value} | Select-object -First 1 | Out-String -OutVariable IAOsig
            [regex]$IAOtimestamp = 'Date: 20\d\d.\d\d.\d\d \d\d:\d\d:\d\d'
            $IAOtimestamp.Matches($IAO) | foreach-object {$_.Value} | Select-object -First 1 | Out-String -OutVariable IAOdate
                $IAO = "$IAOsig$IAOdate"
                #>
                  
$security1 = "31. SECURITY MANAGER SIGNATURE" 
    $securitylinenumber1 = $text | select-string $security1 -CaseSensitive
    $securitylinenumber1 = $securitylinenumber1.linenumber 
$security2 = "PART IV - COMPLETION BY AUTHORIZED STAFF PREPARING ACCOUNT INFORMATION"
    $securitylinenumber2 = $text | select-string $security2 -CaseSensitive
    $securitylinenumber2 = $securitylinenumber2.linenumber
        $security = $text[($securitylinenumber1-1)..($securitylinenumber2-1)] 
	    $security = $security -replace '11\.',''
            $security = $security -replace '12\.',''
            $security = $security -replace '13\.',''
            $security = $security -replace '14\.',''
            $security = $security -replace '15\.',''
            $security = $security -replace '16\.',''
            $security = $security -replace '17\.',''
            $security = $security -replace '18\.',''
            $security = $security -replace '19\.',''
            $security = $security -replace '20\.',''
            $security = $security -replace '21\.',''
            $security = $security -replace '21.\.',''
            $security = $security -replace '22\.',''
            $security = $security -replace '23\.',''
            $security = $security -replace '24\.',''
            $security = $security -replace '25\.',''
            $security = $security -replace '26\.',''
            $security = $security -replace '27\.',''
            $security = $security -replace '28\.',''
            $security = $security -replace '29\.',''
            $security = $security -replace '30\.',''
            $security = $security -replace '31\.',''
            $security = $security -replace '32\.',''
            $security = $security -replace '33\.',''
            $security = $security -replace 'user',''
            $security = $security -replace "supervisor",''
            $security = $security -replace 'security ',''
            $security = $security -replace 'iao',''
            $security = $security -replace 'digitally ',''
            $security = $security -replace 'signed',''
            $security = $security -replace 'signature',''
            $security = $security -replace ' by ',''
            $security = $security -replace "'s ",''
            $security = $security -replace 'organization\/',''
            $security = $security -replace 'CECOM',''
            $security = $security -replace 'SEC',''
            $security = $security -replace 'manager',''
            $security = $security -replace 'SERVICES',''
            $security = $security -replace 'PREVIOUS',''
            $security = $security -replace 'EDITION',''
            $security = $security -replace 'IS OBSOLETE',''
            $security = $security -replace 'PHONE NUMBER',''
            $security = $security -replace 'department',''
            $security = $security -replace 'part ',''
            $security = $security -replace 'II ',''
            $security = $security -replace 'III ',''
            $security = $security -replace 'IV ',''
            $security = $security -replace ' - ',''
            $security = $security -replace 'T\.R\.',''
            $security = $security -replace 'ENDORSEMENT OF ACCESS BY INFORMATION OWNER',''
            $security = $security -replace 'COMPLETIONAUTHORIZED STAFF PREPARING ACCOUNT INFORMATION',''
            $security = $security -replace 'security security OR GOVERNMENT SPONSOR',''
            $security = $security -replace 'DATE \(YYYYMMDD\)',''
            $security = $security -replace '2018\d\d\d\d',''
            $security = $security -replace "-\d\d'\d\d'",''
            $security = $security -replace " \d\d:\d\d:\d\d ",''
            $security = $security -replace "Date: 20\d\d\.",''
            $security = $security -replace 'SIGN HERE ',''
            $security = $security -replace ' DD FORM 2875, ',''
            $security = $security -replace "DN: ",''
            $security = $security -replace "US, ",''
            $security = $security -replace "U\.S\. ",''
            $security = $security -replace "DOD, ",''
            $security = $security -replace "=CONTRACTOR",''
            $security = $security -replace "PKI, ",''
            $security = $security -replace "Government, ",''
            $security = $security -replace "c=",''
            $security = $security -replace "o=",''
            $security = $security -replace "ou=",''
            $security = $security -replace "cn=",''
            $security = $security -replace ' \s ',''
            $security = $security -replace ' \S ',''
            $security = $security -replace ',',''
            $security = $security -replace '  *'," "
            $security = $security -match '\.[ -~]*.\d\d\d\d\d\d\d\d\d*\d*' | select -first 1
            <#
            [regex]$securitysignature = '\w*\s*\S*\w+\.\w+\.\w+\S*\.\d{10}'
            $securitysignature.Matches($security) | foreach-object {$_.Value} | Select-object -First 1 | Out-String -OutVariable secsig
            [regex]$securitytimestamp = 'Date: 20\d\d.\d\d.\d\d \d\d:\d\d:\d\d'
            $securitytimestamp.Matches($security) | foreach-object {$_.Value} | Select-object -First 1 | Out-String -OutVariable secdate
                $security = "$secsig$secdate"
                #>
                                 
Add-Type -AssemblyName System.Windows.Forms

Function FormGUI #create GUI form
{ 
    $AccountPath = @()
	#OUHeads function when creating users who may be placed in one of many different OUs
    ##$OUHeads = (Get-ADOrganizationalUnit -searchbase "OU=Labs-Development,DC=sec,DC=c3sys,DC=army,DC=mil"  -SearchScope OneLevel -Filter *).name
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $SCRIPT:Form                                         = New-Object system.Windows.Forms.Form
    $Form.ClientSize                                     = '539,562'
    $Form.text                                           = "Account Creation"
    $Form.TopMost                                        = $false
    $Form.add_FormClosing({Window-Close-TXT})

   
   function Window-Close-TXT #clean up workingdir if user clicks x button to close form
   {
	Get-Childitem -Path $workingdir -Filter *.pdf | select-object -first 1 | Rename-Item -NewName $oldname -Erroraction 'silentlycontinue'
		Move-Item -Path $workingdir\*.pdf -Destination $storagedir -Erroraction 'silentlycontinue'
   }

    $SCRIPT:FirstNameBox                                 = New-Object system.Windows.Forms.TextBox
    $FirstNameBox.multiline                              = $false
    $FirstNameBox.width                                  = 140
    $FirstNameBox.height                                 = 20
    $FirstNameBox.location                               = New-Object System.Drawing.Point(13,21)
    $FirstNameBox.Font                                   = 'Microsoft Sans Serif,10'
    $FirstNameBox.Text                                   = $first
    $FirstNameBoxLabel                                   = New-Object system.Windows.Forms.Label
    $FirstNameBoxLabel.text                              = "First Name*"
    $FirstNameBoxLabel.AutoSize                          = $true
    $FirstNameBoxLabel.width                             = 25
    $FirstNameBoxLabel.height                            = 10
    $FirstNameBoxLabel.location                          = New-Object System.Drawing.Point(13,45)
    $FirstNameBoxLabel.Font                              = 'Microsoft Sans Serif,10,style=Bold'
    $FirstNameBoxLabel.ForeColor                         = "#d0021b"

    $SCRIPT:MiddleNameBox                                = New-Object system.Windows.Forms.TextBox
    $MiddleNameBox.multiline                             = $false
    $MiddleNameBox.width                                 = 37
    $MiddleNameBox.height                                = 20
    $MiddleNameBox.location                              = New-Object System.Drawing.Point(158,21)
    $MiddleNameBox.Font                                  = 'Microsoft Sans Serif,10'
    $MiddleNameBox.Text                                  = $MI
    $MiddleNameBoxLabel                                  = New-Object system.Windows.Forms.Label
    $MiddleNameBoxLabel.text                             = "MI"
    $MiddleNameBoxLabel.AutoSize                         = $true
    $MiddleNameBoxLabel.width                            = 25
    $MiddleNameBoxLabel.height                           = 10
    $MiddleNameBoxLabel.location                         = New-Object System.Drawing.Point(158,45)
    $MiddleNameBoxLabel.Font                             = 'Microsoft Sans Serif,10,style=Bold'
    $MiddleNameBoxLabel.ForeColor                        = "#d0021b"

    $SCRIPT:LastNameBox                                  = New-Object system.Windows.Forms.TextBox
    $LastNameBox.multiline                               = $false
    $LastNameBox.width                                   = 140
    $LastNameBox.height                                  = 20
    $LastNameBox.location                                = New-Object System.Drawing.Point(200,21)
    $LastNameBox.Font                                    = 'Microsoft Sans Serif,10'
    $LastNameBox.Text                                    = $last
    $LastNameBoxLabel                                    = New-Object system.Windows.Forms.Label
    $LastNameBoxLabel.text                               = "Last Name*"
    $LastNameBoxLabel.AutoSize                           = $true
    $LastNameBoxLabel.width                              = 25
    $LastNameBoxLabel.height                             = 10
    $LastNameBoxLabel.location                           = New-Object System.Drawing.Point(200,45)
    $LastNameBoxLabel.Font                               = 'Microsoft Sans Serif,10,style=Bold'
    $LastNameBoxLabel.ForeColor                          = "#d0021b"

    $SCRIPT:UsernameBox                 = New-Object system.Windows.Forms.TextBox
    $UsernameBox.multiline              = $false
    $UsernameBox.width                  = 180
    $UsernameBox.height                 = 20
    $UsernameBox.location               = New-Object System.Drawing.Point(345,21)
    $UsernameBox.Font                   = 'Microsoft Sans Serif,10'
    $UsernameBox.Text                   = $username
    $UserNameBoxLabel                   = New-Object system.Windows.Forms.Label
    $UserNameBoxLabel.text              = "Username*"
    $UserNameBoxLabel.AutoSize          = $true
    $UserNameBoxLabel.width             = 25
    $UserNameBoxLabel.height            = 10
    $UserNameBoxLabel.location          = New-Object System.Drawing.Point(345,45)
    $UserNameBoxLabel.Font              = 'Microsoft Sans Serif,10,style=Bold'
    $UserNameBoxLabel.ForeColor         = "#d0021b"

    $SCRIPT:EDIPIBox                    = New-Object system.Windows.Forms.TextBox
    $EDIPIBox.multiline                 = $false
    $EDIPIBox.width                     = 85
    $EDIPIBox.height                    = 20
    $EDIPIBox.location                  = New-Object System.Drawing.Point(440,119)
    $EDIPIBox.Font                      = 'Microsoft Sans Serif,10'
    $EDIPIBox.Text                      = $EDIPI
    $EDIPIBoxLabel                      = New-Object system.Windows.Forms.Label
    $EDIPIBoxLabel.text                 = "EDIPI*"
    $EDIPIBoxLabel.AutoSize             = $true
    $EDIPIBoxLabel.width                = 25
    $EDIPIBoxLabel.height               = 10
    $EDIPIBoxLabel.location             = New-Object System.Drawing.Point(445,143)
    $EDIPIBoxLabel.Font                 = 'Microsoft Sans Serif,9,style=Bold'
    $EDIPIBoxLabel.ForeColor            = "#d0021b"

    $SCRIPT:EmailBox                    = New-Object system.Windows.Forms.TextBox
    $EmailBox.multiline                 = $false
    $EmailBox.width                     = 280
    $EmailBox.height                    = 20
    $EmailBox.location                  = New-Object System.Drawing.Point(14,70)
    $EmailBox.Font                      = 'Microsoft Sans Serif,10'
    $EmailBox.Text                      = $email
    $EmailBoxLabel                      = New-Object system.Windows.Forms.Label
    $EmailBoxLabel.text                 = "Email Address*"
    $EmailBoxLabel.AutoSize             = $true
    $EmailBoxLabel.width                = 25
    $EmailBoxLabel.height               = 10
    $EmailBoxLabel.location             = New-Object System.Drawing.Point(14,94)
    $EmailBoxLabel.Font                 = 'Microsoft Sans Serif,10,style=Bold'
    $EmailBoxLabel.ForeColor            = "#d0021b"

    $SCRIPT:PhoneBox                    = New-Object system.Windows.Forms.TextBox
    $PhoneBox.multiline                 = $false
    $PhoneBox.width                     = 134
    $PhoneBox.height                    = 20
    $PhoneBox.location                  = New-Object System.Drawing.Point(300,70)
    $PhoneBox.Font                      = 'Microsoft Sans Serif,10'
    $PhoneBox.Text                      = $phone
    $PhoneBoxLabel                      = New-Object system.Windows.Forms.Label
    $PhoneBoxLabel.text                 = "Phone Number"
    $PhoneBoxLabel.AutoSize             = $true
    $PhoneBoxLabel.width                = 25
    $PhoneBoxLabel.height               = 10
    $PhoneBoxLabel.location             = New-Object System.Drawing.Point(300,94)
    $PhoneBoxLabel.Font                 = 'Microsoft Sans Serif,10,style=Bold'
    $PhoneBoxLabel.ForeColor            = "#d0021b"

    $SCRIPT:BuildingBox                 = New-Object system.Windows.Forms.TextBox
    $BuildingBox.multiline              = $false
    $BuildingBox.width                  = 85
    $BuildingBox.height                 = 20
    $BuildingBox.location               = New-Object System.Drawing.Point(440,70)
    $BuildingBox.Font                   = 'Microsoft Sans Serif,10'
    $BuildingBox.Text                   = $building
    $BuildingBoxLabel                   = New-Object system.Windows.Forms.Label
    $BuildingBoxLabel.text              = "Building"
    $BuildingBoxLabel.AutoSize          = $true
    $BuildingBoxLabel.width             = 25
    $BuildingBoxLabel.height            = 10
    $BuildingBoxLabel.location          = New-Object System.Drawing.Point(440,94)
    $BuildingBoxLabel.Font              = 'Microsoft Sans Serif,10,style=Bold'
    $BuildingBoxLabel.ForeColor         = "#d0021b"

    $SCRIPT:JustificationBox            = New-Object system.Windows.Forms.TextBox
    $JustificationBox.multiline         = $true
    $JustificationBox.width             = 280
    $JustificationBox.height            = 114
    $JustificationBox.location          = New-Object System.Drawing.Point(14,119)
    $JustificationBox.Font              = 'Microsoft Sans Serif,10'
    $JustificationBox.Text              = $justification
    $JustificationBoxLabel              = New-Object system.Windows.Forms.Label
    $JustificationBoxLabel.text         = "Description / Justification"
    $JustificationBoxLabel.AutoSize     = $true
    $JustificationBoxLabel.width        = 25
    $JustificationBoxLabel.height       = 20
    $JustificationBoxLabel.location     = New-Object System.Drawing.Point(14,233)
    $JustificationBoxLabel.Font         = 'Microsoft Sans Serif,10,style=Bold'
    $JustificationBoxLabel.ForeColor    = "#d0021b"

    $SCRIPT:OfficeBox                   = New-Object system.Windows.Forms.TextBox
    $OfficeBox.multiline                = $false
    $OfficeBox.width                    = 134
    $OfficeBox.height                   = 20
    $OfficeBox.location                 = New-Object System.Drawing.Point(300,119)
    $OfficeBox.Font                     = 'Microsoft Sans Serif,10'
    $OfficeBox.Text                     = $office
    $OfficeBoxLabel                     = New-Object system.Windows.Forms.Label
    $OfficeBoxLabel.text                = "Office/ORG"
    $OfficeBoxLabel.AutoSize            = $true
    $OfficeBoxLabel.width               = 25
    $OfficeBoxLabel.height              = 10
    $OfficeBoxLabel.location            = New-Object System.Drawing.Point(300,143)
    $OfficeBoxLabel.Font                = 'Microsoft Sans Serif,10,style=Bold'
    $OfficeBoxLabel.ForeColor           = "#d0021b"

    $SCRIPT:OUSelectBox                 = New-Object system.Windows.Forms.ComboBox
    ##$OUSelectBox.Items                 = $OUHeads | Out-String
    $OUSelectBox.width                  = 225
    $OUSelectBox.height                 = 20
    $OUSelectBox.location               = New-Object System.Drawing.Point(300,208)
    $OUSelectBox.Font                   = 'Microsoft Sans Serif,10'
    $OUSelectBox.Text                   = 'Automatically Created Users'
    ##ForEach ($OU in $OUHeads){$OUSelectBox.Items.Add($OU)}

    $OUSelectBoxLabel                   = New-Object system.Windows.Forms.Label
    $OUSelectBoxLabel.text              = "OU Selection*"
    $OUSelectBoxLabel.AutoSize          = $true
    $OUSelectBoxLabel.width             = 25
    $OUSelectBoxLabel.height            = 10
    $OUSelectBoxLabel.location          = New-Object System.Drawing.Point(300,233)
    $OUSelectBoxLabel.Font              = 'Microsoft Sans Serif,10,style=Bold'
    $OUSelectBoxLabel.ForeColor         = "#d0021b"

    $AccountTypeBox                     = New-Object system.Windows.Forms.ListBox
    $AccountTypeBox.width               = 225
    $AccountTypeBox.height              = 35
    $AccountTypeBox.location            = New-Object System.Drawing.Point(300,168)
    $AccountTypeBox.Font                = 'Microsoft Sans Serif,10'
    $AccountTypeBox.Items.AddRange(@("Standard User","System Administrator"))
    $AccountTypeBoxLabel                = New-Object system.Windows.Forms.Label
    $AccountTypeBoxLabel.text           = "Account Type*"
    $AccountTypeBoxLabel.AutoSize       = $true
    $AccountTypeBoxLabel.width          = 25
    $AccountTypeBoxLabel.height         = 10
    $AccountTypeBoxLabel.location       = New-Object System.Drawing.Point(300,188)
    $AccountTypeBoxLabel.Font           = 'Microsoft Sans Serif,10,style=Bold'
    $AccountTypeBoxLabel.ForeColor      = "#d0021b"

	#signatures
    $UserSignatureLabel                        = New-Object system.Windows.Forms.Textbox
    $UserSignatureLabel.multiline              = $true
    $UserSignatureLabel.text                   = $user
    $UserSignatureLabel.AutoSize               = $true
    $UserSignatureLabel.width                  = 511
    $UserSignatureLabel.height                 = 40
    $UserSignatureLabel.location               = New-Object System.Drawing.Point(14,262)
    $UserSignatureLabel.Font                   = 'Microsoft Sans Serif,10'
    $UserSignatureLabel.ForeColor              = "#000000"
	$UserSignatureLabel.ReadOnly               = $true
	$UserSignatureLabelHeader                  = New-Object system.Windows.Forms.Label
    $UserSignatureLabelHeader.text             = '^^^ User' + ' ' + 'Signature ^^^'
    $UserSignatureLabelHeader.AutoSize         = $true
    $UserSignatureLabelHeader.width            = 25
    $UserSignatureLabelHeader.height           = 10
    $UserSignatureLabelHeader.location         = New-Object System.Drawing.Point(14,302)
    $UserSignatureLabelHeader.Font             = 'Microsoft Sans Serif,10,style=Bold'
    $UserSignatureLabelHeader.ForeColor        = "#000000"
    
    $SupervisorSignatureLabel                  = New-Object system.Windows.Forms.Textbox
    $SupervisorSignatureLabel.multiline        = $true
    $SupervisorSignatureLabel.text             = $supervisor
    $SupervisorSignatureLabel.AutoSize         = $true
    $SupervisorSignatureLabel.width            = 511
    $SupervisorSignatureLabel.height           = 40
    $SupervisorSignatureLabel.location         = New-Object System.Drawing.Point(14,322)
    $SupervisorSignatureLabel.Font             = 'Microsoft Sans Serif,10'
    $SupervisorSignatureLabel.ForeColor        = "#000000"
	$SupervisorSignatureLabel.ReadOnly         = $true
	$SupervisorSignatureLabelHeader            = New-Object system.Windows.Forms.Label
    $SupervisorSignatureLabelHeader.text       = '^^^ Supervisor' + ' ' + 'Signature ^^^'
    $SupervisorSignatureLabelHeader.AutoSize   = $true
    $SupervisorSignatureLabelHeader.width      = 25
    $SupervisorSignatureLabelHeader.height     = 10
    $SupervisorSignatureLabelHeader.location   = New-Object System.Drawing.Point(14,362)
    $SupervisorSignatureLabelHeader.Font       = 'Microsoft Sans Serif,10,style=Bold'
    $SupervisorSignatureLabelHeader.ForeColor  = "#000000"

    $IAOSignatureLabel                         = New-Object system.Windows.Forms.Textbox
    $IAOSignatureLabel.multiline               = $true
    $IAOSignatureLabel.text                    = $IAO
    $IAOSignatureLabel.AutoSize                = $true
    $IAOSignatureLabel.width                   = 511
    $IAOSignatureLabel.height                  = 40
    $IAOSignatureLabel.location                = New-Object System.Drawing.Point(14,382)
    $IAOSignatureLabel.Font                    = 'Microsoft Sans Serif,10'
    $IAOSignatureLabel.ForeColor               = "#000000"
	$IAOSignatureLabel.ReadOnly                = $true
	$IAOSignatureLabelHeader                   = New-Object system.Windows.Forms.Label
    $IAOSignatureLabelHeader.text              = '^^^ IAO' + ' ' + 'Signature ^^^'
    $IAOSignatureLabelHeader.AutoSize          = $true
    $IAOSignatureLabelHeader.width             = 25
    $IAOSignatureLabelHeader.height            = 10
    $IAOSignatureLabelHeader.location          = New-Object System.Drawing.Point(14,422)
    $IAOSignatureLabelHeader.Font              = 'Microsoft Sans Serif,10,style=Bold'
    $IAOSignatureLabelHeader.ForeColor         = "#000000"

    $SecuritySignatureLabel                    = New-Object system.Windows.Forms.Textbox
    $SecuritySignatureLabel.multiline          = $true
    $SecuritySignatureLabel.text               = $security
    $SecuritySignatureLabel.AutoSize           = $true
    $SecuritySignatureLabel.width              = 511
    $SecuritySignatureLabel.height             = 40
    $SecuritySignatureLabel.location           = New-Object System.Drawing.Point(14,442)
    $SecuritySignatureLabel.Font               = 'Microsoft Sans Serif,10'
    $SecuritySignatureLabel.ForeColor          = "#000000"
    $SecuritySignatureLabel.ReadOnly           = $true   
    $SecuritySignatureLabelHeader              = New-Object system.Windows.Forms.Label
    $SecuritySignatureLabelHeader.text         = '^^^ Security' + ' ' + 'Signature ^^^'
    $SecuritySignatureLabelHeader.AutoSize     = $true
    $SecuritySignatureLabelHeader.width        = 25
    $SecuritySignatureLabelHeader.height       = 10
    $SecuritySignatureLabelHeader.location     = New-Object System.Drawing.Point(14,482)
    $SecuritySignatureLabelHeader.Font         = 'Microsoft Sans Serif,10,style=Bold'
    $SecuritySignatureLabelHeader.ForeColor    = "#000000"	

    $CreateAccount                          = New-Object system.Windows.Forms.Button
    $CreateAccount.text                     = "Create User"
    $CreateAccount.width                    = 180
    $CreateAccount.height                   = 60
    $CreateAccount.location                 = New-Object System.Drawing.Point(345,492)
    $CreateAccount.Font                     = 'Arial,15,style=Bold'
    $CreateAccount.Add_Click({CheckUser})

    $ViewText                               = New-Object system.Windows.Forms.Button
    $ViewText.text                          = "View Text"
    $ViewText.width                         = 100
    $ViewText.height                        = 30
    $ViewText.location                      = New-Object System.Drawing.Point(14,522)
    $ViewText.Font                          = 'Arial,10,style=Bold'
    $ViewText.Add_Click({ViewText})
	
    $ViewPDF                               = New-Object system.Windows.Forms.Button
    $ViewPDF.text                          = "View PDF"
    $ViewPDF.width                         = 100
    $ViewPDF.height                        = 30
    $ViewPDF.location                      = New-Object System.Drawing.Point(134,522)
    $ViewPDF.Font                          = 'Arial,10,style=Bold'
    $ViewPDF.Add_Click({ViewPDF})

    $Form.controls.AddRange(@($ViewText,$ViewPDF,$FirstNameBox,$MiddleNameBox,$LastNameBox,$UsernameBox,$EDIPIBox,$FirstNameBoxLabel,$MiddleNameBoxLabel,$LastNameBoxLabel,$UserNameBoxLabel,$EDIPIBoxLabel,$EmailBox,$EmailBoxLabel,$PhoneBox,$PhoneBoxLabel,$BuildingBox,$BuildingBoxLabel,$JustificationBox,$JustificationBoxLabel,$OfficeBox,$OfficeBoxLabel,$OUSelectBox,$AccountTypeBox,$CreateAccount,$AccountTypeBoxLabel,$OUSelectBoxLabel,$SupervisorSignatureLabel,$SupervisorSignatureLabelHeader,$IAOSignatureLabel,$IAOSignatureLabelHeader,$UserSignatureLabel,$UserSignatureLabelHeader,$SecuritySignatureLabel,$SecuritySignatureBoxLabelHeader))
    [void]$Form.ShowDialog()
}
Function CheckUser #ensure text fields have the proper characteristics
{ 
    $wshell = New-Object -ComObject Wscript.Shell

    If ($FirstNameBox.text -eq "")                {$wshell.Popup("FIRST NAME REQUIRED!",0,"ERROR",0x10);RETURN}
    If ($LastNameBox.Text -eq "")                 {$wshell.Popup("LAST NAME REQUIRED!",0,"ERROR",0x10);RETURN}
    If ($UsernameBox.Text -eq "")                 {$wshell.Popup("USERNAME REQUIRED!",0,"ERROR",0x10);RETURN}
        $UsernameLength = $UsernameBox.Text | Measure-Object -Character
            $Count = $UsernameLength.Characters
            If ($Count -gt 20)                    {$wshell.Popup("USERNAME TOO LONG!",0,"ERROR",0x10);RETURN} #username cannot be >20 characters
    If ($EDIPIBox.Text -eq "")                    {$wshell.Popup("EDIPI REQUIRED!",0,"ERROR",0x10);RETURN}
    If ($EmailBox.Text -eq "")                    {$wshell.Popup("EMAIL REQUIRED!",0,"ERROR",0x10);RETURN}
    If ($PhoneBox.Text -eq "")                    {$wshell.Popup("PHONE NUMBER REQUIRED!",0,"ERROR",0x10);RETURN}
    If ($PhoneBox.Text -notlike "*.*.*")          {$wshell.Popup("PHONE NUMBER - USE FORMAT XXX.XXX.XXXX",0,"ERROR",0x10);RETURN}
    If ($BuildingBox.Text -eq "")                 {$wshell.Popup("BUILDING# REQUIRED!",0,"ERROR",0x10);RETURN}
    if(Get-aduser $UsernameBox.Text)              {$wshell.Popup("USERNAME Already in use!",0,"ERROR",0x10)}
    if(Get-ADUser -filter {(UserPrincipalName -eq $EDIPI) -or (UserPrincipalName -eq $EDIPI2)})       {$wshell.Popup("EDIPI Already in use!",0,"ERROR",0x10);RETURN} 
    #if ($OUSelectBox.SelectedItem -eq $null)         {$wshell.Popup("OU - REQUIRED!",0,"ERROR",0x10);RETURN}
   
   $SCRIPT:EDIPI = $EDIPIBox.Text+"@Mil"
   $SCRIPT:EDIPI2 = $EDIPIBox.Text+"@sec.c3sys.army.mil"
      
   AssignValues
}
Function AssignValues
{
    #reassign all auto-filled form variables to currently entered form data in case user edited a field
    $First = $FirstNameBox.Text 
    $MI = $MiddleNameBox.Text 
    $Last = $LastNameBox.Text
    $Username = $UsernameBox.Text
    $EDIPI = $EDIPIBox.Text

    $Email = $EmailBox.Text
    $Phone = $PhoneBox.Text

    $Office = $OfficeBox.Text
    $Building = $BuildingBox.Text
    $Room = $RoomBox.Text

    $Justification = $JustificationBox.Text
    
    $AccountType = $AccountTypeBox.selectedItem
    $OU = $OUSelectBox.SelectedItem

    CreateAccount
}
Function CreateAccount
{

$ADPassword = '12345!@#$qwertQWERT'
$ADPassword = ConvertTo-SecureString -Asplaintext -Force -String $ADPassword

#create user's account with specified parameters in AD 
New-ADUser -AccountPassword $ADPassword -ChangePasswordAtLogon 1 -Enabled 1 -PasswordNeverExpires $false -SmartcardLogonRequired $true -Office $Building -Organization 'CECOM SEC' -Division 'APG' -Department $office -Displayname $Username -UserPrincipalName $EDIPI@sec.c3sys.army.mil -Name $Username -EmailAddress $Email -OfficePhone $Phone -Givenname $First -Initials $Mi -Surname $Last -Description $Office -StreetAddress ($Building + ' ' + 'combat drive') -Path 'OU=Automatically-Created-Users,OU=Users,OU=6002-D5120,OU=Labs-Development,DC=sec,DC=c3sys,DC=army,DC=mil' 

#disable account by default once it's created
Disable-ADAccount -Identity $Username

    Move-Item -Path $workingdir\*.pdf -Destination $storagedir #move 2875 to secondary location for archiving
    Remove-Item -Path $workingdir\*output.txt #delete temporary 2875 file output text file

$Form.Close()
}
Function ViewText
{
Start-Process -FilePath $workingdir\DD2875-output.txt
}
Function ViewPDF
{
Start-Process -FilePath $workingdir\*.pdf
}

FormGUI