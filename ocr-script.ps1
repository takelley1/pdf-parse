$workingdir = c:\users\waldo\desktop\ad-auto-account-creation\2875 #define shortcut for main directory

New-Item -Path $workingdir\working2875 -ItemType directory #create temp directory for 2875s currently being parsed

Get-Childitem -Path $env:USERPROFILE\Downloads -Filter *2875*.pdf | Sort-Object LastAccessTime -Descending | Select-Object -First 1 |
Move-Item -Destination $workingdir\working2875 #move most recently downloaded 2875 to temp directory for parsing

$item = Get-Childitem -Path $workingdir\working2875 
$oldname = $item.Name #store 2875's original name in variable since it must be renamed for pdftotext utility

Get-Childitem -Path $workingdir\working2875 | Rename-Item -Newname {"DD2875-input" + $_.extension} #rename 2875 so pdftotext can read it

Start-Process -FilePath $workingdir\pdftopngbatch.bat -Wait #run pdftotext

Get-Childitem -Path $workingdir\working2875 -Filter *input* | Rename-Item -Newname $oldname #rename 2875 back to its original name

#parse here

Move-Item -Path $workingdir\working2875\*.pdf -Destination $workingdir\old2875s #move 2875 to archiving location

Remove-Item -Path $workingdir\working2875 #delete temp directory







