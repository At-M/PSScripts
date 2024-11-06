Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine
######################################################
# This has to be run by the affected user themselves #
######################################################
# This is a script that generates a signature for someone that uses Office 365

# This Script does the following:

# get the lowest office version installed on the current system
# uses that information to get the users mail out of the registry
# read the Active Directory to get their Full Name, Department, Phone Extension, Fax Number and ShortenedUsername ($userfolder)
# insert the userspecific variables in two signature templates (standard + short)
# save said templates to both htm and text, and copy over the needed image/themefiles from a specific path
# edit the registry so that the standard signature will be used by default, short for answering other mails

# Made by David S. 13.05.22, adapted somewhen late 23 / early 24
# Version 1.0

###############################
# Use-case specific variables #
###############################

# "Destination" Office Version has to be manually set to the currently newest version (Office 2019 -> 16)
# DONT CHANGE, this is a remnant of me trying to adapt this to office 365 (used 2013 before) and i've just opted to fully make this o365 only
$officever = "16"

# Sourcepath for signaturespecific theme-files can be found in DSFunctions, near the end of function "savesignature", around line 160

# If for some Reason the ActiveDirectory Functions wont work, uncomment the following line
#import-module ActiveDirectory



################################
<# Known errors / todo:
- When the Username is wildly different from the Users folder, this script might not find the correct path to put the signature files in
#>
################################
################################
# Code, nothing to change here #
################################

# Import Functions made by David S.
Import-Module "\files\o365\DSFunctions.psm1"


## Messagebox for user interaction
Add-Type -AssemblyName System.Windows.Forms
$mb = [System.Windows.Forms.MessageBox]
$mbIcon = [System.Windows.Forms.MessageBoxIcon]
$mbBtn = [System.Windows.Forms.MessageBoxButtons]

[System.Windows.Forms.Application]::EnableVisualStyles()

# Find out if script is called from somewhere else
if (-not $MyInvocation.PSCommandPath) {
  # Script was called from the PowerShell prompt or via the PowerShell CLI.
  $result = $mb::Show("Has Outlook been opened atleast once and is currently closed?", "Signaturegenerator - Alert", $mbBtn::YesNo, $mbIcon::Warning)
}
else {
  # Script was called from the script whose path is reflected in
  # $MyInvocation.PSCommandPath
  $result = "Yes"
}



if ($result -eq "No") {
  # Stop code with exit message
  $mb::Show("Please open Outlook atleast once, and then run this after it has been closed again.", "Signaturegenerator - Stopping Program..", $mbBtn::OK, $mbIcon::Error) | Out-Null
  break
} 
else {
  $mb::Show("Please wait, do not start outlook before the next infobox comes up..", "Signaturegenerator - Info..", $mbBtn::OK, $mbIcon::Information) | Out-Null

  # Get Usermail via ADuser
  $usermail = Get-ADUser -Filter "SamAccountName -eq '$env:USERNAME'" -Properties EmailAddress | Select-Object -ExpandProperty EmailAddress
  $o365mail = Get-ADUser -Filter "SamAccountName -eq '$env:USERNAME'" -Properties EmailAddress | Select-Object -ExpandProperty EmailAddress

  Write-Output("E-Mail:" + $usermail) # Useful for Debug only

  #create other queries depending on the email
  #$userdepartment = GetADPropertybyMail $usermail "Department" #now called position
  $usertitle = GetADPropertybyMail $usermail "Title" #now called position
  $name = GetADPropertybyMail $usermail "GivenName"
  $surname = GetADPropertybyMail $usermail "SurName"
  $usertel = GetADPropertybyMail $usermail "OfficePhone"
  $userfax = GetADPropertybyMail $usermail "FacsimileTelephoneNumber"
  $userfolder = GetADPropertybyMail $usermail "SamAccountName"

  <# # If specific Coworkers that married have different foldernames than their samaccountname, quickfix to work around that problem
if($userfolder -eq "JohnD"){
  $userfolder = "JohnE"
}
else{
  if($userfolder -eq "JaneD"){
    $userfolder = "JaneE"
  }
  else{
    
  }
}
#>

  Write-Output("Title:" + $usertitle) # Useful for Debug only
  Write-Output("Name:" + $name) # Useful for Debug only
  Write-Output("Surname:" + $surname) # Useful for Debug only
  Write-Output("Tel.:" + $usertel) # Useful for Debug only
  Write-Output("SAMAcc:" + $userfolder) # Useful for Debug only

  $username = $name + " " + $surname

  #### STANDARD START
  $signaturename = "Standard"

  ### TXT
  $signone = "Copy TXT Contents here"
  #Name will be inserted here
  #Department will be inserted here
  $signtwo = "Copy TXT Contents here"
  #Extension will be inserted here
  $signthree = "Copy TXT Contents here"
  #Faxno. will be inserted here
  #Usermail will be inserted here
  $signfour = "Copy TXT Contents here"
  $signfive = "Copy TXT Contents here"
  $signsix = "Copy TXT Contents here"
  $signseven = "Copy TXT Contents here"
  # can't remember what these were for to be honest, sorry. i think it was because of some viewing issue with our spacings
  $utspaced = "  " + $usertitle
  $umspaced = "  " + $usermail
  $unspaced = "  " + $username
  # Save the signature to the specific file and also create registry entry 0 = TXT, 1 = HTML
  savesignature $signaturename $umspaced $unspaced $utspaced $usertel $userfax $userfolder $signone $signtwo $signthree $signfour $signfive $signsix $signseven "0" $officever
  ### TXT END

  ### HTM START
<# 

signature path:

C:\Users\username\AppData\Roaming\Microsoft\Signatures

Find the htm/ txt and copy the contents in the variables below/above.
the files in the signature folders with similar names to their signatures just need to be be made to a .zip with the same name as the folder
some things might need to be changed

some lines that have to be changed to look like this:
(this refers to your zip file names i think)

<link rel=File-List href=`"Standard (" + $o365mail + ")-files/filelist.xml`">
<link rel=Edit-Time-Data href=`"Standard (" + $o365mail + ")-Dateien/editdata.mso`">
<link rel=themeData href=`"Standard (" + $o365mail + ")-Dateien/themedata.thmx`">
<link rel=colorSchemeMapping href=`"Standard (" + $o365mail + ")-Dateien/colorschememapping.xml`">

if there are pictures in the mail, you need to change the path for them aswell, for example with "image001.png":

src=`"Standard (" + $o365mail + ")-Dateien/image001.png`"

if there's german umlauts in your signature, you might need to replace them, for example:

"Mit freundlichen Grüßen / kind regards" becomes:
 <p class=MsoNormal><span style='color:#404040'>Mit freundlichen Gr"+ [char]0x00FC + [char]0x00DF + "en / kind
  regards<br>
#>

  $signone = "Copy HTM Contents here"
  #UserName
  $signtwo = "Copy HTM Contents here"
  #UserDepartment
  $signthree = "Copy HTM Contents here"
  #UserTel
  $signfour = "Copy HTM Contents here"
  #UserFax
  $signfive = "Copy HTM Contents here"
  #UserMail
  $signsix = "Copy HTM Contents here"
  #UserMail again
  $signseven = "Copy HTM Contents here"

  # Save the signature to the specific file and also create registry entry
  savesignature $signaturename $usermail $username $usertitle $usertel $userfax $userfolder $signone $signtwo $signthree $signfour $signfive $signsix $signseven "1" $officever
  #### STANDARD END

### Now we do basically the same, but for the short version of our signature, i left some comments here without explanation because it's similar to above

  #### SHORT START
  $signaturename = "Short"

  ### TXT
  $signone = "Copy TXT Contents here"
  #username
  #department
  $signtwo = "Copy TXT Contents here"
  #extension
  #usermail
  $signthree = "Copy TXT Contents here"
  $signfour = "Copy TXT Contents here"
  $signfive = "Copy TXT Contents here"
  $signsix = "Copy TXT Contents here"
  $signseven = "Copy TXT Contents here"
  $utspaced = "  " + $usertitle
  $umspaced = "  " + $usermail
  # Save the signature to the specific file and also create registry entry
  savesignature $signaturename $umspaced $username $utspaced $usertel $userfax $userfolder $signone $signtwo $signthree $signfour $signfive $signsix $signseven "0" $officever
  ### TXT END
  ### HTM Start
  $signone = "Copy HTM Contents here"
  #Usermail
  $signtwo = "Copy HTM Contents here"
  #Username
  $signthree = "Copy HTM Contents here"
  #Department
  $signfour = "Copy HTM Contents here"
  #Usertel
  $signfive = "Copy HTM Contents here"
  ### HTM END
  # Save the signature to the specific file and also create registry entry
  savesignature $signaturename $usermail $username $usertitle $usertel $userfax $userfolder $signone $signtwo $signthree $signfour $signfive $signsix $signseven "1" $officever

  $mb::Show("Outlook can now be opened again.", "Signaturegenerator - Done..", $mbBtn::OK, $mbIcon::Information) | Out-Null
  #### SHORT END
}
# Remove Function Module to be able to run this in VSCode without always restarting the console
Remove-Module DSFunctions
Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope LocalMachine
