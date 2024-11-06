#######################################################################
# This contains multiple functions that might be handy across scripts #
#                  (Think of it like a weird headerfile)                    #
#######################################################################

# Current features:
# 
    # Get the lowest office version installed and output the versionnumber
    # Test if a folder/file exists and delete it (by force if needed)
    # Search the Active Directory with a mail-input for further fields
    # Outputs a string into a file, not garbling german umlauts
    # Save strings to an outlook-signaturefile, in a very specific way

# Made by David S. 13.05.22
# Version 1.0

###############################
# Use-case specific variables #
###############################

# none yet, this is ment to be used like a library of short codesnippets

################################
# Code, nothing to change here #
################################


# Tests for the lowest Version of MS Office detected in the registry and gives back its version number
function GetLowestOfficeVersion {
  $o2013path = "HKCU:\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002"
  $o2019path = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002"
  
  $o2013test = Test-Path $o2013path
  $o2019test = Test-Path $o2019path
  
  if ($o2013test) {
    return "15"
  }
  elseif ($o2019test) {
    return "16"
  }
}
# Tests recursive if folder/file exists and deletes it (by force if needed), leaving the folder itself though
function TestDelete($testpath, $delpath, $deloldpath, $force) {
  # Check if folder exists
  if (Test-Path $testpath) {
    # Delete Folder Contents
    if ($force -eq "1") {
      Remove-Item "$delpath" -Recurse -Force
    }
    #The folder is purposefully not deleted, it can + will be reused
    else {
      Remove-Item "$delpath"
    }
  }
  else {
    # Files are not created, delete only old version
    Remove-Item "$deloldpath"
  }
}
# Searches the AD for a specific mail and reads the given value from the user that has this mail
function GetADPropertybyMail($mail, $prop) {
  $ret = Get-ADUser -Filter "EmailAddress -eq '$mail'" -Properties $prop | Select-Object -ExpandProperty $prop
  return $ret
}
# Update registry to fit office signature names
#
# Possible keynames: Reply-Forward Signature  New Signature
#
# 05/22 Update changed registry value from REG_BINARY to REG_SZ and now we can just insert Strings
# Before that it used to be like this:
# string has to be a byte, every second hex no. must be NULL + at the end there need to be three NULLS
function updateofficereg($version, $keyname, $value) {
  $path = "HKCU:\SOFTWARE\Microsoft\Office\$version.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002"
  # Check if registry path exists
  if (!(Test-Path $path)) {
    # If not, create it
    New-Item -Path $path -Force | Out-Null

    New-ItemProperty -Path $path -Name $keyname -Type String -Value $value

    Write-Output("The Script had to create the Registry Output $keyname with $value")
  }

  else {
    # if it does, update it instead
    Set-ItemProperty -Path $path -Name $keyname -Type String -Value $value

    Write-Output("The Script updated the Registrysetting $keyname to $value")
  }
}
# Outputs a string into a file, with replacing umlauts accordingly
# thanks to:
# https://stackoverflow.com/questions/38842789/exporting-german-umlauts-with-powershell-3-by-out-file-cmdlet
function GermanOutFile ($str, $file) {
  $umlauts = @(
    @('Ä', [char]0x00C4),
    @('Ö', [char]0x00D6),
    @('Ü', [char]0x00DC),
    @('ä', [char]0x00E4),
    @('ö', [char]0x00F6),
    @('ü', [char]0x00FC),
    @('ß', [char]0x00DF)
  )

  foreach ($umlaut in $umlauts) {
    $str = $str -replace $umlaut[0], $umlaut[1]
  }

  $str | Out-File $file
  return $str
}
# Save signature to file in a specific way
#o365 needs usermail in filename aswell
function savesignature($signaturename, $usermail, $username, $userdepartment, $usertel, $userfax, $userfolder, $signone, $signtwo, $signthree, $signfour, $signfive, $signsix, $signseven, $signver, $officever) {
  $pre_signaturepath = $env:USERPROFILE
  $post_signaturepath = "\AppData\Roaming\Microsoft\Signatures"
  $o365mail = Get-ADUser -Filter "SamAccountName -eq '$env:USERNAME'" -Properties EmailAddress | Select-Object -ExpandProperty EmailAddress

  # Build Signature depending on type and default values
  if ($signver -eq "1") {
    # Put everything together (HTML)
    $signpath = $pre_signaturepath + $post_signaturepath + "\" + $signaturename + " (" + $o365mail + ").htm"
    $oldsignpath = $pre_signaturepath + $post_signaturepath + "\" + $signaturename + ".htm"

    if ($signaturename -eq "Short") {
      $signature = $signone + $usermail + $signtwo + $username + $signthree + $userdepartment + $signfour + $usertel + $signfive
    }
    elseif ($signaturename -eq "Standard") {
      $signature = $signone + $username + $signtwo + $userdepartment + $signthree + $usertel + $signfour + $userfax + $signfive + $usermail + $signsix + $usermail + $signseven
    }

  }
  elseif ( $signver -eq "0") {
    #TXT
    $signpath = $pre_signaturepath + $post_signaturepath + "\" + $signaturename + " (" + $o365mail + ").txt"
    $oldsignpath = $pre_signaturepath + $post_signaturepath + "\" + $signaturename + ".txt"
    # Put everything together (TXT) (the additional r is a newline)
    if ($signaturename -eq "Short") {
      $signature = $signone + $username + "`r" + $userdepartment + "`r" + $signtwo + $usertel + "`r" + $usermail + $signthree 
    }
    elseif ($signaturename -eq "Standard") {
      $signature = $signone + "`r" + $username + "`r" + $userdepartment + "`r" + $signtwo + $usertel + $signthree + $userfax + "`r" + $usermail + "`r" + $signfour
    }
  }

  # Delete old Signaturefile if exists
  TestDelete $signpath $signpath $oldsignpath "1"

  GermanOutFile $signature $signpath > $null

  ######## COPY SIGNATURE FILES

  ## Copy Files from zipped "XYZ-Dateien" folder to needed folder
  ## zipping those beforehand is needed, since it otherwise would be converted into a file (because of some kind of linkage between the htm and the folder)

  $destpath = $pre_signaturepath + $post_signaturepath + "\" + $signaturename + " (" + $o365mail + ")-Dateien\"

  $delpath = $destpath + "*"

  $olddestpath = $pre_signaturepath + $post_signaturepath + "\" + $signaturename + "-Dateien\"
  $olddelpath = $olddestpath + "*"

    $sourcepath = "C:\Signaturegenerator\files\" + $signaturename + "-Dateien.zip"
  # Delete old Signature image files (etc.) if exists, force the deletion (1)
  TestDelete $destpath $delpath $olddelpath "1"
  # Unzip from source to destination
  Expand-Archive $sourcepath -DestinationPath $destpath

  ########## SET SIGNATURE IN REGISTRY
  # Standard Values:
  if ($signaturename -eq "Standard") {
    $keyname = "New Signature"
  }
  elseif ($signaturename -eq "Short") {
    $keyname = "Reply-Forward Signature"
  }
  else {
    $keyname = "New Signature"
  }
  updateofficereg $officever $keyname $signaturename
}