######################################################
# This has to be run by the affected user themselves #
######################################################

# This Script does the following:

    # gets information about the user and  takes their email
    # checks the "current user / my" path for expiring certificates (timeframe => $daystocheck)
    # if there is nothing coming back, do nothing, script is done
    # if there actually is a certificate that expires in the set timeframe, alert the user

# Made by David S. 12.08.22
# Version 1.0

###############################
# Use-case specific variables #
###############################

# How many days in advance do we want to check?
$daystocheck = 21

################################
# Code, nothing to change here #
################################

# Get Usermail by reading the Outlook value
#$usermail = [System.Text.Encoding]::Unicode.GetString((Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002")."Account Name")


$searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
$usermail = $searcher.FindOne().Properties.mail

#Write-Output("E-Mail:" + $usermail) # Debug only

$test = Get-ChildItem -Path Cert:\CurrentUser\My -ExpiringInDays $daystocheck | where{$_.Subject -like "*$usermail*"} | Select-Object @{n=’expire’;e={($_.notafter – (Get-Date)).Days}}
$expirdate = $($test.expire)
if($test -eq $null)
{
# Certs are still good, do nothing
}
else{
# There is a cert for your E-Mail that is expiring
[System.Windows.MessageBox]::Show("Das Zertifikat für die folgende E-Mail läuft in $expirdate Tage(n) ab.`r`n$usermail `r`nBitte melde dich rechtzeitig bei der IT.", "E-Mail Zertifikat läuft bald aus", 0, 48)
}
# Remove variable to not have false positives when running this again
Remove-Variable -Name "test"