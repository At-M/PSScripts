######################################################
# This has to be run by the affected user themselves #
######################################################

# This Script does the following:

    # trys to ping the server anhdc (which should be available at all times while in our network)
    # if not possible to reach (no internet or no active FortiClient connection), waits for 15 minutes and tries again (max. 3 times by default)
    # runs gpupdate

# Made by David S. 31.08.22
# Version 1.0

###############################
# Use-case specific variables #
###############################

# How many seconds do we want to wait to try again? (Default: 900)
$secstowait = 900
# How often do we want to try it with running this once? (Default: 3)
$maxtries = 3

################################
# Code, nothing to change here #
################################

function testping {
    # Ping the DC, cause that should always be up
    $ping = Test-Connection -ComputerName anhdc -Quiet
    foreach ($p in $ping) {
        if ($p -eq $false) {
            # If the ping is unsuccessful, return wait time for next try
            Write-Host "Firmeninternes Netz nicht erreichbar! - ist das Gerät mit dem Internet verbunden und FortiClient aktiv?"
            return $secstowait
        }
        else { 
            # Ping is successful, return 0 to not wait
            Write-Host "Firmeninternes Netz erreichbar! Führe fort..."
            return 0
        }
    }
}
$counter = 1 # to reset the counter
$countstop = $maxtries + 1 # for defining the end of the loop
$countskip = $countstop + 1 # for skipping out of the loop

Write-Host "Versuche die Gruppenrichtlinien zu aktualisieren..."

# Do this max 3 times
do {
    Write-Host "Versuch $counter von $maxtries:"
    $testconn = testping
    #  Sleep for predetermined time
    Start-Sleep -Seconds $testconn
    $counter++
    if ($testconn -eq 0) {
        $counter = $countskip
    }
} while ($counter -lt $countstop)

# Try to gpupdate
if ($counter -eq $countskip) {
    Invoke-Command -ScriptBlock { gpupdate }
    Write-Host "Fertig."
}
else {
    Write-Host "Firmeninternes Netzwerk mehrfach nicht erreichbar, Gruppenrichtlinien-Updateversuch gestoppt."
}