##########################################
# This HAS be run with normal privileges #
##########################################

# This Script does the following:

## Removes MS Teams from the current Userprofile, including all old folders
## Reinstall & Starts MS Teams from the Machine Wide Installer


# Made by David S. 21.07.23

# updated for public use 25.08.23 ~12:10 by At-M

###############################
# Use-case specific variables #
###############################

# Where to find the Machinewide installer
$TeamsMachWideExePath = "C:\Program Files (x86)\Teams Installer\Teams.exe"

################################
#       PUBLIC-ONLY INFO       #
################################

## v1.2: 

# logging functions had to be removed for now, most of it was replaced with Write-Output
# Outlook has to be closed first to run it, since one of the folders messes with Outlook aswell (TeamsPresenceAddIn)
# This modified script has been tested once (25.08.23 12:07), it worked but it's not the prettiest

# To run this as a user, the easiest way would be to create a .bat with the following content that the user can just open and run
# start powershell.exe -executionpolicy bypass -file "path\to\this\psfile.ps1"

################################
# Code, nothing to change here #
################################
$progname = "MS Teams Reinstaller"
$ver = "1.2"

function Remove_TeamsItself () {
    Write-Output "Starting Remove_TeamsItself"
        
    Write-Output "Please wait until Teams opens up again.."

    $TeamsUpdateExePath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams', 'Update.exe')
    $TeamsPath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams')
    $teamsfolderone = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'TeamsMeetingAddin')
    $teamsfoldertwo = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'TeamsPresenceAddin')
    $teamsfolderthree = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft', 'Teams')
    $teamsfolderfour = [System.IO.Path]::Combine($env:APPDATA, 'Teams')
    $removedfolders = 0

    try {
        if (Test-Path -Path $TeamsUpdateExePath) {
            Write-Output "Deinstalling MS Teams Application.."

            # Uninstall app
            $proc = Start-Process -FilePath $TeamsUpdateExePath -ArgumentList "-uninstall -s" -PassThru
            $proc.WaitForExit()
        }
        else {
            Write-Output "The MS Teams application is not installed or could not be found..."
        }
        Write-Output "Removing MS Teams related folders.."
        if (Test-Path -Path $TeamsPath) {
                
            Remove-Item -Path $TeamsPath -Recurse
            $removedfolders++
        }
        if (Test-Path -Path $teamsfolderone) {
            Remove-Item -LiteralPath $teamsfolderone -Force -Recurse
            $removedfolders++
        }
        else {}
        if (Test-Path -Path $teamsfoldertwo) {
            Remove-Item -LiteralPath $teamsfoldertwo -Force -Recurse
            $removedfolders++
        }
        else {}
        if (Test-Path -Path $teamsfolderthree) {
            Remove-Item -LiteralPath $teamsfolderthree -Force -Recurse
            $removedfolders++
        }
        else {}
        if (Test-Path -Path $teamsfolderfour) {
            Remove-Item -LiteralPath $teamsfolderfour -Force -Recurse
            $removedfolders++
        }
        else {}
        Write-Output "$removedfolders MS Teams folder were removed."
    }
    catch {
        Write-Error -ErrorRecord $_
        Write-Output "$_"
        exit /b 1
    }
}

# Output GUI

Write-Host(" ")
Write-Host(" ")
Write-Host("#v$ver########################################################")
Write-Host("#############################################################")
Write-Host("#############################################################")
Write-Host("|                                                           |")
Write-Host("|                   $progname                    |")
Write-Host("|                                                           |")
Write-Host("|         This Window will close itself when done.          |")
Write-Host("#############################################################")
Write-Host("#############################################################")
Write-Host("###########################################################DS")
Write-Host(" ")
Write-Host(" ")


Add-Type -AssemblyName System.Windows.Forms
$mb = [System.Windows.Forms.MessageBox]
$mbIcon = [System.Windows.Forms.MessageBoxIcon]
$mbBtn = [System.Windows.Forms.MessageBoxButtons]

[System.Windows.Forms.Application]::EnableVisualStyles()

Write-Output "Showing Outlook-Alert Window.."
$result = $mb::Show("Has Outlook already been completely closed?", "$progname - Alert", $mbBtn::YesNo, $mbIcon::Warning)
if ($result -eq "No") {
    # Stop code with exit message
    Write-Output "Outlook-Alert -> NO"
    $mb::Show("Please close all Outlook windows first, cancelling..", "$progname - Stopping program..", $mbBtn::OK, $mbIcon::Error) | Out-Null
    Write-Output "Stopping.."
    break
} 
else {
    Write-Output "Outlook-Alert -> YES"

    # Remove Teams from Userprofile
    Remove_TeamsItself

    # Try to start the MachineWide Installer to open up Teams again
    Try {

        $dateA = ((Get-Item $TeamsMachWideExePath).LastWriteTime)

        # Format the shown date to a date that does make sense
        $dateA = ($dateA).ToString("dd.MM.yyyy")

        Write-Output "The Teams Installer was last updated on $dateA and is now starting.."
        Start-Process -FilePath $TeamsMachWideExePath -ArgumentList 'OPTIONS="noAutoStart=false"'

    }
    Catch {
        Write-Output "Teams Installer could not be found - and therefore could not be started."
        Write-Output "If nothing happens even after logging off/on once, please notify IT via ticket and mention the server/client name $env:COMPUTERNAME."
    }
}
Write-Output "$progname done."