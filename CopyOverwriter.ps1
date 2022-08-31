##########################################
# This can be run with normal privileges #
##########################################

# This Script does the following:

## checks if the original file exists
## checks if there is already a file and confirms with the user (or the variable set below) if overwriting is allowed
## copies the original file to the new location
## sets new file to read-only if possible and wanted

# Made by David S. 30.03.22
# Version 1.0

###############################
# Use-case specific variables #
###############################

# Filename
$filename = "filename.txt"
# Path to original File
$originalfile = "\\server\path\to\file\"
# Path to new File
$newfile = "Driveletter:\Path\to\store\new\file\"

# Is it ok to automatically overwrite the new file, if it exists? (1 = yes, 0 = no)
$overwrite = 0

# Should the new file be read-only?  (1 = yes, 0 = no)
$readonly = 1

################################
# Code, nothing to change here #
################################
Add-Type -AssemblyName System.Windows.Forms

$oldfilepath = $originalfile + $filename
$newfilepath = $newfile + $filename
# 0 = Not done at all, 1 = Done writing, 2= done writing, but with errors, 3 = errored, 4= completely done
$finished = 0

Write-Output "#########"
Write-Output "# Start #"
Write-Output "#########"

$testoriginal = Test-Path -Path $oldfilepath
if ($testoriginal) {
    Write-Output "Original file has been found, continuing..."

    $testnew = Test-Path -Path $newfile
    if ($testnew) {

        Write-Output "File already exists in new folder, trying to overwrite..."
        if ($overwrite) {
            # Overwrite this stuff
            Copy-Item -Path $oldfilepath -Destination $newfilepath -Force
            $finished = 1
        }
        else {

            $msgBoxInput = [System.Windows.Forms.MessageBox]::Show('Es existiert bereits eine Datei im Zielordner, soll diese überschrieben werden?', 'Dateivorgang', '4', 'Question')

            If ($msgboxInput -eq "Yes") {
                Write-Output "Overwriting manually allowed, trying to overwrite..."
                Copy-Item -Path $oldfilepath -Destination $newfilepath -Force
                $finished = 1
            }
            else {
                Write-Output "Overwriting not allowed although file exists, stopping..."
                $finished = 3
            }
    
        }
    }
    else {
        Write-Output "File does not exist yet, writing..."
        Copy-Item -Path $oldfilepath -Destination $newfilepath
        $finished = 1
    }

    # File has been written, now make it read-only

    if (($finished -eq 1) -And ($readonly -eq 1)) {
        $file = Get-Item -Path $newfilepath
        $file.IsReadOnly = $true
        $checkreadonly = $file.IsReadOnly
        if ($checkreadonly) {
            Write-Output "New file has been written is now read only."
            $finished = 4
        }
        else {
            Write-Output "Could not change new file to read only."
            $finished = 2
        }
    }
    else {
        Write-Output "Either the file has not been written yet, or read-only is disabled. Stopping..."
        $finished = 3
    }

}
else {
    Write-Output "Original file has NOT been found, stopping."
    $finished = 3
}
switch ($finished) {
    0 { Write-Output "Script somehow aborted, this case should not be happening." }
    1 { Write-Output "File has been copied, nothing else." }
    2 { Write-Output "File has been copied with errors afterwards." }
    3 { Write-Output "An error has occurred, stopped." }
    4 { Write-Output "Script ran successfully, the file should now be copied and read-only." }
    default { Write-Output "Script somehow aborted, this case should not be happening." }
}