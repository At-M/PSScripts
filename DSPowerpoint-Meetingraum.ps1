# Update registry to put second screen as screen for "notes" in powerpoint
function updateofficereg() {
    $path = "HKCU:\SOFTWARE\Microsoft\Office\16.0\powerpoint\options"

    if (Test-Path $path) {
        $value = "\\\\.\\DISPLAY2"
        $keyname = "DisplayMonitor"
        $keynametwo = "UseAutoMonSelection"
        $valuetwo = "0"

        ## write to registry
  
        # check if path exists
        # of not, create it
        if (!(Test-Path $path)) {
            New-Item -Path $path -Force | Out-Null
            # Set second screen as screen for notes
            New-ItemProperty -Path $path -Name $keyname -Value $value
            # To set Auto Monitor Selection to 0
            Set-ItemProperty -Path $path -Name $keynametwo -Value $valuetwo
        }
        # if it does, update it
        else {
            # set second screen as screen for notes
            Set-ItemProperty -Path $path -Name $keyname -Value $value
            # to set auto monitor selection to 0
            Set-ItemProperty -Path $path -Name $keynametwo -Value $valuetwo
        }
    }
}
# run above function
updateofficereg