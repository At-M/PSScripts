# Update registry to put second screen as screen for "notes" in powerpoint
function updateofficereg() {

    $path = "HKCU:\SOFTWARE\Microsoft\Office\16.0\powerpoint\options"
    if (Test-Path $path) {
    $value = "\\\\.\\DISPLAY2"
    # Write to registry
  
      # Check if path exists
      if (!(Test-Path $path)) {
        # If not, create it
        New-Item -Path $path -Force | Out-Null
  
        New-ItemProperty -Path $path -Name $keyname -Value $value
      }
  
      else {
        # if it does, update it
        Set-ItemProperty -Path $path -Name $keyname -Value $value
      }
    }
  }

  updateofficereg