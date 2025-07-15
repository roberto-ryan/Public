function Get-vtsDeviceChange {
  <#
  .SYNOPSIS
  This function continuously monitors for changes in the connected devices.
  
  .DESCRIPTION
  The Get-vtsDeviceChange function uses the Get-PnpDevice cmdlet to continuously scan for devices that are currently connected to the system. It compares the results of two consecutive scans to detect any changes in the connected devices. If a device is connected or removed between the two scans, it will output a message indicating the change.
  
  .PARAMETER None
  This function does not take any parameters.
  
  .EXAMPLE
  PS C:\> Get-vtsDeviceChange
  This command will start the function and continuously monitor for any changes in the connected devices. If a device is connected or removed, it will output a message indicating the change.
  
  .NOTES
  To stop the function, use the keyboard shortcut for stopping a running command in your shell (usually Ctrl+C).
  
  .LINK
  Device Management
  #>
  while ($true) {
    $FirstScan = Get-PnpDevice | Where-Object Present -eq True | Select-Object -expand FriendlyName

    Start-Sleep -Seconds 1

    $SecondScan = Get-PnpDevice | Where-Object Present -eq True | Select-Object -expand FriendlyName

    $Changes = Compare-Object $FirstScan $SecondScan

    foreach ($Device in $Changes) {
      if ($Device.SideIndicator -eq "=>") {
        "`"$($Device.InputObject)`" connected."
      }
      if ($Device.SideIndicator -eq "<=") {
        "`"$($Device.InputObject)`" removed."
      }
    }

    $FirstScan = $SecondScan
  }
}

