function Stop-vtsPacketCapture {
  <#
  .SYNOPSIS
  Stops the currently running packet capture.
  
  .DESCRIPTION
  The Stop-vtsPacketCapture function stops the currently running packet capture started by Start-vtsPacketCapture.
  
  .EXAMPLE
  Stop-vtsPacketCapture
  Stops the currently running packet capture.
  
  .LINK
  Network
  #>
  Write-Host "Stopping packet capture..."
  Get-Process tshark | Stop-Process -Confirm:$false
  Write-Host "Packet capture stopped."
}

