function Get-vtsUSB {
  <#
  .DESCRIPTION
  Returns USB devices and their current status.
  .EXAMPLE
  PS> Get-vtsUSB
  
  Output:
  FriendlyName                                          Present Status
  ------------                                          ------- ------
  Microsoft LifeCam VX-3000                                True OK
  EPSON Utility                                            True OK
  American Power Conversion USB UPS                        True OK
  Microsoft LifeCam VX-3000.                               True OK
  FULL HD 1080P Webcam                                     True OK
  SmartSource Pro/Value                                    True OK
  EPSON ES-400                                             True OK
  .EXAMPLE
  PS> Get-vtsUSB epson
  
  Output:
  FriendlyName                                          Present Status
  ------------                                          ------- ------
  EPSON Utility                                            True OK
  EPSON ES-400                                             True OK
  
  .LINK
  Device Management
  #>
  param (
    $searchTerm
  )
    
  get-pnpdevice -friendlyName *$searchTerm* |
  Where-Object { $_.InstanceId -like "*usb*" } |
  Select-Object FriendlyName, Present, Status -unique |
  Sort-Object Present -Descending
}

