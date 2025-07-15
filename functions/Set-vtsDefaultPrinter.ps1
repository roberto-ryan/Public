function Set-vtsDefaultPrinter {
  <#
  .Description
  Sets the default printer.
  
  .EXAMPLE
  PS> Set-vtsDefaultPrinter -Name <"Printer">
  
  .EXAMPLE
  PS> Set-vtsDefaultPrinter -Name "HP Laserjet"
  
  .LINK
  Print Management
  #>
  Param(
    [Parameter(
      Mandatory = $true)]
    $Name
  )
  $printerName = Get-printer "*$Name*" | Select-Object -ExpandProperty Name -First 1
  if ($null -ne $printerName) {
    $confirm = Read-Host -Prompt "Set $printerName as the default printer? (y/n)"
    if ($confirm -eq "y") {
      Write-Host "Setting $printerName as the default printer. Estimated time: 42 seconds"
      $wsh = New-Object -ComObject WScript.Network
      $wsh.SetDefaultPrinter($printerName)
    }
    else {
      Write-Host "exiting..."
    }
  }
  else {
    Write-Host "There are no matching printers."
  }
}

