function Get-vtsPrinterByPortAddress {
  <#
  .SYNOPSIS
      A function to get printer(s) by port address.
  
  .DESCRIPTION
      This function retrieves printer(s) associated with a given port address. If no port address is provided, it retrieves all printers.
  
  .PARAMETER PrinterHostAddress
      The port address of the printer. This parameter is optional. If not provided, the function retrieves all printers.
  
  .EXAMPLE
      PS C:\> Get-vtsPrinterByPortAddress -PrinterHostAddress "192.168.1.10"
      This command retrieves the printer(s) associated with the port address "192.168.1.10".
  
  .EXAMPLE
      PS C:\> Get-vtsPrinterByPortAddress
      This command retrieves all printers as no port address is provided.
  
  .NOTES
      The function returns a custom object that includes the printer host address and the associated printers.
  
  .LINK
      Print Management
  #>
  param (
    # PrinterHostAddress is not mandatory
    [Parameter(Mandatory = $false)]
    [string]$PrinterHostAddress
  )

  # If PrinterHostAddress is provided
  if ($PrinterHostAddress) {
    # Get the port names of the printer
    $PortNames = Get-PrinterPort | Where-Object { $_.PrinterHostAddress -eq "$PrinterHostAddress" } | Select-Object -ExpandProperty Name
    # Initialize an array to store printers
    $Printers = @()
    # Loop through each port name
    foreach ($PortName in $PortNames) {
      # Get the printer with the port name and add it to the printers array
      $Printers += Get-Printer | Where-Object { $_.PortName -eq "$PortName" } | Sort-Object Name
    }
    # If printers are found
    if ($Printers.Name) {
      # Create a custom object to store the printer host address and the printers
      $Result = [PSCustomObject]@{
        PrinterHostAddress = $PrinterHostAddress
        Printers           = $Printers.Name -join ", "
      }
      # Return the result
      $Result
    }
  }
  # If PrinterHostAddress is not provided
  else {
    # Get all printer ports
    Get-PrinterPort | ForEach-Object {
      # Get the port name
      $PortName = $_.Name
      # Get the printer with the port name
      $Printers = Get-Printer | Where-Object { $_.PortName -eq "$PortName" } | Sort-Object Name
      # If printers are found
      if ($Printers.Name) {
        # Create a custom object to store the printer host address and the printers
        $Result = [PSCustomObject]@{
          PrinterHostAddress = $_.PrinterHostAddress
          Printers           = $Printers.Name -join ", "
        }
        # Return the result
        $Result
      }
    }
  }
}

