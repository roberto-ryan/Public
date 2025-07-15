function Get-vtsPrinterAddressMapping {
  <#
  .SYNOPSIS
  Retrieves a mapping of printer names to their corresponding port names and host addresses.
  
  .DESCRIPTION
  The Get-vtsPrinterAddressMapping function retrieves all printer ports with their host addresses and maps them to the printers using those ports. It returns a list of custom objects containing the printer name, port name, and printer host address.
  
  .EXAMPLE
  PS C:\> Get-vtsPrinterAddressMapping
  
  This command retrieves and displays a list of printers along with their port names and host addresses.
  
  .OUTPUTS
  System.Object
  A custom object with the following properties:
  - PrinterName: The name of the printer.
  - PortName: The name of the port.
  - PrinterHostAddress: The host address of the printer port.
  
  .LINK
  Print Management
  #>

  $PortNames = Get-PrinterPort | Where-Object { $_.PrinterHostAddress } | Select-Object Name, PrinterHostAddress
  # Initialize an array to store results
  $Results = @()
  
  # Loop through each port
  foreach ($Port in $PortNames) {
      # Get the printer with the port name
      $Printers = Get-Printer | Where-Object { $_.PortName -eq $Port.Name } | Sort-Object Name
      
      # If printers are found
      if ($Printers.Name) {
          # Create a custom object for each printer
          foreach ($Printer in $Printers) {
              $Result = [PSCustomObject]@{
                  PrinterName = $Printer.Name
                  PortName = $Port.Name
                  PrinterHostAddress = $Port.PrinterHostAddress
              }
              $Results += $Result
          }
      }
  }
  
  $Results
}

