function Add-vtsPrinter {
  <#
  .SYNOPSIS
      A function to add a printer from a print server.
  
  .DESCRIPTION
      This function takes a server name and a printer name as input, searches for the printer on the server, and adds it to the local machine. 
      It prompts the user for confirmation before adding the printer. The function uses the Get-Printer cmdlet to get the list of printers from the server, 
      and the Add-Printer cmdlet to add the printer.
  
  .PARAMETER Server
      The name of the server where the printer is located.
  
  .PARAMETER Name
      The name of the printer to be added. Wildcards can be used for searching.
  
  .EXAMPLE
      PS C:\> Add-vtsPrinter -Server "ch-dc" -Name "*P18*"
  
      This command will search for a printer with a name that includes "P18" on the server "ch-dc", and add it to the local machine after user confirmation.
  
  .LINK
      Print Management
      #>
  param(
    [Parameter(Mandatory = $true)]
    [string]$Server,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  # Get the list of printers from the print server
  $printers = @(Get-Printer -ComputerName $Server -Name "*$Name*")

  if ($printers.Count -eq 0) {
    Write-Output "No printer found with the name $Name on server $Server"
    return
  }
  else {
    # Create a hashtable of printers
    $printerTable = @{}
    for ($i = 0; $i -lt $printers.Count; $i++) {
      $printerTable.Add($i + 1, $printers[$i])
      Write-Output "$($i+1): $($printers[$i].Name)"
    }

    # Ask the user which printers to install
    $userInput = Read-Host "Enter the numbers of the printers you want to install, separated by commas, or enter * to install all printers"

    if ($userInput -eq '*') {
      $keys = $printerTable.Keys
    }
    else {
      $keys = $userInput.Split(',') | ForEach-Object { [int]$_ }
    }

    foreach ($key in $keys) {
      $printer = $printerTable[$key]
      # Add the printer
      try {
        Add-Printer -ConnectionName "\\$Server\$($printer.Name)"
        if ($null -ne (Get-Printer -Name "\\$Server\$($printer.Name)")) {
          Write-Host "`nPrinter $($printer.Name) added successfully.`n" -f Green
          $newPrinter = (Get-Printer -Name "\\$Server\$($printer.Name)")
          Write-Host "Name  : $($newPrinter.Name)`nDriver: $($newPrinter.DriverName)`nPort  : $($newPrinter.PortName)`n"
          if (($($Printer.DriverName)) -ne ($($newPrinter.DriverName))) {
            Write-Host "Driver mismatch. Printer server is using: `n$($Printer.DriverName)" -f Yellow
          }
        }
        else {
          Write-Error "Failed to add printer $($printer.Name)."
        }
      }
      catch {
        Write-Error "Failed to add printer $($printer.Name). $($_.Exception.Message)"
      }
    }
  }
}

