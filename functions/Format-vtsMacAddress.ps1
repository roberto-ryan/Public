function Format-vtsMacAddress {
  <#
  .SYNOPSIS
  This function formats a MAC address input by the user or from the clipboard if not specified.
  
  .DESCRIPTION
  The function Format-vtsMacAddress takes a MAC address as a parameter, removes any separators, converts it to lowercase, checks if it is 12 characters long, and then formats it by inserting colons after every 2 characters. If the MacAddress parameter is not specified, the function will use the current content of the clipboard. The formatted MAC address is then copied to the clipboard.
  
  .PARAMETER MacAddress
  A string representing the MAC address to be formatted. If not specified, the function will use the current content of the clipboard.
  
  .EXAMPLE
  Format-vtsMacAddress -MacAddress "00:0a:95:9d:68:16"
  
  This example formats the provided MAC address.
  
  .EXAMPLE
  Format-vtsMacAddress
  
  This example formats the MAC address currently in the clipboard.
  
  .LINK
  Utilities
  #>
  param (
    $MacAddress = (Get-Clipboard)
  )

  # Remove any separators from the MAC address and convert it to lowercase
  $cleanMac = (($MacAddress -replace '[-:.]', '').ToLower()).Trim() | Where-Object Length -eq 12

  # Check if the MAC address is 12 characters long
  if ($cleanMac.Length -ne 12) {
    Write-Host "The MAC address is invalid. It should be 12 hexadecimal characters."
  }
  else {
    # Insert colons after every 2 characters to format the MAC address
    $outputMac = $cleanMac -replace '(.{2})', '$1:'
    # Remove the trailing colon
    $outputMac = $outputMac.TrimEnd(':')

    # Output the formatted MAC address
    Write-Host "Copied to clipboard:"
    Write-Host "show mac address-table dynamic address $outputMac"
    "show mac address-table dynamic address $outputMac" | Set-Clipboard
  }
}

