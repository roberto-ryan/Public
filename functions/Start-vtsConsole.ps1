function Start-vtsConsole {
  <#
  .SYNOPSIS
  Starts a console session for serial communication using PuTTY's plink.exe.
  
  .DESCRIPTION
  The Start-vtsConsole function checks for the presence of plink.exe and Chocolatey, installs them if necessary, retrieves the COM port name of devices with class 'Ports' and status 'OK', and starts a console session using plink.exe for serial communication.
  
  .PARAMETER PlinkPath
  The path to the plink.exe executable. Default is "C:\Program Files\PuTTY\plink.exe".
  
  .PARAMETER ChocolateyBinPath
  The path to the Chocolatey bin directory. Default is "C:\ProgramData\chocolatey\bin\choco.exe".
  
  .EXAMPLE
  PS C:\> Start-vtsConsole
  
  This command starts a console session using the default paths for plink.exe and Chocolatey.
  
  .EXAMPLE
  PS C:\> Start-vtsConsole -PlinkPath "D:\Tools\PuTTY\plink.exe" -ChocolateyBinPath "D:\choco\bin\choco.exe"
  
  This command starts a console session using the specified paths for plink.exe and Chocolatey.
  
  .INPUTS
  None
  
  .OUTPUTS
  None
  
  .LINK
  Network
  #>
  param (
      [string]$PlinkPath = "C:\Program Files\PuTTY\plink.exe",
      [string]$ChocolateyBinPath = "C:\ProgramData\chocolatey\bin\choco.exe"
  )

  # Check if plink.exe exists in the specified path
  if (!(Test-Path $PlinkPath)){
      Write-Host "plink.exe not found. Checking for Chocolatey installation..."
      # Check if Chocolatey is installed and install it if not
      if (!(Test-Path $ChocolateyBinPath)){
          Write-Host "Chocolatey not found. Installing Chocolatey..."
          # Remove any existing Chocolatey directory forcefully
          Remove-Item "C:\ProgramData\chocolatey" -Recurse -Force
          # Call the function to install Chocolatey
          Install-vtsChoco
          Write-Host "Chocolatey installed. Installing PuTTY..."
          # Install PuTTY using Chocolatey
          choco install putty -y
          Write-Host "PuTTY installed."
      } else {
          Write-Host "Chocolatey is already installed."
      }
  } else {
      Write-Host "plink.exe found."
  }

  # Retrieve the COM port name of devices with class 'Ports' and status 'OK'
  Write-Host "Retrieving COM port name..."
  $COMPORT = (Get-PnpDevice |
  Where-Object { $_.Class -eq 'Ports' -and $_.Status -eq 'OK' } |
  Select-Object -ExpandProperty FriendlyName) -replace "USB Serial Port \(" -replace "\)"
  Write-Host "COM port name retrieved: $COMPORT"

  # Execute plink.exe with the serial communication port
  Write-Host "Executing plink.exe with the serial communication port: $COMPORT"
  & $PlinkPath -serial $COMPORT
}

