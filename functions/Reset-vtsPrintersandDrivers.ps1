function Reset-vtsPrintersandDrivers {
  <#
  .SYNOPSIS
      This function resets the printer drivers and settings on a Windows machine.
  
  .DESCRIPTION
      The Reset-vtsPrintersandDrivers function is a destructive process that resets the printer drivers and settings on a Windows machine. It should be used as a last resort. The function first prompts the user for confirmation before proceeding. It then checks for the RunAsUser module and installs it if not present. The function then gets the network printers and saves them to a temporary directory. It attempts to remove the driver and registry paths with the spooler service running, then stops the spooler service and tries again. Finally, it starts the spooler service and removes the printer drivers.
  
  .PARAMETER driverPath
      The path to the printer drivers. Default is "C:\Windows\System32\spool\drivers".
  
  .PARAMETER printProcessorRegPath
      The registry path to the print processors. Default is "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Print Processors".
  
  .PARAMETER driverRegPath
      The registry path to the drivers. Default is "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers".
  
  .PARAMETER printProcessorName
      The name of the print processor. Default is "winprint".
  
  .PARAMETER printProcessorDll
      The DLL of the print processor. Default is "winprint.dll".
  
  .EXAMPLE
      Reset-vtsPrintersandDrivers
      This command will reset the printer drivers and settings on the machine with the default parameters.
  
  .EXAMPLE
      Reset-vtsPrintersandDrivers -driverPath "C:\CustomPath\drivers"
      This command will reset the printer drivers and settings on the machine with a custom driver path.
  
  .LINK
      Print Management
  #>
  param (
    [string]$driverPath = "C:\Windows\System32\spool\drivers",
    [string]$printProcessorRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Print Processors",
    [string]$driverRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers",
    [string]$printProcessorName = "winprint",
    [string]$printProcessorDll = "winprint.dll"
  )

  Write-Host "Starting Reset-vtsPrintersandDrivers function..."

  $userConfirmation = Read-Host -Prompt "This is a destructive process and should be used as a last resort. Are you sure you want to proceed? (yes/no)"
  if ($userConfirmation -ne 'yes') {
    Write-Host "Operation cancelled by user."
    return
  }

  Write-Host "Checking for RunAsUser module..."
  if (!(Get-Module -ListAvailable -Name RunAsUser)) {
    Write-Host "Installing RunAsUser module..."
    Install-Module RunAsUser -Force
  }

  Invoke-AsCurrentUser {
    Write-Host "Getting network printers..."
    $printers = Get-Printer "\\*" | Select-Object -ExpandProperty Name
    $tempPath = "C:\temp"
    if (!(Test-Path $tempPath)) {
      Write-Host "Creating temp directory..."
      New-Item -ItemType Directory -Path $tempPath
    }
    $printers | Out-File "$tempPath\printers.txt" -Append
    Write-Host "Network printers saved to $tempPath\printers.txt" -ForegroundColor Yellow
  }

  Write-Host "Attempting to remove driver and registry paths with spooler running..."
  $items = Get-ChildItem -Path $driverPath -Recurse -Depth 0
  $items | Sort-Object -Property FullName -Descending | ForEach-Object { Remove-Item -Path $_.FullName -Recurse -Force -Confirm:$false }
  Remove-Item -Path $printProcessorRegPath -Recurse -Force -Confirm:$false
  Remove-Item -Path $driverRegPath -Recurse -Force -Confirm:$false
  
  Write-Host "Stopping spooler service..."
  net stop spooler
  
  Write-Host "Attempting to remove driver and registry paths with spooler stopped..."
  $items = Get-ChildItem -Path $driverPath -Recurse -Depth 0
  $items | Sort-Object -Property FullName -Descending | ForEach-Object { Remove-Item -Path $_.FullName -Recurse -Force -Confirm:$false }
  Remove-Item -Path $printProcessorRegPath -Recurse -Force -Confirm:$false
  Remove-Item -Path $driverRegPath -Recurse -Force -Confirm:$false

  Write-Host "Starting spooler service..."
  net start spooler

  Write-Host "Removing printer drivers with spooler started..."
  Get-PrinterDriver | Remove-PrinterDriver -Confirm:$false

  if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run PowerShell as an Administrator."
    return
  }

  Try {
    Write-Host "Checking for existing Print Processor..."
    if (Test-Path "$printProcessorRegPath\$printProcessorName") {
      Write-Host "Print Processor '$printProcessorName' already exists. Consider updating existing registry entries instead."
    }

    Write-Host "Creating new registry entries for Print Processor..."
    New-Item -Path "$printProcessorRegPath\$printProcessorName" -Force | Out-Null
    New-ItemProperty -Path "$printProcessorRegPath\$printProcessorName" -Name "Driver" -Value $printProcessorDll -PropertyType String -Force | Out-Null
    $path = "HKLM:\SYSTEM\CURRENTCONTROLSET\CONTROL\PRINT\ENVIRONMENTS\WINDOWS X64\PRINT PROCESSORS\winprint"
    if (!(Test-Path $path)) {
      New-Item -Path $path -Force | Out-Null
    }
    Set-ItemProperty -Path $path -Name "Driver" -Value "winprint.dll" | Out-Null

    Write-Host "Registry entries for Print Processor '$printProcessorName' have been (re)created successfully."
  }
  Catch {
    Write-Error "An error occurred while recreating registry entries for Print Processor '$printProcessorName': $_"
  }

  # Define the registry keys and values
  $registryKeys = @(
    @{
      Path  = "HKLM:\Software\Policies\Microsoft\Windows\DriverInstall\Restrictions"
      Name  = "AllowUserDeviceClasses"
      Value = 1
    },
    @{
      Path  = "HKLM:\Software\Policies\Microsoft\Windows NT\Printers\PointAndPrint"
      Name  = "RestrictDriverInstallationToAdministrators"
      Value = 0
    }
  )

  # Loop through each registry key
  foreach ($registryKey in $registryKeys) {
    # Check if the registry key exists
    if (!(Test-Path $registryKey.Path)) {
      # If the registry key doesn't exist, create it
      New-Item -Path $registryKey.Path -Force | Out-Null
      Write-Host "Created registry key $($registryKey.Path)"
    }

    try {
      # Get the current value of the registry key
      $property = Get-ItemProperty -Path $registryKey.Path -ErrorAction SilentlyContinue
      $currentValue = $property.($registryKey.Name)

      if ($currentValue -eq $registryKey.Value) {
        # If the current value is the same as the desired value, print a message and continue to the next iteration
        Write-Host "Registry key $($registryKey.Path)\$($registryKey.Name) is already set to $($registryKey.Value). No change was made."
      }

      # Check if the property exists
      if ($null -eq $currentValue) {
        # If the property doesn't exist, create it
        New-ItemProperty -Path $registryKey.Path -Name $registryKey.Name -Value $registryKey.Value -PropertyType DWORD -Force | Out-Null
        Write-Host "Created property $($registryKey.Name) with value $($registryKey.Value) in $($registryKey.Path)"
      }
      else {
        # If the property exists, set its value
        Set-ItemProperty -Path $registryKey.Path -Name $registryKey.Name -Value $registryKey.Value -ErrorAction Stop
        Write-Host "Successfully set $($registryKey.Name) to $($registryKey.Value) in $($registryKey.Path)"
      }
    }
    catch {
      # Catch any errors
      Write-Host "Failed to set $($registryKey.Name) to $($registryKey.Value) in $($registryKey.Path): $_"
    }
  }

  Invoke-AsCurrentUser {
    Write-Host "Restoring network printers..."
    $printers = Get-Content "C:\temp\printers.txt" | Select-Object -Unique
    foreach ($p in $printers) {
      Write-Host "Adding printer $p..."
      Add-Printer -ConnectionName "$p"
    }
  }

  Write-Host "Reset-vtsPrintersandDrivers function completed."
}

