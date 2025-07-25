function Install-vtsChoco {
  <#
  .DESCRIPTION
  Installs Chocolatey.
  .EXAMPLE
  PS> Install-vtsChoco
  
  Output:
  Forcing web requests to allow TLS v1.2 (Required for requests to Chocolatey.org)
  Getting latest version of the Chocolatey package for download.
  Not using proxy.
  Getting Chocolatey from https://community.chocolatey.org/api/v2/package/chocolatey/1.2.0.
  Downloading https://community.chocolatey.org/api/v2/package/chocolatey/1.2.0 to C:\Users\rober\AppData\Local\Temp\chocolatey\chocoInstall\chocolatey.zip
  Not using proxy.
  Extracting C:\Users\rober\AppData\Local\Temp\chocolatey\chocoInstall\chocolatey.zip to C:\Users\rober\AppData\Local\Temp\chocolatey\chocoInstall
  Installing Chocolatey on the local machine
  Creating ChocolateyInstall as an environment variable (targeting 'Machine')
    Setting ChocolateyInstall to 'C:\ProgramData\chocolatey'
  WARNING: It's very likely you will need to close and reopen your shell
    before you can use choco.
  Restricting write permissions to Administrators
  We are setting up the Chocolatey package repository.
  The packages themselves go to 'C:\ProgramData\chocolatey\lib'
    (i.e. C:\ProgramData\chocolatey\lib\yourPackageName).
  A shim file for the command line goes to 'C:\ProgramData\chocolatey\bin'
    and points to an executable in 'C:\ProgramData\chocolatey\lib\yourPackageName'.
  
  Creating Chocolatey folders if they do not already exist.
  
  WARNING: You can safely ignore errors related to missing log files when
    upgrading from a version of Chocolatey less than 0.9.9.
    'Batch file could not be found' is also safe to ignore.
    'The system cannot find the file specified' - also safe.
  chocolatey.nupkg file not installed in lib.
   Attempting to locate it from bootstrapper.
  PATH environment variable does not have C:\ProgramData\chocolatey\bin in it. Adding...
  Adding Chocolatey to the profile. This will provide tab completion, refreshenv, etc.
  WARNING: Chocolatey profile installed. Reload your profile - type . $profile
  Chocolatey (choco.exe) is now ready.
  You can call choco from anywhere, command line or powershell by typing choco.
  Run choco /? for a list of functions.
  You may need to shut down and restart powershell and/or consoles
   first prior to using choco.
  Ensuring Chocolatey commands are on the path
  Ensuring chocolatey.nupkg is in the lib folder
  
  .LINK
  Package Management
  #>
  Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
}

