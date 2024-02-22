<#
.Description
Searches the last 500 System and Application logs using a term.
.EXAMPLE
PS> Search-vtsEventLog <term>
.EXAMPLE
PS> Search-vtsEventLog "driver"

Output:
TimeGenerated : 9/13/2022 9:14:30 AM
Message       : Media disconnected on NIC /DEVICE/{90E7B0EA-AE78-4836-8CBC-B73F1BCD5894} (Friendly Name: Microsoft
                Network Adapter Multiplexor Driver).
Log           : System

.LINK
Log Management
#>
function Search-vtsEventLog {
  [CmdletBinding()]
  Param(
    [Parameter(Position = 0, Mandatory,
      ParameterSetName = 'SearchTerm')]
    [string]$SearchTerm
  )
  [array]$Logname = @(
    "System"
    "Application"
  )
        
  $result = @()

  foreach ($log in $Logname) {
    Get-EventLog -LogName $log -EntryType Error, Warning -Newest 500 2>$null |
    Where-Object Message -like "*$SearchTerm*" |
    Select-Object TimeGenerated, Message |
    ForEach-Object {
      $result += [PSCustomObject]@{
        TimeGenerated = $_.TimeGenerated
        Message       = $_.Message
        Log           = $log
      }
    }
  }
    
  foreach ($log in $Logname) {
    if ($null -eq ($result | Where-Object Log -like "$log")) {
      Write-Host "$($log) Log - No Matches Found" -ForegroundColor Yellow
    }
  }

  $result | Sort-Object TimeGenerated | Format-List
}

<#
.Description
Displays Mapped Drives information from the Windows Registry.
.EXAMPLE
PS> Get-vtsMappedDrive

Output:
Username            : VTS-ROBERTO\rober
DriveLetter         : Y
RemotePath          : https://live.sysinternals.com
ConnectWithUsername : rober
SID                 : S-1-5-21-376445358-2603134888-3166729622-1001

.LINK
Drive Management
#>
function Get-vtsMappedDrive {
  # This is required for Verbose to work correctly.
  # If you don't want the Verbose message, remove "-Verbose" from the Parameters field.
  [CmdletBinding()]
  param ()

  # On most OSes, HKEY_USERS only contains users that are logged on.
  # There are ways to load the other profiles, but it can be problematic.
  $Drives = Get-ItemProperty "Registry::HKEY_USERS\*\Network\*" 2>$null

  # See if any drives were found
  if ( $Drives ) {

    ForEach ( $Drive in $Drives ) {

      # PSParentPath looks like this: Microsoft.PowerShell.Core\Registry::HKEY_USERS\S-1-5-21-##########-##########-##########-####\Network
      $SID = ($Drive.PSParentPath -split '\\')[2]

      [PSCustomObject]@{
        # Use .NET to look up the username from the SID
        Username            = ([System.Security.Principal.SecurityIdentifier]"$SID").Translate([System.Security.Principal.NTAccount])
        DriveLetter         = $Drive.PSChildName
        RemotePath          = $Drive.RemotePath

        # The username specified when you use "Connect using different credentials".
        # For some reason, this is frequently "0" when you don't use this option. I remove the "0" to keep the results consistent.
        ConnectWithUsername = $Drive.UserName -replace '^0$', $null
        SID                 = $SID
      }

    }

  }
  else {

    Write-Verbose "No mapped drives were found"

  }
}

<#
.Description
Continuous Ping Report. Tracks failed ping times and outputs data to a logfile.
.EXAMPLE
PS> Start-vtsPingReport google.com
.EXAMPLE
PS> Start-vtsPingReport 8.8.8.8

Output:
Start Time : 09/14/2022 10:31:20

Ping Target: 8.8.8.8

Total Ping Count     : 10
Successful Ping Count: 10
Failed Ping Count    : 0

Last Successful Ping : 09/14/2022 10:31:30

Press Ctrl-C to exit
logfile saved to C:\temp\PingResults-8.8.8.8.log

.LINK
Network
#>
function Start-vtsPingReport {
  Param(
    [Parameter(
      Mandatory = $true,
      ValueFromPipeline = $true)]
    $PingTarget
  )
    
  try {
    $output = "C:\temp\PingResults-$PingTarget.log"
    if (-not (Test-Path $output)) {
      New-Item -Path $output -ItemType File -Force | Out-Null
    }
    $startTime = (Get-Date)
    $lastSuccess = $null
    $failedTimes = @()
    
    $successCount = 0
    $failCount = 0
    $totalPingCount = 0
    
    while ($true) {
      $totalPingCount++
      $pingResult = Test-Connection $PingTarget -Count 1 2>$null
      if (($pingResult.StatusCode -eq 0) -or ($pingResult.Status -eq "Success")) {
        $successCount++
        $lastSuccess = (Get-Date)
      }
      else {
        $failCount++
        $failedTimes += "$(Get-Date) - Ping#$totalPingCount"
      }
      Clear-Host
      Write-Host "Start Time : $startTime"
      Write-Host "`nPing Target: $PingTarget"
      Write-Host "`nTotal Ping Count     : $totalPingCount"
      Write-Host "Successful Ping Count: $successCount" -ForegroundColor Green
      Write-Host "Failed Ping Count    : $failCount" -ForegroundColor DarkRed
      Write-Host "`nLast Successful Ping : $lastSuccess" -ForegroundColor Green
            
      if ($failCount -gt 0) {
        Write-Host "`n-----Last 30 Failed Pings-----" -ForegroundColor DarkRed
        $failedTimes | Select-Object -last 30 | Sort-Object -Descending
        Write-Host "------------------------------" -ForegroundColor DarkRed
      }
      Write-Host "`nPress Ctrl-C to exit" -ForegroundColor Yellow
    
      Start-Sleep 1    
    }
  }
  finally {
    Write-Host "logfile saved to $output"
    Write-Output "Start Time : $startTime" | Out-File $output
    Write-Output "End Time   : $(Get-Date)" | Out-File $output -Append
    Write-Output "`nPing Target: $PingTarget" | Out-File $output -Append
    Write-Output "`nTotal Ping Count     : $totalPingCount" | Out-File $output -Append
    Write-Output "Successful Ping Count: $successCount" | Out-File $output -Append
    Write-Output "Failed Ping Count    : $failCount" | Out-File $output -Append
    Write-Output "`nLast Successful Ping : $lastSuccess" | Out-File $output -Append
    if ($failCount -gt 0) {
      Write-Output "`n-------Pings Failed at:-------" | Out-File $output -Append
      $failedTimes | Out-File $output -Append
      Write-Output "------------------------------" | Out-File $output -Append
    }
  }
}

<#
.Description
Generates a random 12 character password and copies it to the clipboard.
.EXAMPLE
PS> New-vtsRandomPassword

Output:
Random Password Copied to Clipboard

.LINK
Utilities
#>
function New-vtsRandomPassword {
  $numbers = 0..9
  $symbols = '!', '@', '#', '$', '%', '*', '?', '+', '='
  $string = ( -join ((0x30..0x39) + ( 0x41..0x5A) + ( 0x61..0x7A) |
      Get-Random -Count 12  |
      ForEach-Object { [char]$_ }))
  $number = $numbers | Get-Random
  $symbol = $symbols | Get-Random
  $NewPW = $string + $number + $symbol

  $NewPW | Set-Clipboard

  Write-Output "Random Password Copied to Clipboard"
}

<#
.Description
Converts strings to the phonetic alphabet.
.EXAMPLE
PS> "RandomString" | Out-vtsPhoneticAlphabet

Output:
ROMEO
alfa
november
delta
oscar
mike
SIERRA
tango
romeo
india
november
golf

.LINK
Utilities
#>
function Out-vtsPhoneticAlphabet {
  [CmdletBinding()]
  [OutputType([String])]
  Param
  (
    # Input string to convert
    [Parameter(Mandatory = $true, 
      ValueFromPipeline = $true,
      Position = 0)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $InputObject
  )
  $result = @()
  $nato = @{
    '0' = '(ZERO)'
    '1' = '(ONE)'
    '2' = '(TWO)'
    '3' = '(THREE)'
    '4' = '(FOUR)'
    '5' = '(FIVE)'
    '6' = '(SIX)'
    '7' = '(SEVEN)'
    '8' = '(EIGHT)'
    '9' = '(NINE)'
    'a' = 'alfa'
    'b' = 'bravo'
    'c' = 'charlie'
    'd' = 'delta'
    'e' = 'echo'
    'f' = 'foxtrot'
    'g' = 'golf'
    'h' = 'hotel'
    'i' = 'india'
    'j' = 'juliett'
    'k' = 'kilo'
    'l' = 'lima'
    'm' = 'mike'
    'n' = 'november'
    'o' = 'oscar'
    'p' = 'papa'
    'q' = 'quebec'
    'r' = 'romeo'
    's' = 'sierra'
    't' = 'tango'
    'u' = 'uniform'
    'v' = 'victor'
    'w' = 'whiskey'
    'x' = 'xray'
    'y' = 'yankee'
    'z' = 'zulu'
    '.' = '(PERIOD)'
    '-' = '(DASH)'
  }

  $chars = ($InputObject).ToCharArray()

  foreach ($char in $chars) {
    switch -Regex -CaseSensitive ($char) {
      '\d' {
        $result += ($nato["$char"])
        break
      }
      '[a-z]' {
        $result += ($nato["$char"]).ToLower()
        break
      }
      '[A-Z]' {
        $result += ($nato["$char"]).ToUpper()
        break
      }

      Default { $result += $char }
    }
  }
  $result
}

<#
.Description
Displays monitor connection type (HDMI, DisplayPort, etc.)
.EXAMPLE
PS> Get-vtsDisplayConnectionType

Output:
GSM M2362D (DisplayPort (external))
GSM M2362D (HDMI)

.LINK
System Information
#>
function Get-vtsDisplayDetails {
  $adapterTypes = @{
    '-2'         = 'Unknown'
    '-1'         = 'Unknown'
    '0'          = 'VGA'
    '1'          = 'S-Video'
    '2'          = 'Composite'
    '3'          = 'Component'
    '4'          = 'DVI'
    '5'          = 'HDMI'
    '6'          = 'LVDS'
    '8'          = 'D-Jpn'
    '9'          = 'SDI'
    '10'         = 'DisplayPort (external)'
    '11'         = 'DisplayPort (internal)'
    '12'         = 'Unified Display Interface'
    '13'         = 'Unified Display Interface (embedded)'
    '14'         = 'SDTV dongle'
    '15'         = 'Miracast'
    '16'         = 'Internal'
    '2147483648' = 'Internal'
  }
  $arrMonitors = @()
  $monitors = Get-WmiObject WmiMonitorID -Namespace root/wmi
  $connections = Get-WmiObject WmiMonitorConnectionParams -Namespace root/wmi
  foreach ($monitor in $monitors) {
    $manufacturer = $monitor.ManufacturerName
    $name = $monitor.UserFriendlyName
    $serialNumber = $monitor.SerialNumberID
    $connectionType = ($connections | Where-Object { $_.InstanceName -eq $monitor.InstanceName }).VideoOutputTechnology
    if ($manufacturer -ne $null) { $manufacturer = [System.Text.Encoding]::ASCII.GetString($manufacturer -ne 0) }
    if ($name -ne $null) { $name = [System.Text.Encoding]::ASCII.GetString($name -ne 0) }
    if ($serialNumber -ne $null) { $serialNumber = [System.Text.Encoding]::ASCII.GetString($serialNumber).Trim([char]0) }
    $connectionType = $adapterTypes."$connectionType"
    if ($connectionType -eq $null) { $connectionType = 'Unknown' }
    if (($manufacturer -ne $null) -or ($name -ne $null) -or ($serialNumber -ne $null)) { 
      $arrMonitors += "$manufacturer $name, Serial: $serialNumber ($connectionType)" 
    }
  }
  $i = 0
  $strMonitors = ''
  if ($arrMonitors.Count -gt 0) {
    foreach ($monitor in $arrMonitors) {
      if ($i -eq 0) { $strMonitors += $arrMonitors[$i] }
      else { $strMonitors += "`n"; $strMonitors += $arrMonitors[$i] }
      $i++
    }
  }
  if ($strMonitors -eq '') { $strMonitors = 'None Found' }
  $strMonitors
}


<#
.DESCRIPTION
Forces the installation of a specified Google Chrome extension.
.EXAMPLE
PS> Install-vtsChromeExtension -extensionId "ddloeodolhdfbohkokiflfbacbfpjahp"

.LINK
Package Management
#>
function Install-vtsChromeExtension {
  param(
    [string]$extensionId,
    [switch]$info
  )
  if ($info) {
    $InformationPreference = "Continue"
  }
  if (!($extensionId)) {
    # Empty Extension
    $result = "No Extension ID"
  }
  else {
    Write-Information "ExtensionID = $extensionID"
    $extensionId = "$extensionId;https://clients2.google.com/service/update2/crx"
    $regKey = "HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist"
    if (!(Test-Path $regKey)) {
      New-Item $regKey -Force
      Write-Information "Created Reg Key $regKey"
    }
    # Add Extension to Chrome
    $extensionsList = New-Object System.Collections.ArrayList
    $number = 0
    $noMore = 0
    do {
      $number++
      Write-Information "Pass : $number"
      try {
        $install = Get-ItemProperty $regKey -name $number -ErrorAction Stop
        $extensionObj = [PSCustomObject]@{
          Name  = $number
          Value = $install.$number
        }
        $extensionsList.add($extensionObj) | Out-Null
        Write-Information "Extension List Item : $($extensionObj.name) / $($extensionObj.value)"
      }
      catch {
        $noMore = 1
      }
    }
    until($noMore -eq 1)
    $extensionCheck = $extensionsList | Where-Object { $_.Value -eq $extensionId }
    if ($extensionCheck) {
      $result = "Extension Already Exists"
      Write-Information "Extension Already Exists"
    }
    else {
      $newExtensionId = $extensionsList[-1].name + 1
      New-ItemProperty HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist -PropertyType String -Name $newExtensionId -Value $extensionId
      $result = "Installed"
    }
  }
  $result
}

<#
.DESCRIPTION
Returns temperature of thermal sensor on motherboard. Not accurate for CPU temp.
.EXAMPLE
PS> Get-vtsTemperature

Output:
27.85 C : 82.1300000000001 F : 301K

.LINK
System Information
#>
function Get-vtsTemperature {
  $t = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi"
  $returntemp = @()

  foreach ($temp in $t.CurrentTemperature) {


    $currentTempKelvin = $temp / 10
    $currentTempCelsius = $currentTempKelvin - 273.15

    $currentTempFahrenheit = (9 / 5) * $currentTempCelsius + 32

    $returntemp += $currentTempCelsius.ToString() + " C : " + $currentTempFahrenheit.ToString() + " F : " + $currentTempKelvin + "K"  
  }
  return $returntemp
}

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
function Get-vtsUSB {
  param (
    $searchTerm
  )
    
  get-pnpdevice -friendlyName *$searchTerm* |
  Where-Object { $_.InstanceId -like "*usb*" } |
  Select-Object FriendlyName, Present, Status -unique |
  Sort-Object Present -Descending
}

<#
.DESCRIPTION
Returns physical disk stats. Defaults to C drive if no driver letter is specified.
.EXAMPLE
PS> Get-vtsDiskStat

Output:
Drive Property                                   Value
----- --------                                   -----
C     avg. disk bytes/read                       65536
C     disk write bytes/sec            44645.6174840863
C     avg. disk bytes/write           4681.14285714286
C     % idle time                     96.9313065314281
C     split io/sec                    6.88540550759652
C     disk transfers/sec              0.99463553271784
C     % disk write time              0.716335878527066
C     avg. disk read queue length  0.00802039105365663
C     avg. disk write queue length 0.00792347156238492
C     avg. disk sec/write          0.00246081666666667
C     avg. disk sec/transfer                         0
C     avg. disk sec/read                             0
C     disk reads/sec                                 0
C     disk writes/sec                                0
C     disk bytes/sec                                 0
C     disk read bytes/sec                            0
C     % disk read time                               0
C     avg. disk bytes/transfer                       0
C     avg. disk queue length                         0
C     % disk time                                    0
C     current disk queue length                      0
.EXAMPLE
PS> Get-vtsDiskStat -DriveLetter D

Output:
Drive Property                                   Value
----- --------                                   -----
D     avg. disk bytes/read                       65536
D     disk write bytes/sec            44645.6174840863
D     avg. disk bytes/write           4681.14285714286
D     % idle time                     96.9313065314281
D     split io/sec                    6.88540550759652
D     disk transfers/sec              0.99463553271784
D     % disk write time              0.716335878527066
D     avg. disk read queue length  0.00802039105365663
D     avg. disk write queue length 0.00792347156238492
D     avg. disk sec/write          0.00246081666666667
D     avg. disk sec/transfer                         0
D     avg. disk sec/read                             0
D     disk reads/sec                                 0
D     disk writes/sec                                0
D     disk bytes/sec                                 0
D     disk read bytes/sec                            0
D     % disk read time                               0
D     avg. disk bytes/transfer                       0
D     avg. disk queue length                         0
D     % disk time                                    0
D     current disk queue length                      0

.LINK
System Information
#>
function Get-vtsDiskStat {
  param (
    $DriveLetter = "C"
  )
        
  $a = (Get-Counter -List PhysicalDisk).PathsWithInstances |
  Select-String "$DriveLetter`:" |
  Foreach-object {
    Get-Counter -Counter "$_"
  }
    
  if ($null -eq $a) {
    Write-Output "$DriveLetter drive not found."
  }

  $stats = @()
        
  foreach ($i in $a) {
    $stats += [PSCustomObject]@{
      Drive    = ($DriveLetter).ToUpper()
      Property = (($i.CounterSamples.Path).split(")") | Select-String ^\\[a-z%]) -replace '\\', ''
      Value    = $i.CounterSamples.CookedValue
    }
  }
        
  $stats | Sort-Object Value -Descending
}

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
function Install-vtsChoco {
  Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
}

<#
.DESCRIPTION
Runs speedtest by Ookla. Installs via chocolatey.
.EXAMPLE
PS> Start-vtsSpeedTest

Output:
   Speedtest by Ookla

     Server: Sparklight - Anniston, AL (id = 8829)
        ISP: Spectrum Business
    Latency:    25.11 ms   (0.18 ms jitter)
   Download:   236.80 Mbps (data used: 369.4 MB )
     Upload:   309.15 Mbps (data used: 526.0 MB )
Packet Loss:     0.0%
 Result URL: https://www.speedtest.net/result/c/23d057dd-8de5-4d62-aef9-72beb122d7a4

 .LINK
Network
 #>
function Start-vtsSpeedTest {
  if (Test-Path "C:\ProgramData\chocolatey\bin\speedtest.exe") {
    C:\ProgramData\chocolatey\bin\speedtest.exe
  }
  elseif (Test-Path C:\ProgramData\chocolatey\lib\speedtest\tools\speedtest.exe) {
    C:\ProgramData\chocolatey\lib\speedtest\tools\speedtest.exe
  }
  elseif (Test-Path "C:\ProgramData\chocolatey\choco.exe") {
    choco install speedtest -y
    speedtest
  }
  else {
    Install-vtsChoco
    choco install speedtest -y
    speedtest
  }
}
 
<#
.DESCRIPTION
Installs all pending Windows Updates.
.EXAMPLE
PS> Install-vtsWindowsUpdate

Output:
X ComputerName Result     KB          Size Title
- ------------ ------     --          ---- -----
1 CH-BIMA-W... Accepted   KB5018202   68MB 2022-10 Cumulative Update Preview for .NET Framework 3.5, 4.8 and 4.8...
2 CH-BIMA-W... Downloaded KB5018202   68MB 2022-10 Cumulative Update Preview for .NET Framework 3.5, 4.8 and 4.8...
3 CH-BIMA-W... Installed  KB5018202   68MB 2022-10 Cumulative Update Preview for .NET Framework 3.5, 4.8 and 4.8...
Reboot is required. Do it now? [Y / N] (default is 'N')

.LINK
Package Management
#>
function Install-vtsWindowsUpdate {
  $NuGet = Get-PackageProvider -Name NuGet
  if ($null -eq $NuGet) {
    Install-PackageProvider -Name NuGet -Force
  }

  $PSWindowsUpdate = Get-Module -Name PSWindowsUpdate
  if ($null -eq $PSWindowsUpdate) {
    Install-Module PSWindowsUpdate -Force -Confirm:$false
  }

  Import-Module PSWindowsUpdate -Force
  Get-WindowsUpdate
  Install-WindowsUpdate -AcceptAll
}

<#
.DESCRIPTION
Searches the Remote Desktop Gateway connection log.
.EXAMPLE
PS> Search-vtsRDPGatewayLog

Output:
TimeCreated : 11/17/2022 9:38:46 AM
Message     : The user "domain\user1", on client computer "100.121.185.200", connected to resource
              "PC-21.domain.local". Connection protocol used: "HTTP".
Log         : Microsoft-Windows-TerminalServices-Gateway/Operational

TimeCreated : 11/17/2022 9:45:40 AM
Message     : The user "domain\user2", on client computer "172.56.65.179", disconnected from the following
              network resource: "PC-p2.domain.local". Before the user disconnected, the client transferred 1762936
              bytes and received 6198054 bytes. The client session duration was 4947 seconds. Connection protocol
              used: "HTTP".
Log         : Microsoft-Windows-TerminalServices-Gateway/Operational

TimeCreated : 11/17/2022 9:46:01 AM
Message     : The user "domain\user1", on client computer "100.121.185.200", disconnected from the following network
              resource: "PC-21.domain.local". Before the user disconnected, the client transferred 1348808 bytes
              and received 4463546 bytes. The client session duration was 435 seconds. Connection protocol used:
              "HTTP".
Log         : Microsoft-Windows-TerminalServices-Gateway/Operational
.EXAMPLE
PS> Search-vtsRDPGatewayLog robert

Output:
TimeCreated : 11/17/2022 9:45:40 AM
Message     : The user "domain\robert.ryan", on client computer "172.56.65.179", disconnected from the following
              network resource: "PC-p2.domain.local". Before the user disconnected, the client transferred 1762936
              bytes and received 6198054 bytes. The client session duration was 4947 seconds. Connection protocol
              used: "HTTP".
Log         : Microsoft-Windows-TerminalServices-Gateway/Operational

TimeCreated : 11/17/2022 9:46:01 AM
Message     : The user "domain\robert.ryan, on client computer "100.121.185.200", disconnected from the following network
              resource: "PC-21.domain.local". Before the user disconnected, the client transferred 1348808 bytes
              and received 4463546 bytes. The client session duration was 435 seconds. Connection protocol used:
              "HTTP".
Log         : Microsoft-Windows-TerminalServices-Gateway/Operational

.LINK
Log Management
#>
function Search-vtsRDPGatewayLog {
  [CmdletBinding()]
  Param(
    [string]$SearchTerm
  )
  [array]$Logname = @(
    "Microsoft-Windows-TerminalServices-Gateway/Operational"
  )
        
  $result = @()

  foreach ($log in $Logname) {
    Get-WinEvent -LogName $log 2>$null |
    Where-Object Message -like "*$SearchTerm*" |
    Select-Object TimeCreated, Message |
    ForEach-Object {
      $result += [PSCustomObject]@{
        TimeCreated = $_.TimeCreated
        Message     = $_.Message
        Log         = $log
      }
    }
  }
    
  foreach ($log in $Logname) {
    if ($null -eq ($result | Where-Object Log -like "$log")) {
      Write-Host "$($log) Log - No Matches Found" -ForegroundColor Yellow
    }
  }

  $result | Sort-Object TimeCreated | Format-List
}

<#
.DESCRIPTION
Uninstall ninja and remove keys, files and services.
.EXAMPLE
PS> Uninstall-vtsNinja

.LINK
Package Management
#>
function Uninstall-vtsNinja {
  param (
    [Parameter(Mandatory = $false)]
    [switch]$DelTeamViewer = $false,
    [Parameter(Mandatory = $false)]
    [switch]$Cleanup = $true,
    [Parameter(Mandatory = $false)]
    [switch]$Uninstall = $true
  )
    
  #Set-PSDebug -Trace 2
    
  if ([system.environment]::Is64BitOperatingSystem) {
    $ninjaPreSoftKey = 'HKLM:\SOFTWARE\WOW6432Node\NinjaRMM LLC'
    $uninstallKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    $exetomsiKey = 'HKLM:\SOFTWARE\WOW6432Node\EXEMSI.COM\MSI Wrapper\Installed'
  }
  else {
    $ninjaPreSoftKey = 'HKLM:\SOFTWARE\NinjaRMM LLC'
    $uninstallKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    $exetomsiKey = 'HKLM:\SOFTWARE\EXEMSI.COM\MSI Wrapper\Installed'
  }
    
  $ninjaSoftKey = Join-Path $ninjaPreSoftKey -ChildPath 'NinjaRMMAgent'
    
  $ninjaDir = [string]::Empty
  $ninjaDataDir = Join-Path -Path $env:ProgramData -ChildPath "NinjaRMMAgent"
    
  ###################################################################################################
  # locating NinjaRMMAgent
  ###################################################################################################
  $ninjaDirRegLocation = $(Get-ItemPropertyValue $ninjaSoftKey -Name Location)
  if ($ninjaDirRegLocation) {
    if (Join-Path -Path $ninjaDirRegLocation -ChildPath "NinjaRMMAgent.exe" | Test-Path) {
      #location confirmed from registry location
      $ninjaDir = $ninjaDirRegLocation
    }
  }
    
  if (!$ninjaDir) {
    #attempt to get the path from service
    $ss = Get-WmiObject win32_service -Filter 'Name Like "NinjaRMMAgent"'
    if ($ss) {
      $ninjaDirService = ($(Get-WmiObject win32_service -Filter 'Name Like "NinjaRMMAgent"').PathName | Split-Path).Replace("`"", "")
      if (Join-Path -Path $ninjaDirService -ChildPath "NinjaRMMAgentPatcher.exe" | Test-Path) {
        #location confirmed from service location
        $ninjaDir = $ninjaDirService
      }
    }
  }
    
  if ($ninjaDir) {
    $ninjaDir.Replace('/', '\')
    
    if ($Uninstall) {
      #there are few measures agent takes to prevent accidental uninstllation
      #disable those measures now
      #it automatically takes care if those measures are already removed
      #it is not possible to check those measures outside of the agent since agent's development comes parralel to this script
      & "$ninjaDir\NinjaRMMAgent.exe" -disableUninstallPrevention NOUI
    
      # Executes uninstall.exe in Ninja install directory
      $Arguments = @(
        "/uninstall"
        $(Get-WmiObject -Class win32_product -Filter "Name='NinjaRMMAgent'").IdentifyingNumber
        "/quiet"
        "/log"
        "NinjaRMMAgent_uninstall.log"
        "/L*v"
        "WRAPPED_ARGUMENTS=`"--mode unattended`""
      )
    
      Start-Process -FilePath "msiexec.exe"  -Verb RunAs -Wait -NoNewWindow -WhatIf -ArgumentList $Arguments
    }
  }
    
    
  if ($Cleanup) {
    $service = Get-Service "NinjaRMMAgent" 
    if ($service) {
      Stop-Service $service -Force
      & sc.exe DELETE NinjaRMMAgent
      #Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NinjaRMMAgent
    }
    
    # Delete Ninja install directory and all contents
    if (Test-Path $ninjaDir) {
      & cmd.exe /c rd /s /q $ninjaDir
    }
    
    if (Test-Path $ninjaDataDir) {
      & cmd.exe /c rd /s /q $ninjaDataDir
    }
    
    #Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\NinjaRMM LLC\NinjaRMMAgent
    Remove-Item -Path  -Recurse -Force
    
    # Will search registry locations for NinjaRMMAgent value and delete parent key
    # Search $uninstallKey
    $keys = Get-ChildItem $uninstallKey | Get-ItemProperty -name 'DisplayName'
    foreach ($key in $keys) {
      if ($key.'DisplayName' -eq 'NinjaRMMAgent') {
        Remove-Item $key.PSPath -Recurse -Force
      }
    }
    
    #Search $installerKey
    $keys = Get-ChildItem 'HKLM:\SOFTWARE\Classes\Installer\Products' | Get-ItemProperty -name 'ProductName'
    foreach ($key in $keys) {
      if ($key.'ProductName' -eq 'NinjaRMMAgent') {
        Remove-Item $key.PSPath -Recurse -Force
      }
    }
    # Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\A0313090625DD2B4F824C1EAE0958B08\InstallProperties
    $keys = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products'
    foreach ($key in $keys) {
      #ridiculous, MS sucks
      $kn = $key.Name -replace 'HKEY_LOCAL_MACHINE' , 'HKLM:'; 
      $k1 = Join-Path $kn -ChildPath 'InstallProperties';
      if ( $(Get-ItemProperty -Path $k1 -Name DisplayName).DisplayName -eq 'NinjaRMMAgent') {
        $toremove = 
        Get-Item -LiteralPath $kn | Remove-Item -Recurse -Force
      }
    }
    
    #Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\EXEMSI.COM\MSI Wrapper\Installed\NinjaRMMAgent 5.3.3681
    Get-ChildItem $exetomsiKey | Where-Object -Property Name -CLike '*NinjaRMMAgent*'  | Remove-Item -Recurse -Force
    
    #HKLM:\SOFTWARE\WOW6432Node\NinjaRMM LLC
    Get-Item -Path $ninjaPreSoftKey | Remove-Item -Recurse -Force
    
    # agent creates this key by mistake but we delete it here
    Get-Item -Path "HKLM:\SOFTWARE\WOW6432Node\WOW6432Node\NinjaRMM LLC" | Remove-Item -Recurse -Force
    
  }
    
  if (Get-Item -Path $ninjaPreSoftKey) {
    Write-Output "Failed to remove NinjaRMMAgent reg keys ", $ninjaPreSoftKey
  }
    
  if (Get-Service "NinjaRMMAgent") {
    Write-Output "Failed to remove NinjaRMMAgent service"
  }
    
  if ($ninjaDir) {
    if (Test-Path $ninjaDir) {
      Write-Output "Failed to remove NinjaRMMAgent program folder"
      if (Join-Path -Path $ninjaDir -ChildPath "NinjaRMMAgent.exe" | Test-Path) {
        Write-Output "Failed to remove NinjaRMMAgent.exe"
      }
    
      if (Join-Path -Path $ninjaDir -ChildPath "NinjaRMMAgentPatcher.exe" | Test-Path) {
        Write-Output "Failed to remove NinjaRMMAgentPatcher.exe"
      }
    }
  }
    
    
    
  # Uninstall TeamViewer only if -DelTeamViewer parameter specified
  if ($DelTeamViewer -eq $true) {
    $tvProcess = Get-Process -Name 'teamviewer'
    Stop-Process -InputObject $tvProcess -Force # Stops TeamViewer process
    & '%ProgramFiles%\TeamViewer\uninstall.exe' /S | out-null
    & '%ProgramFiles(x86)%\TeamViewer\uninstall.exe' /S | out-null
    Remove-Item -path HKLM:\SOFTWARE\TeamViewer -Recurse
    Remove-Item -path HKLM:\SOFTWARE\WOW6432Node\TeamViewer -Recurse
    Remove-Item -path HKLM:\SOFTWARE\WOW6432Node\TVInstallTemp -Recurse
    Remove-Item -path HKLM:\SOFTWARE\TeamViewer -Recurse
    Remove-Item -path HKLM:\SOFTWARE\Wow6432Node\TeamViewer -Recurse
  }    
}

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
function Set-vtsDefaultPrinter {
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

<#
.DESCRIPTION
Returns users toast notifications. Duplicates notifications are removed for brevity.
.EXAMPLE
Show notifications for all users
PS> Show-vtsToastNotification
.EXAMPLE
Show notifications for a selected user
PS> Show-vtsToastNotification -user john.doe

.LINK
Device Management
#>
function Show-vtsToastNotification {
  param(
    $user = (Get-ChildItem C:\Users\ | Select-Object -ExpandProperty Name)
  )

  $db = foreach ($u in $user) {
    Get-Content "C:\Users\$u\AppData\Local\Microsoft\Windows\Notifications\wpndatabase.db-wal" 2>$null
  }

  $tags = @(
    'text>'
    'text id="1">'
    'text id="2">'
  )

  $notification = foreach ($tag in $tags) {
    ($db -split '<' |
    Select-String $tag |
    Select-Object -ExpandProperty Line) -replace "$tag", "" -replace "</text>", "" |
    Select-String -NotMatch '/'
  }

  Write-Host "Duplicates removed for brevity." -ForegroundColor Yellow
  $notification | Select-Object -Unique
}

<#
.DESCRIPTION
Maps a remote drive.
.EXAMPLE
PS> New-vtsMappedDrive -Letter A -Path \\192.168.0.4\sharedfolder
.EXAMPLE
PS> New-vtsMappedDrive -Letter A -Path "\\192.168.0.4\folder with spaces"

.LINK
Drive Management
#>
function New-vtsMappedDrive {
  param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Letter,
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )
  New-PSDrive -Name "$Letter" -PSProvider FileSystem -Root "$Path" -Persist -Scope Global
}

<#
.DESCRIPTION
Install PowerShell 7
.EXAMPLE
PS> Install-vtsPwsh

.LINK
Package Management
#>
function Install-vtsPwsh {
  msiexec.exe /i "https://github.com/PowerShell/PowerShell/releases/download/v7.3.3/PowerShell-7.3.3-win-x64.msi" /qn
  Write-Host "Installing PowerShell 7... Please wait" -ForegroundColor Cyan
  While (-not (Test-Path "C:\Program Files\PowerShell\7\pwsh.exe" 2>$null)) { Start-Sleep 5 }
  & "C:\Program Files\PowerShell\7\pwsh.exe" 2>$null
}

<#
.DESCRIPTION
    Get the default application associated with a specific file extension. If no extension is provided, it returns a list of all file extensions and their associated default applications.

.NOTES
    Version    : 1.2.0
    Author(s)  : Danyfirex & Dany3j
    Credits    : https://bbs.pediy.com/thread-213954.htm
                 LMongrain - Hash Algorithm PureBasic Version
    License    : MIT License
    Copyright  : 2022 Danysys. <danysys.com>

.EXAMPLE
    Get-FTA .pdf
    Show Default Application Program Id for an Extension

.LINK
File Association Management
    #>
function Get-FTA {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [String]
    $Extension
  )

  
  if ($Extension) {
    Write-Verbose "Get File Type Association for $Extension"
    
    $assocFile = (Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension\UserChoice" -ErrorAction SilentlyContinue).ProgId
    Write-Output $assocFile
  }
  else {
    Write-Verbose "Get File Type Association List"

    $assocList = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\* |
    ForEach-Object {
      $progId = (Get-ItemProperty "$($_.PSParentPath)\$($_.PSChildName)\UserChoice" -ErrorAction SilentlyContinue).ProgId
      if ($progId) {
        "$($_.PSChildName), $progId"
      }
    }
    Write-Output $assocList
  }
  
}

<#
.DESCRIPTION
    Get the default application associated with a specific protocol. If no protocol is provided, it returns a list of all protocols and their associated default applications.

.NOTES
    Version    : 1.2.0
    Author(s)  : Danyfirex & Dany3j
    Credits    : https://bbs.pediy.com/thread-213954.htm
                 LMongrain - Hash Algorithm PureBasic Version
    License    : MIT License
    Copyright  : 2022 Danysys. <danysys.com>

.EXAMPLE
    Get-PTA

.LINK
File Association Management
    #>
function Get-PTA {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [String]
    $Protocol
  )

  if ($Protocol) {
    Write-Verbose "Get Protocol Type Association for $Protocol"

    $assocFile = (Get-ItemProperty "HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\$Protocol\UserChoice" -ErrorAction SilentlyContinue).ProgId
    Write-Output $assocFile
  }
  else {
    Write-Verbose "Get Protocol Type Association List"

    $assocList = Get-ChildItem HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\* |
    ForEach-Object {
      $progId = (Get-ItemProperty "$($_.PSParentPath)\$($_.PSChildName)\UserChoice" -ErrorAction SilentlyContinue).ProgId
      if ($progId) {
        "$($_.PSChildName), $progId"
      }
    }
    Write-Output $assocList
  }
}

<#
.DESCRIPTION
    Register an application and set it as the default for a specified file extension or protocol.

.NOTES
    Version    : 1.2.0
    Author(s)  : Danyfirex & Dany3j
    Credits    : https://bbs.pediy.com/thread-213954.htm
                 LMongrain - Hash Algorithm PureBasic Version
    License    : MIT License
    Copyright  : 2022 Danysys. <danysys.com>

.EXAMPLE
    Register-FTA "C:\SumatraPDF.exe" .pdf -Icon "shell32.dll,100"
    This example registers SumatraPDF as the default PDF reader.

.LINK
File Association Management
    #>
function Register-FTA {
  [CmdletBinding()]
  param (
    [Parameter( Position = 0, Mandatory = $true)]
    [ValidateScript( { Test-Path $_ })]
    [String]
    $ProgramPath,

    [Parameter( Position = 1, Mandatory = $true)]
    [Alias("Protocol")]
    [String]
    $Extension,
    
    [Parameter( Position = 2, Mandatory = $false)]
    [String]
    $ProgId,
    
    [Parameter( Position = 3, Mandatory = $false)]
    [String]
    $Icon
  )

  Write-Verbose "Register Application + Set Association"
  Write-Verbose "Application Path: $ProgramPath"
  if ($Extension.Contains(".")) {
    Write-Verbose "Extension: $Extension"
  }
  else {
    Write-Verbose "Protocol: $Extension"
  }
  
  if (!$ProgId) {
    $ProgId = "SFTA." + [System.IO.Path]::GetFileNameWithoutExtension($ProgramPath).replace(" ", "") + $Extension
  }
  
  $progCommand = """$ProgramPath"" ""%1"""
  Write-Verbose "ApplicationId: $ProgId" 
  Write-Verbose "ApplicationCommand: $progCommand"
  
  try {
    $keyPath = "HKEY_CURRENT_USER\SOFTWARE\Classes\$Extension\OpenWithProgids"
    [Microsoft.Win32.Registry]::SetValue( $keyPath, $ProgId, ([byte[]]@()), [Microsoft.Win32.RegistryValueKind]::None)
    $keyPath = "HKEY_CURRENT_USER\SOFTWARE\Classes\$ProgId\shell\open\command"
    [Microsoft.Win32.Registry]::SetValue($keyPath, "", $progCommand)
    Write-Verbose "Register ProgId and ProgId Command OK"
  }
  catch {
    throw "Register ProgId and ProgId Command FAILED"
  }
  
  Set-FTA -ProgId $ProgId -Extension $Extension -Icon $Icon
}

<#
.DESCRIPTION
    Remove the file type association for a given application.

.NOTES
    Version    : 1.2.0
    Author(s)  : Danyfirex & Dany3j
    Credits    : https://bbs.pediy.com/thread-213954.htm
                 LMongrain - Hash Algorithm PureBasic Version
    License    : MIT License
    Copyright  : 2022 Danysys. <danysys.com>

.LINK
File Association Management
    #>
function Remove-FTA {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [Alias("ProgId")]
    [String]
    $ProgramPath,

    [Parameter(Mandatory = $true)]
    [String]
    $Extension
  )
  
  function local:Remove-UserChoiceKey {
    param (
      [Parameter( Position = 0, Mandatory = $True )]
      [String]
      $Key
    )

    $code = @'
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Win32;
    
    namespace Registry {
      public class Utils {
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern int RegOpenKeyEx(UIntPtr hKey, string subKey, int ulOptions, int samDesired, out UIntPtr hkResult);
    
        [DllImport("advapi32.dll", SetLastError=true, CharSet = CharSet.Unicode)]
        private static extern uint RegDeleteKey(UIntPtr hKey, string subKey);

        public static void DeleteKey(string key) {
          UIntPtr hKey = UIntPtr.Zero;
          RegOpenKeyEx((UIntPtr)0x80000001u, key, 0, 0x20019, out hKey);
          RegDeleteKey((UIntPtr)0x80000001u, key);
        }
      }
    }
'@

    try {
      Add-Type -TypeDefinition $code
    }
    catch {}

    try {
      [Registry.Utils]::DeleteKey($Key)
    }
    catch {} 
  } 

  function local:Update-Registry {
    $code = @'
    [System.Runtime.InteropServices.DllImport("Shell32.dll")] 
    private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
    public static void Refresh() {
        SHChangeNotify(0x8000000, 0, IntPtr.Zero, IntPtr.Zero);    
    }
'@ 

    try {
      Add-Type -MemberDefinition $code -Namespace SHChange -Name Notify
    }
    catch {}

    try {
      [SHChange.Notify]::Refresh()
    }
    catch {} 
  }

  if (Test-Path -Path $ProgramPath) {
    $ProgId = "SFTA." + [System.IO.Path]::GetFileNameWithoutExtension($ProgramPath).replace(" ", "") + $Extension
  }
  else {
    $ProgId = $ProgramPath
  }

  try {
    $keyPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension\UserChoice"
    Write-Verbose "Remove User UserChoice Key If Exist: $keyPath"
    Remove-UserChoiceKey $keyPath

    $keyPath = "HKCU:\SOFTWARE\Classes\$ProgId"
    Write-Verbose "Remove Key If Exist: $keyPath"
    Remove-Item -Path $keyPath -Recurse -ErrorAction Stop | Out-Null
    
  }
  catch {
    Write-Verbose "Key No Exist: $keyPath"
  }

  try {
    $keyPath = "HKCU:\SOFTWARE\Classes\$Extension\OpenWithProgids"
    Write-Verbose "Remove Property If Exist: $keyPath Property $ProgId"
    Remove-ItemProperty -Path $keyPath -Name $ProgId -ErrorAction Stop | Out-Null
    
  }
  catch {
    Write-Verbose "Property No Exist: $keyPath Property: $ProgId"
  } 

  Update-Registry
  Write-Output "Removed: $ProgId" 
}

<#
.DESCRIPTION
    Set the default file type association. It can be used to set the default application for a specific file type or protocol.
.NOTES
    Version    : 1.2.0
    Author(s)  : Danyfirex & Dany3j
    Credits    : https://bbs.pediy.com/thread-213954.htm
                 LMongrain - Hash Algorithm PureBasic Version
    License    : MIT License
    Copyright  : 2022 Danysys. <danysys.com>
    
.EXAMPLE
    Set-FTA AcroExch.Document.DC .pdf
    Set Acrobat Reader DC as Default .pdf reader
 
.EXAMPLE
    Set-FTA Applications\SumatraPDF.exe .pdf
    Set Sumatra PDF as Default .pdf reader

.LINK
File Association Management
    #>
function Set-FTA {

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [String]
    $ProgId,

    [Parameter(Mandatory = $true)]
    [Alias("Protocol")]
    [String]
    $Extension,
      
    [String]
    $Icon,

    [switch]
    $DomainSID
  )
  
  if (Test-Path -Path $ProgId) {
    $ProgId = "SFTA." + [System.IO.Path]::GetFileNameWithoutExtension($ProgId).replace(" ", "") + $Extension
  }

  Write-Verbose "ProgId: $ProgId"
  Write-Verbose "Extension/Protocol: $Extension"

  
  #Write required Application Ids to ApplicationAssociationToasts
  #When more than one application associated with an Extension/Protocol is installed ApplicationAssociationToasts need to be updated
  function local:Write-RequiredApplicationAssociationToasts {
    param (
      [Parameter( Position = 0, Mandatory = $True )]
      [String]
      $ProgId,

      [Parameter( Position = 1, Mandatory = $True )]
      [String]
      $Extension
    )
    
    try {
      $keyPath = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\ApplicationAssociationToasts"
      [Microsoft.Win32.Registry]::SetValue($keyPath, $ProgId + "_" + $Extension, 0x0) 
      Write-Verbose ("Write Reg ApplicationAssociationToasts OK: " + $ProgId + "_" + $Extension)
    }
    catch {
      Write-Verbose ("Write Reg ApplicationAssociationToasts FAILED: " + $ProgId + "_" + $Extension)
    }
    
    $allApplicationAssociationToasts = Get-ChildItem -Path HKLM:\SOFTWARE\Classes\$Extension\OpenWithList\* -ErrorAction SilentlyContinue | 
    ForEach-Object {
      "Applications\$($_.PSChildName)"
    }

    $allApplicationAssociationToasts += @(
      ForEach ($item in (Get-ItemProperty -Path HKLM:\SOFTWARE\Classes\$Extension\OpenWithProgids -ErrorAction SilentlyContinue).PSObject.Properties ) {
        if ([string]::IsNullOrEmpty($item.Value) -and $item -ne "(default)") {
          $item.Name
        }
      })

    
    $allApplicationAssociationToasts += Get-ChildItem -Path HKLM:SOFTWARE\Clients\StartMenuInternet\* , HKCU:SOFTWARE\Clients\StartMenuInternet\* -ErrorAction SilentlyContinue | 
    ForEach-Object {
    (Get-ItemProperty ("$($_.PSPath)\Capabilities\" + (@("URLAssociations", "FileAssociations") | Select-Object -Index $Extension.Contains("."))) -ErrorAction SilentlyContinue).$Extension
    }
    
    $allApplicationAssociationToasts | 
    ForEach-Object { if ($_) {
        if (Set-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\ApplicationAssociationToasts $_"_"$Extension -Value 0 -Type DWord -ErrorAction SilentlyContinue -PassThru) {
          Write-Verbose  ("Write Reg ApplicationAssociationToastsList OK: " + $_ + "_" + $Extension)
        }
        else {
          Write-Verbose  ("Write Reg ApplicationAssociationToastsList FAILED: " + $_ + "_" + $Extension)
        }
      } 
    }

  }

  function local:Update-RegistryChanges {
    $code = @'
    [System.Runtime.InteropServices.DllImport("Shell32.dll")] 
    private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
    public static void Refresh() {
        SHChangeNotify(0x8000000, 0, IntPtr.Zero, IntPtr.Zero);    
    }
'@ 

    try {
      Add-Type -MemberDefinition $code -Namespace SHChange -Name Notify
    }
    catch {}

    try {
      [SHChange.Notify]::Refresh()
    }
    catch {} 
  }
  

  function local:Set-Icon {
    param (
      [Parameter( Position = 0, Mandatory = $True )]
      [String]
      $ProgId,

      [Parameter( Position = 1, Mandatory = $True )]
      [String]
      $Icon
    )

    try {
      $keyPath = "HKEY_CURRENT_USER\SOFTWARE\Classes\$ProgId\DefaultIcon"
      [Microsoft.Win32.Registry]::SetValue($keyPath, "", $Icon) 
      Write-Verbose "Write Reg Icon OK"
      Write-Verbose "Reg Icon: $keyPath"
    }
    catch {
      Write-Verbose "Write Reg Icon FAILED"
    }
  }


  function local:Write-ExtensionKeys {
    param (
      [Parameter( Position = 0, Mandatory = $True )]
      [String]
      $ProgId,

      [Parameter( Position = 1, Mandatory = $True )]
      [String]
      $Extension,

      [Parameter( Position = 2, Mandatory = $True )]
      [String]
      $ProgHash
    )
    

    function local:Remove-UserChoiceKey {
      param (
        [Parameter( Position = 0, Mandatory = $True )]
        [String]
        $Key
      )

      $code = @'
      using System;
      using System.Runtime.InteropServices;
      using Microsoft.Win32;
      
      namespace Registry {
        public class Utils {
          [DllImport("advapi32.dll", SetLastError = true)]
          private static extern int RegOpenKeyEx(UIntPtr hKey, string subKey, int ulOptions, int samDesired, out UIntPtr hkResult);
      
          [DllImport("advapi32.dll", SetLastError=true, CharSet = CharSet.Unicode)]
          private static extern uint RegDeleteKey(UIntPtr hKey, string subKey);
  
          public static void DeleteKey(string key) {
            UIntPtr hKey = UIntPtr.Zero;
            RegOpenKeyEx((UIntPtr)0x80000001u, key, 0, 0x20019, out hKey);
            RegDeleteKey((UIntPtr)0x80000001u, key);
          }
        }
      }
'@
  
      try {
        Add-Type -TypeDefinition $code
      }
      catch {}

      try {
        [Registry.Utils]::DeleteKey($Key)
      }
      catch {} 
    } 

    
    try {
      $keyPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension\UserChoice"
      Write-Verbose "Remove Extension UserChoice Key If Exist: $keyPath"
      Remove-UserChoiceKey $keyPath
    }
    catch {
      Write-Verbose "Extension UserChoice Key No Exist: $keyPath"
    }
  

    try {
      $keyPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension\UserChoice"
      [Microsoft.Win32.Registry]::SetValue($keyPath, "Hash", $ProgHash)
      [Microsoft.Win32.Registry]::SetValue($keyPath, "ProgId", $ProgId)
      Write-Verbose "Write Reg Extension UserChoice OK"
    }
    catch {
      throw "Write Reg Extension UserChoice FAILED"
    }
  }


  function local:Write-ProtocolKeys {
    param (
      [Parameter( Position = 0, Mandatory = $True )]
      [String]
      $ProgId,

      [Parameter( Position = 1, Mandatory = $True )]
      [String]
      $Protocol,

      [Parameter( Position = 2, Mandatory = $True )]
      [String]
      $ProgHash
    )
      

    try {
      $keyPath = "HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\$Protocol\UserChoice"
      Write-Verbose "Remove Protocol UserChoice Key If Exist: $keyPath"
      Remove-Item -Path $keyPath -Recurse -ErrorAction Stop | Out-Null
    
    }
    catch {
      Write-Verbose "Protocol UserChoice Key No Exist: $keyPath"
    }
  

    try {
      $keyPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\$Protocol\UserChoice"
      [Microsoft.Win32.Registry]::SetValue( $keyPath, "Hash", $ProgHash)
      [Microsoft.Win32.Registry]::SetValue($keyPath, "ProgId", $ProgId)
      Write-Verbose "Write Reg Protocol UserChoice OK"
    }
    catch {
      throw "Write Reg Protocol UserChoice FAILED"
    }
    
  }

  
  function local:Get-UserExperience {
    [OutputType([string])]
    $hardcodedExperience = "User Choice set via Windows User Experience {D18B6DD5-6124-4341-9318-804003BAFA0B}"
    $userExperienceSearch = "User Choice set via Windows User Experience"
    $userExperienceString = ""
    $user32Path = [Environment]::GetFolderPath([Environment+SpecialFolder]::SystemX86) + "\Shell32.dll"
    $fileStream = [System.IO.File]::Open($user32Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $binaryReader = New-Object System.IO.BinaryReader($fileStream)
    [Byte[]] $bytesData = $binaryReader.ReadBytes(5mb)
    $fileStream.Close()
    $dataString = [Text.Encoding]::Unicode.GetString($bytesData)
    $position1 = $dataString.IndexOf($userExperienceSearch)
    $position2 = $dataString.IndexOf("}", $position1)
    try {
      $userExperienceString = $dataString.Substring($position1, $position2 - $position1 + 1)
    }
    catch {
      $userExperienceString = $hardcodedExperience
    }
    Write-Output $userExperienceString
  }
  

  function local:Get-UserSid {
    [OutputType([string])]
    $userSid = ((New-Object System.Security.Principal.NTAccount([Environment]::UserName)).Translate([System.Security.Principal.SecurityIdentifier]).value).ToLower()
    Write-Output $userSid
  }

  #use in this special case
  #https://github.com/DanysysTeam/PS-SFTA/pull/7
  function local:Get-UserSidDomain {
    if (-not ("System.DirectoryServices.AccountManagement" -as [type])) {
      Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    }
    [OutputType([string])]
    $userSid = ([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).SID.Value.ToLower()
    Write-Output $userSid
  }



  function local:Get-HexDateTime {
    [OutputType([string])]

    $now = [DateTime]::Now
    $dateTime = [DateTime]::New($now.Year, $now.Month, $now.Day, $now.Hour, $now.Minute, 0)
    $fileTime = $dateTime.ToFileTime()
    $hi = ($fileTime -shr 32)
    $low = ($fileTime -band 0xFFFFFFFFL)
    $dateTimeHex = ($hi.ToString("X8") + $low.ToString("X8")).ToLower()
    Write-Output $dateTimeHex
  }
  
  function Get-Hash {
    [CmdletBinding()]
    param (
      [Parameter( Position = 0, Mandatory = $True )]
      [string]
      $BaseInfo
    )


    function local:Get-ShiftRight {
      [CmdletBinding()]
      param (
        [Parameter( Position = 0, Mandatory = $true)]
        [long] $iValue, 
            
        [Parameter( Position = 1, Mandatory = $true)]
        [int] $iCount 
      )
    
      if ($iValue -band 0x80000000) {
        Write-Output (( $iValue -shr $iCount) -bxor 0xFFFF0000)
      }
      else {
        Write-Output  ($iValue -shr $iCount)
      }
    }
    

    function local:Get-Long {
      [CmdletBinding()]
      param (
        [Parameter( Position = 0, Mandatory = $true)]
        [byte[]] $Bytes,
    
        [Parameter( Position = 1)]
        [int] $Index = 0
      )
    
      Write-Output ([BitConverter]::ToInt32($Bytes, $Index))
    }
    

    function local:Convert-Int32 {
      param (
        [Parameter( Position = 0, Mandatory = $true)]
        [long] $Value
      )
    
      [byte[]] $bytes = [BitConverter]::GetBytes($Value)
      return [BitConverter]::ToInt32( $bytes, 0) 
    }

    [Byte[]] $bytesBaseInfo = [System.Text.Encoding]::Unicode.GetBytes($baseInfo) 
    $bytesBaseInfo += 0x00, 0x00  
    
    $MD5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    [Byte[]] $bytesMD5 = $MD5.ComputeHash($bytesBaseInfo)
    
    $lengthBase = ($baseInfo.Length * 2) + 2 
    $length = (($lengthBase -band 4) -le 1) + (Get-ShiftRight $lengthBase  2) - 1
    $base64Hash = ""

    if ($length -gt 1) {
    
      $map = @{PDATA = 0; CACHE = 0; COUNTER = 0 ; INDEX = 0; MD51 = 0; MD52 = 0; OUTHASH1 = 0; OUTHASH2 = 0;
        R0 = 0; R1 = @(0, 0); R2 = @(0, 0); R3 = 0; R4 = @(0, 0); R5 = @(0, 0); R6 = @(0, 0); R7 = @(0, 0)
      }
    
      $map.CACHE = 0
      $map.OUTHASH1 = 0
      $map.PDATA = 0
      $map.MD51 = (((Get-Long $bytesMD5) -bor 1) + 0x69FB0000L)
      $map.MD52 = ((Get-Long $bytesMD5 4) -bor 1) + 0x13DB0000L
      $map.INDEX = Get-ShiftRight ($length - 2) 1
      $map.COUNTER = $map.INDEX + 1
    
      while ($map.COUNTER) {
        $map.R0 = Convert-Int32 ((Get-Long $bytesBaseInfo $map.PDATA) + [long]$map.OUTHASH1)
        $map.R1[0] = Convert-Int32 (Get-Long $bytesBaseInfo ($map.PDATA + 4))
        $map.PDATA = $map.PDATA + 8
        $map.R2[0] = Convert-Int32 (($map.R0 * ([long]$map.MD51)) - (0x10FA9605L * ((Get-ShiftRight $map.R0 16))))
        $map.R2[1] = Convert-Int32 ((0x79F8A395L * ([long]$map.R2[0])) + (0x689B6B9FL * (Get-ShiftRight $map.R2[0] 16)))
        $map.R3 = Convert-Int32 ((0xEA970001L * $map.R2[1]) - (0x3C101569L * (Get-ShiftRight $map.R2[1] 16) ))
        $map.R4[0] = Convert-Int32 ($map.R3 + $map.R1[0])
        $map.R5[0] = Convert-Int32 ($map.CACHE + $map.R3)
        $map.R6[0] = Convert-Int32 (($map.R4[0] * [long]$map.MD52) - (0x3CE8EC25L * (Get-ShiftRight $map.R4[0] 16)))
        $map.R6[1] = Convert-Int32 ((0x59C3AF2DL * $map.R6[0]) - (0x2232E0F1L * (Get-ShiftRight $map.R6[0] 16)))
        $map.OUTHASH1 = Convert-Int32 ((0x1EC90001L * $map.R6[1]) + (0x35BD1EC9L * (Get-ShiftRight $map.R6[1] 16)))
        $map.OUTHASH2 = Convert-Int32 ([long]$map.R5[0] + [long]$map.OUTHASH1)
        $map.CACHE = ([long]$map.OUTHASH2)
        $map.COUNTER = $map.COUNTER - 1
      }

      [Byte[]] $outHash = @(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
      [byte[]] $buffer = [BitConverter]::GetBytes($map.OUTHASH1)
      $buffer.CopyTo($outHash, 0)
      $buffer = [BitConverter]::GetBytes($map.OUTHASH2)
      $buffer.CopyTo($outHash, 4)
    
      $map = @{PDATA = 0; CACHE = 0; COUNTER = 0 ; INDEX = 0; MD51 = 0; MD52 = 0; OUTHASH1 = 0; OUTHASH2 = 0;
        R0 = 0; R1 = @(0, 0); R2 = @(0, 0); R3 = 0; R4 = @(0, 0); R5 = @(0, 0); R6 = @(0, 0); R7 = @(0, 0)
      }
    
      $map.CACHE = 0
      $map.OUTHASH1 = 0
      $map.PDATA = 0
      $map.MD51 = ((Get-Long $bytesMD5) -bor 1)
      $map.MD52 = ((Get-Long $bytesMD5 4) -bor 1)
      $map.INDEX = Get-ShiftRight ($length - 2) 1
      $map.COUNTER = $map.INDEX + 1

      while ($map.COUNTER) {
        $map.R0 = Convert-Int32 ((Get-Long $bytesBaseInfo $map.PDATA) + ([long]$map.OUTHASH1))
        $map.PDATA = $map.PDATA + 8
        $map.R1[0] = Convert-Int32 ($map.R0 * [long]$map.MD51)
        $map.R1[1] = Convert-Int32 ((0xB1110000L * $map.R1[0]) - (0x30674EEFL * (Get-ShiftRight $map.R1[0] 16)))
        $map.R2[0] = Convert-Int32 ((0x5B9F0000L * $map.R1[1]) - (0x78F7A461L * (Get-ShiftRight $map.R1[1] 16)))
        $map.R2[1] = Convert-Int32 ((0x12CEB96DL * (Get-ShiftRight $map.R2[0] 16)) - (0x46930000L * $map.R2[0]))
        $map.R3 = Convert-Int32 ((0x1D830000L * $map.R2[1]) + (0x257E1D83L * (Get-ShiftRight $map.R2[1] 16)))
        $map.R4[0] = Convert-Int32 ([long]$map.MD52 * ([long]$map.R3 + (Get-Long $bytesBaseInfo ($map.PDATA - 4))))
        $map.R4[1] = Convert-Int32 ((0x16F50000L * $map.R4[0]) - (0x5D8BE90BL * (Get-ShiftRight $map.R4[0] 16)))
        $map.R5[0] = Convert-Int32 ((0x96FF0000L * $map.R4[1]) - (0x2C7C6901L * (Get-ShiftRight $map.R4[1] 16)))
        $map.R5[1] = Convert-Int32 ((0x2B890000L * $map.R5[0]) + (0x7C932B89L * (Get-ShiftRight $map.R5[0] 16)))
        $map.OUTHASH1 = Convert-Int32 ((0x9F690000L * $map.R5[1]) - (0x405B6097L * (Get-ShiftRight ($map.R5[1]) 16)))
        $map.OUTHASH2 = Convert-Int32 ([long]$map.OUTHASH1 + $map.CACHE + $map.R3) 
        $map.CACHE = ([long]$map.OUTHASH2)
        $map.COUNTER = $map.COUNTER - 1
      }
    
      $buffer = [BitConverter]::GetBytes($map.OUTHASH1)
      $buffer.CopyTo($outHash, 8)
      $buffer = [BitConverter]::GetBytes($map.OUTHASH2)
      $buffer.CopyTo($outHash, 12)
    
      [Byte[]] $outHashBase = @(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
      $hashValue1 = ((Get-Long $outHash 8) -bxor (Get-Long $outHash))
      $hashValue2 = ((Get-Long $outHash 12) -bxor (Get-Long $outHash 4))
    
      $buffer = [BitConverter]::GetBytes($hashValue1)
      $buffer.CopyTo($outHashBase, 0)
      $buffer = [BitConverter]::GetBytes($hashValue2)
      $buffer.CopyTo($outHashBase, 4)
      $base64Hash = [Convert]::ToBase64String($outHashBase) 
    }

    Write-Output $base64Hash
  }

  Write-Verbose "Getting Hash For $ProgId   $Extension"
  If ($DomainSID.IsPresent) { Write-Verbose  "Use Get-UserSidDomain" } Else { Write-Verbose  "Use Get-UserSid" } 
  $userSid = If ($DomainSID.IsPresent) { Get-UserSidDomain } Else { Get-UserSid } 
  $userExperience = Get-UserExperience
  $userDateTime = Get-HexDateTime
  Write-Debug "UserDateTime: $userDateTime"
  Write-Debug "UserSid: $userSid"
  Write-Debug "UserExperience: $userExperience"

  $baseInfo = "$Extension$userSid$ProgId$userDateTime$userExperience".ToLower()
  Write-Verbose "baseInfo: $baseInfo"

  $progHash = Get-Hash $baseInfo
  Write-Verbose "Hash: $progHash"
  
  #Write AssociationToasts List
  Write-RequiredApplicationAssociationToasts $ProgId $Extension

  #Handle Extension Or Protocol
  if ($Extension.Contains(".")) {
    Write-Verbose "Write Registry Extension: $Extension"
    Write-ExtensionKeys $ProgId $Extension $progHash

  }
  else {
    Write-Verbose "Write Registry Protocol: $Extension"
    Write-ProtocolKeys $ProgId $Extension $progHash
  }

   
  if ($Icon) {
    Write-Verbose  "Set Icon: $Icon"
    Set-Icon $ProgId $Icon
  }

  Update-RegistryChanges 

}

<#
.DESCRIPTION
    Set the default application for a specific protocol. It takes a ProgId, a Protocol, and an optional Icon as parameters. 
    The ProgId is the identifier of the application to be set as default. The Protocol is the protocol for which the application will be set as default. 
    The Icon is an optional parameter that sets the icon for the application.

.NOTES
    Version    : 1.2.0
    Author(s)  : Danyfirex & Dany3j
    Credits    : https://bbs.pediy.com/thread-213954.htm
                 LMongrain - Hash Algorithm PureBasic Version
    License    : MIT License
    Copyright  : 2022 Danysys. <danysys.com>
  
.EXAMPLE
    Set-PTA ChromeHTML http
    Set Google Chrome as Default for http Protocol

.LINK
File Association Management
    #>
function Set-PTA {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [String]
    $ProgId,

    [Parameter(Mandatory = $true)]
    [String]
    $Protocol,
      
    [String]
    $Icon
  )

  Set-FTA -ProgId $ProgId -Protocol $Protocol -Icon $Icon
}

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
function Add-vtsPrinter {
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
      Add-Printer -ConnectionName "\\$Server\$($printer.Name)"
      Write-Output "Printer $($printer.Name) added successfully.`n"
      Write-Host "Name  : $($printer.Name)`nDriver: $($printer.DriverName)`nPort  : $($printer.PortName)`n"
    }
  }
}

<#
.SYNOPSIS
This script downloads, extracts, and installs a printer driver from a given URL.

.DESCRIPTION
The Add-vtsPrinterDriver function downloads a printer driver from a specified URL, extracts it using 7Zip, and installs it on the local machine. 
It also handles the creation of necessary directories and the installation of required software (Chocolatey and 7Zip).

.PARAMETER WorkingDir
The directory where the printer driver will be downloaded and extracted. Default is "C:\temp\PrinterDrivers".

.PARAMETER DriverURL
The URL of the printer driver to be downloaded. This is a mandatory parameter.

.EXAMPLE
Add-vtsPrinterDriver -WorkingDir "C:\temp\MyDrivers" -DriverURL "http://example.com/driver.zip"

This example downloads the printer driver from http://example.com/driver.zip, extracts it to C:\temp\MyDrivers, and installs it.

.NOTES
This script requires administrative privileges to install the printer driver and software dependencies.

.LINK
Print Management
#>
function Add-vtsPrinterDriver {
  param(
    $WorkingDir = "C:\temp\PrinterDrivers",
    [Parameter(Mandatory = $true)]
    $DriverURL
  )
  
  try {
    $FileName = (get-date -f hhmmyy) + (1111..999999 | Get-Random)
  
    if (Test-Path "$WorkingDir\$FileName") {
      try {
        Rename-Item "$WorkingDir\$FileName" "$WorkingDir\$FileName.old" -Force -ErrorAction Stop
      }
      catch {
        Write-Host "Failed to rename $FileName. Error: $_" -ForegroundColor Red
      }
    }
  
    # Make $WorkingDir
    if (!(Test-Path $WorkingDir)) {
      $FileName
      Write-Host "`n➜ Creating $WorkingDir..."
      New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
      if ($?) { Write-Host "✔ Successfully created $WorkingDir`n" -ForegroundColor Green }
    }
          
    # Download Driver
    Write-Host "`n➜ Downloading Driver..."
    Invoke-WebRequest -Uri "$DriverURL" -OutFile "$WorkingDir\$FileName"
    if ($?) { Write-Host "✔ Successfully Downloaded Driver`n" -ForegroundColor Green }
  
          
    # Remove Chocolatey folder if choco.exe doesn't exist
    if (!(Test-Path "C:\ProgramData\chocolatey\choco.exe")) {
      Write-Host "`n➜ Removing Chocolatey folder..."
      Remove-Item "C:\ProgramData\chocolatey" -Force -Recurse -Confirm:$False
      if ($?) { Write-Host "✔ Removed Choloatey folder...`n" -ForegroundColor Green }
      # Install Chocolatey
      Write-Host "`n➜ Installing Chocolatey..."
      Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
      if ($?) { Write-Host "✔ Installed Chocolatey`n" -ForegroundColor Green }
    }
          
    # Install 7Zip using Chocolatey
    if (-not (Test-Path "C:\Program Files\7-Zip\7z.exe")) {
      Write-Host "`n➜ Installing 7Zip using Chocolatey..."
      & "C:\ProgramData\chocolatey\choco.exe" install 7zip.install -y
      if ($?) { Write-Host "✔ Installed 7Zip`n" -ForegroundColor Green }
    }
          
    # Use 7Zip to extract driver
    Write-Host "`n➜ Extracting driver using 7Zip..."
    & "C:\Program Files\7-Zip\7z.exe" x -y -o"$WorkingDir\$(($FileName -split '.') | Select-Object -Last 1)" "$WorkingDir\$FileName" | Out-Null
    if ($?) { Write-Host "✔ Files extracted`n" -ForegroundColor Green }
          
  
    $InfFiles = Get-ChildItem -Path "$WorkingDir\$(($FileName -split '.') | Select-Object -Last 1)" -Include "*.inf" -Recurse -File | Select-Object -ExpandProperty FullName
  
    if ($InfFiles) {
      Write-Host "`n➜ Adding Drivers..."
      foreach ($Inf in $InfFiles) {
        # Add driver to driver store
        pnputil.exe /a "$Inf" | Out-Null
        if ($?) { Write-Host "✔ Added $Inf to driver store" -ForegroundColor Green }
      }
      
      $InfFileContent = $InfFiles | ForEach-Object { Get-Content $_ }
      
      $InfFileLines = $InfFileContent -split "`n"
      $DriverNames = @()
      foreach ($line in $InfFileLines) {
        if ($line -match '".*=.*') {
          $parts = $line -split "="
          $driverName = $parts[0].Trim()
          if ($driverName -notmatch 'NULL|{|<|Port|\(DOT|http') {
            $DriverNames += ($driverName -replace '"', '')
          }
        }
        if ($line -match '.*=.*"') {
          $parts = $line -split "="
          $driverName = $parts[1].Trim()
          if ($driverName -notmatch 'NULL|{|<|Port|\(DOT|http') {
            $DriverNames += ($driverName -replace '"', '')
          }
        }
      }
      $UniqueDriverNames = $DriverNames | Select-Object -unique
      
      $totalDrivers = $UniqueDriverNames.Count
      $currentDriver = 0
      foreach ($Driver in $UniqueDriverNames) {
        $currentDriver++
        Write-Progress -Activity "Installing printer drivers..." -Status "$([math]::Round(($currentDriver / $totalDrivers) * 100))%" -PercentComplete (($currentDriver / $totalDrivers) * 100)
        Add-PrinterDriver -Name $Driver 2>$null | Out-Null
        if ($?) { Write-Host "✔ Added $Driver driver" -ForegroundColor Green }
      }
      Write-Progress -Activity "Adding printer drivers..." -Completed
    }
    else {
      Write-Host "No .inf files were detected post driver extraction. Consequently, no drivers have been installed." -ForegroundColor Red
    }
  }
  finally {
    Remove-Item "$WorkingDir" -Force -Recurse -confirm:$false
  }
}
<#
.SYNOPSIS
This function retrieves the license details for a list of users after validating their existence.

.DESCRIPTION
The function Get-vts365UserLicense connects to the Graph API, validates if the users exist, and retrieves the license details for a list of users. The user list is passed as a parameter to the function.

.PARAMETER UserList
A single string or comma separated email addresses for which the license details are to be retrieved.

.EXAMPLE
Get-vts365UserLicense -UserList "user1@domain.com, user2@domain.com"

This example validates and retrieves the license details for user1@domain.com and user2@domain.com.

.LINK
M365
#>
function Get-vts365UserLicense {
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All

  $LicenseDetails = @()

  foreach ($User in $UserList) {
    $User = $User.Trim()
    $UserExists = Get-MgUser -UserId $User -ErrorAction SilentlyContinue
    if ($UserExists) {
      $LicenseDetails += [pscustomobject]@{
        User    = $User
        License = (Get-MgUserLicenseDetail -UserId $User | Select-Object -expand SkuPartNumber) -join ", "
      }
    }
    else {
      Write-Host "User $User does not exist." -ForegroundColor Red
    }
  }

  $LicenseDetails
}

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
function Format-vtsMacAddress {
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

<#
.SYNOPSIS
This function copies emails from a specified sender within a date range from all mailboxes to a target mailbox and folder.

.DESCRIPTION
The Copy-vts365MailToMailbox function connects to Exchange Online PowerShell and creates a new compliance search for emails from a specified sender within a specified date range. It waits for the compliance search to complete, gets the results, parses them, and creates objects from the results. It then gets the mailboxes to search, performs the search, and copies the emails to the target mailbox's specified folder.

.EXAMPLE
Copy-vts365MailToMailbox -senderAddress "sender@example.com" -targetMailbox "target@example.com" -targetFolder "Folder" -startDate "01/01/2020" -endDate "12/31/2020" -SearchName "Search1"
This example copies all emails from sender@example.com sent between 01/01/2020 and 12/31/2020 from all mailboxes to the "Folder" in the target@example.com mailbox.

.LINK
M365
#>
function Copy-vts365MailToMailbox {
  param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the sender's email address")]
    [string]$senderAddress,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the target mailbox to copy the results")]
    [string]$targetMailbox,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the target folder to copy the results")]
    [string]$targetFolder,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the start date in the format 'MM/dd/yyyy'")]
    [string]$startDate,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the end date in the format 'MM/dd/yyyy'")]
    [string]$endDate,

    [string]$SearchName = "Copy-vts365MailToMailbox-$(Get-Date -Format MM-dd-yy-mm-ss)"
  )

  # Connect to Exchange Online PowerShell
  Write-Host "Connecting to Exchange Online PowerShell..."
  Connect-ExchangeOnline -ShowBanner:$false
  Write-Host "Connecting to Security & Compliance PowerShell..."
  Connect-IPPSSession -ShowBanner:$false

  # Create a new compliance search
  Write-Host "Creating new compliance search..."
  New-ComplianceSearch -Name "$SearchName" -ExchangeLocation All -ContentMatchQuery "from:$senderAddress AND received>=$startDate AND received<=$endDate" | Out-Null
  Start-ComplianceSearch -Identity "$SearchName" | Out-Null

  # Wait for the compliance search to complete
  Write-Host "Waiting for compliance search to complete..."
  while ((Get-ComplianceSearch -Identity "$SearchName" | Select-Object -ExpandProperty Status) -ne "Completed") {
    Start-Sleep -Seconds 5
    Get-ComplianceSearch -Identity "$SearchName" | Out-Null
  }
    
  # Get the compliance search results
  Write-Host "Getting compliance search results..."
  $Results = Get-ComplianceSearch -Identity "$SearchName" | Select-Object -expand SuccessResults

  # Parse the results
  Write-Host "Parsing results..."
  $array = ($Results -replace "{|}" -split ",").Trim()

  # Create objects from the results
  Write-Host "Creating objects from results..."
  $objects = for ($i = 0; $i -lt $array.Count; $i += 3) {
    New-Object PSObject -Property @{
      Location  = ($array[$i] -split ": ")[1]
      ItemCount = [int]($array[$i + 1] -split ": ")[1]
      TotalSize = [int]($array[$i + 2] -split ": ")[1]
    }
  }

  # Get the mailboxes to search
  Write-Host "Getting mailboxes to search...`n"
  $MailboxesWithContent = $objects | Where-Object ItemCount -gt 0 | Sort-Object ItemCount -Descending | Select-Object Location, ItemCount, TotalSize

  # Initialize a hashtable for mailboxes
  $mailboxTable = @()
  $key = 1

  # Iterate over each mailbox to search
  foreach ($Box in $MailboxesWithContent) {
    # Add each mailbox to the hashtable with its corresponding details
    $mailboxTable += [pscustomobject]@{
      Key       = $key
      Location  = $Box.Location
      ItemCount = $Box.ItemCount
      TotalSize = $Box.TotalSize
    }
    $key++
  }

  # Output the mailbox details
  $mailboxTable

  if ($null -ne $mailboxTable) {
    # Ask the user which mailboxes to search
    $userInput = Read-Host "`nEnter the numbers of the mailboxes you want to search, separated by commas, or enter * to search all mailboxes"
  
    if ($userInput -eq '*') {
      $SelectedMailboxes = $mailboxTable.Location
    }
    else {
      $SelectedMailboxes = $mailboxTable | Where-Object Key -in ($userInput -split ",") | Select-Object -ExpandProperty Location
    }
  
    foreach ($mailbox in $SelectedMailboxes) {
      # Perform the search and copy the emails to the target mailbox's inbox
      Write-Host "Performing search and copying emails from mailbox: $mailbox..."
      Search-Mailbox -Identity $mailbox -SearchQuery "from:$senderAddress AND received>=$startDate AND received<=$endDate" -TargetMailbox $targetMailbox -TargetFolder $targetFolder -LogLevel Full 3>$null | Out-Null
    }
  }
  else { 
    Write-Host "No matches found." -ForegroundColor Red
  }

  Write-Host "Operation completed."
}

<#
.SYNOPSIS
This function adds licenses to a list of users.

.DESCRIPTION
The function Add-vts365UserLicense connects to the Graph API and adds licenses to a list of users. The user list is passed as a parameter to the function. The function also allows the user to select which licenses to add.

.PARAMETER UserList
A single string or comma separated email addresses to which the licenses are to be added.

.EXAMPLE
Add-vts365UserLicense -UserList "user1@domain.com, user2@domain.com"

.LINK
M365
#>
function Add-vts365UserLicense {
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  $UserList = ($UserList -split ",").Trim()

  Write-Host "Connecting to Graph API..."
  Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All

  Write-Host "Getting available SKUs..."
  $SKUs = Get-MgSubscribedSku

  $key = 0
  $LicenseDisplay = @()
  foreach ($item in $SKUs) {
    $key++
    $LicenseDisplay += [PSCustomObject]@{
      Key           = $key
      SkuPartNumber = $item.SkuPartNumber
      SkuID         = $item.SkuID
    }
  }

  Write-Host "Displaying available licenses..."
  $LicenseDisplay | Out-Host

  $userInput = Read-Host "`nEnter the numbers of the licenses you want to add, separated by commas, or enter * to add all available licenses"
  
  if ($userInput -eq '*') {
    Write-Host "Adding all available licenses..."
    $SelectedLicenses = $LicenseDisplay.SkuID
  }
  else {
    Write-Host "Adding selected licenses..."
    $SelectedLicenses = $LicenseDisplay | Where-Object Key -in ($userInput -split ",") | Select-Object -ExpandProperty SkuID
  }
  
  $FormattedSKUs = @()
  foreach ($Sku in $SelectedLicenses) {
    $FormattedSKUs += @{SkuId = $Sku }
  }
  
  $FormattedSKUs
  
  Write-Host "Adding licenses to users..."
  foreach ($User in $UserList) {
    Write-Host "Adding licenses to $User..."
    Set-MgUserLicense -UserId $User -AddLicenses $FormattedSKUs -RemoveLicenses @()
  }
  Write-Host "Operation completed."
}

<#
.SYNOPSIS
This function removes all licenses from a list of users.

.DESCRIPTION
The function Remove-vts365UserLicense connects to the Graph API and removes all licenses from a list of users. The user list is passed as a parameter to the function.

.PARAMETER UserList
A single string or comma separated email addresses from which the licenses are to be removed.

.EXAMPLE
Remove-vts365UserLicense -UserList "user1@domain.com, user2@domain.com"

.LINK
M365
#>
function Remove-vts365UserLicense {
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  $UserList = ($UserList -split ",").Trim()

  Write-Host "Connecting to Graph API..."
  Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All

  Write-Host "Removing licenses from users..."
  foreach ($User in $UserList) {
    Write-Host "Removing licenses from $User..."
    $UserLicenses = Get-MgUserLicenseDetail -UserId $User | Select-Object -ExpandProperty SkuId
    Set-MgUserLicense -UserId $User -AddLicenses @() -RemoveLicenses $UserLicenses
  }
  Write-Host "Operation completed."
}

<#
.SYNOPSIS
Searches Windows Event Logs for a specific term and exports the results to a CSV or HTML report.

.DESCRIPTION
The Search-vtsAllLogs function searches through Windows Event Logs for a specific term. The function allows the user to specify the logs to search, the term to search for, the type of report to generate (CSV or HTML), and the output directory for the report. If a date is provided, the function will only search logs from that date. If the date is set to 'today', the function will search logs from the current date.

.PARAMETER SearchTerm
The term to search for in the logs. This parameter is mandatory.

.PARAMETER ReportType
The type of report to generate. Valid options are 'csv' and 'html'. If not specified, the function will default to 'html' for results with less than or equal to 150 entries, and 'csv' for results with more than 150 entries.

.PARAMETER OutputDirectory
The directory where the report will be saved. If not specified, the function will default to 'C:\temp'.

.PARAMETER Date
The date of the logs to search. The date should be entered in the format: month/day/year. For example, 12/14/2023. You can also enter 'today' as a value to search logs from the current date.

.EXAMPLE
Search-vtsAllLogs -SearchTerm "Error" -ReportType "csv" -OutputDirectory "C:\temp" -Date "12/14/2023"

This example searches all logs from 12/14/2023 for the term "Error" and generates a CSV report in the 'C:\temp' directory.

.EXAMPLE
Search-vtsAllLogs -SearchTerm "Warning" -Date "today"

This example searches all logs from the current date for the term "Warning" and generates a report based on the number of results found.

.LINK
Log Management
#>
function Search-vtsAllLogs {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $SearchTerm,
    [ValidateSet("csv", "html")]
    [string]$ReportType,
    [string]$OutputDirectory = "C:\temp",
    [Parameter(HelpMessage = "Please enter the date in the format: month/day/year. For example, 12/14/2023. You can also enter 'today' as a value.")]
    $Date
  )

  if ($Date -eq "today") {
    $Date = Get-Date -f MM/dd/yy
  }
    
  # Create Output Dir if not exist
  if (!(Test-Path -Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
  }
  
  # Validate the log name
  $validLogNames = (Get-WinEvent -ListLog *).LogName 2>$null
  
  $LogTable = @()
  $key = 1

  foreach ($log in $validLogNames) {
    $LogTable += [pscustomobject]@{
      Key = $key
      Log = $log
    }
    $key++
  }
  
  if ($(whoami) -eq "nt authority\system") {
    $LogTable | Out-Host
    $userInput = Read-Host "Please input the log numbers you wish to search, separated by commas. Alternatively, input '*' to search all logs."
    if ("$userInput" -eq '*') {
      Write-Host "Searching all available logs..."
      $SelectedLogs = $LogTable.Log
    }
    else {
      Write-Host "Searching selected logs..."
      $SelectedLogs = $LogTable | Where-Object Key -in ("$userInput" -split ",") | Select-Object -ExpandProperty Log
    }
  }
  else {
    $SelectedLogs = $LogTable | Out-GridView -OutputMode Multiple | Select-Object -ExpandProperty Log
  }

  # Get the logs from the Event Viewer based on the provided log name
  $Results = @()
  $LogCount = $SelectedLogs.Count * $SearchTerm.Count
  $CurrentLog = 0
  foreach ($LogName in $SelectedLogs) {
    foreach ($Term in $SearchTerm) {
      $CurrentLog++
      Write-Host "Searching $LogName log for $Term... ($CurrentLog/$LogCount)" -ForegroundColor Yellow
      if ($Date) {
        $Results += Get-WinEvent -LogName "$LogName" -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated.Date -eq $Date } |
        Where-Object Message -like "*$Term*" |
        Tee-Object -Variable temp
      }
      else {
        $Results += Get-WinEvent -LogName "$LogName" -ErrorAction SilentlyContinue |
        Where-Object Message -like "*$Term*" |
        Tee-Object -Variable temp
      }
    }
  }
  
  if ("" -eq $ReportType) {
    if ($Results.Count -le 150) {
      $ReportType = "html"
    }
    else {
      $ReportType = "csv"
    }
  }
  
  $ReportPath = "C:\temp\$($env:COMPUTERNAME)-$(Get-Date -Format MM-dd-yy-mm-ss).$ReportType"

  switch ($ReportType) {
    csv { 
      $Results | Export-Csv $ReportPath -NoTypeInformation
      $openFile = Read-Host "Do you want to open the file? (Y/N)"
      if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ReportPath
      }
    }
    html { 
      # Check if PSWriteHTML module is installed, if not, install it
      if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
        Install-Module -Name PSWriteHTML -Force -Confirm:$false
      }
      
      # Export the results to an HTML file using the PSWriteHTML module
      $Results | Out-HtmlView -FilePath $ReportPath

    }
    Default {
      Write-Host "Invalid ReportType selection. Defaulting to csv."
      $Results | Export-Csv $ReportPath
      $openFile = Read-Host "Do you want to open the file? (Y/N)"
      if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ReportPath
      }
    }
  }
  
  Write-Host "Report saved at $ReportPath"
}
<#
.SYNOPSIS
    Get-vtsNICThroughput is a function that measures the throughput of active network adapters.

.DESCRIPTION
    This function continuously measures the throughput (in Mbps) of all active network adapters on the system. 
    It does this by capturing the initial and final statistics of each network adapter over a 2-second interval, 
    and then calculates the difference in received and sent bytes to determine the throughput.

.PARAMETER adapterName
    The name of the network adapter. If not specified, the function will measure the throughput for all active network adapters.

.EXAMPLE
    PS C:\> Get-vtsNICThroughput

    This command will start the continuous measurement of throughput for all active network adapters.

.EXAMPLE
    PS C:\> Get-vtsNICThroughput -AdapterName "Ethernet"

    This command will start the continuous measurement of throughput for the network adapter named "Ethernet".

.NOTES
    To stop the continuous measurement, use Ctrl+C.

.LINK
    Network
#>
function Get-vtsNICThroughput {
  [CmdletBinding()]
  param (
    $AdapterName = (Get-NetAdapter | Where-Object Status -eq Up | Select-Object -expand Name)
  )

  function CalculateNetworkAdapterThroughput {
    param (
      [Parameter(Mandatory = $true)]
      [string]$adapterName
    )

    # Capture initial statistics of the network adapter
    $statsInitial = Get-NetAdapterStatistics -Name $adapterName

    # Wait for 2 seconds to capture the final statistics
    Start-Sleep -Seconds 1
      
    # Capture final statistics of the network adapter
    $statsFinal = Get-NetAdapterStatistics -Name $adapterName
      
    # Calculate the differences in received and sent bytes
    $bytesReceivedDiff = $statsFinal.ReceivedBytes - $statsInitial.ReceivedBytes
    $bytesSentDiff = $statsFinal.SentBytes - $statsInitial.SentBytes
      
    # Calculate the throughput in Mbps
    $throughputInMbps = [Math]::Round($bytesReceivedDiff * 8 / 1MB / 2, 2)
    $throughputOutMbps = [Math]::Round($bytesSentDiff * 8 / 1MB / 2, 2)
      
    Clear-Host

    # Display the throughput
    Write-Host "Adapter: $adapterName"
    Write-Host "    Throughput In (Mbps): $throughputInMbps"
    Write-Host "    Throughput Out (Mbps): $throughputOutMbps"
  }

  # Infinite loop to continuously measure NIC throughput until Ctrl-C is pressed
  while ($true) {
    # Call CalculateNetworkAdapterThroughput function for each adapterName
    foreach ($adapter in $AdapterName) {
      CalculateNetworkAdapterThroughput -adapterName $adapter
    }
  }
}

<#
.SYNOPSIS
This function retrieves mailbox statistics for a list of email addresses from Exchange Online.

.DESCRIPTION
The Get-vts365MailboxStatistics function connects to Exchange Online and retrieves mailbox statistics for each email address provided. The results are stored in an array and outputted at the end. It also provides an option to export the results to an HTML report.

.PARAMETER EmailAddress
An array of email addresses for which to retrieve mailbox statistics. If not provided, the function retrieves statistics for all mailboxes.

.EXAMPLE
PS C:\> Get-vts365MailboxStatistics -EmailAddress "user1@example.com", "user2@example.com"

This example retrieves mailbox statistics for user1@example.com and user2@example.com.

.EXAMPLE
PS C:\> Get-vts365MailboxStatistics

This example retrieves mailbox statistics for all mailboxes.

.LINK
M365
#>
function Get-vts365MailboxStatistics {
  param(
    $EmailAddress = $(get-mailbox | Select-Object -expand UserPrincipalName)
  )

  # Connect to Exchange Online
  Write-Host "Connecting to Exchange Online..."
  Connect-ExchangeOnline

  # Initialize results array
  $Results = @()

  # Retrieve mailbox statistics for each email address
  foreach ($Address in $EmailAddress) {
    Write-Host "Retrieving mailbox statistics for $Address..."
    $Results += Get-EXOMailboxStatistics -UserPrincipalName $Address
  }

  # Output the results
  Write-Host "Retrieval complete. Here are the results:"
  $Results | Sort-Object TotalItemSize -Descending | Format-Table -AutoSize

  # Ask user if they want to export a report
  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"
  
  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
    # Check if PSWriteHTML module is installed, if not, install it
    if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
      Install-Module -Name PSWriteHTML -Force -Confirm:$false
    }
      
    # Export the results to an HTML file using the PSWriteHTML module
    $Results | Out-HtmlView
  }
}

<#
.SYNOPSIS
    The Start-vtsScreenRecording function starts a screen recording of the current desktop session.

.DESCRIPTION
    The Start-vtsScreenRecording function initiates a screen recording using FFmpeg. It creates a new scheduled task that runs a PowerShell script to start the recording. The recording captures the screen at a rate of 5 frames per second and scales the output to a resolution of 1280x720. The output is saved as an MKV file in the "C:\Windows\Temp\VTS\rc" directory. The filename is formatted as "T[timestamp]-[computername]-[username].mkv".

.PARAMETER No Parameters

.EXAMPLE
    PS C:\> Start-vtsScreenRecording

    This command starts a screen recording of the current desktop session.

.NOTES
    This function requires FFmpeg to be installed and accessible in the system path. If FFmpeg is not found, it will attempt to download and install it using Chocolatey and aria2.

.LINK
    Utilities
#>
function Start-vtsScreenRecording {
  Set-ExecutionPolicy Unrestricted
  if ((Test-Path "C:\Windows\Temp\VTS\rc\start.ps1")){
    Remove-item "C:\Windows\Temp\VTS\rc\start.ps1" -Force -Confirm:$false
  }
  if (-not (Test-Path "C:\Windows\Temp\VTS\rc")) {
    mkdir "C:\Windows\Temp\VTS\rc"
  }

  # Set ACL and NTFS permissions for everyone to have full control
  $acl = Get-Acl "C:\Windows\Temp"
  $permission = "Everyone","FullControl","Allow"
  $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
  $acl.SetAccessRule($accessRule)
  
  # Apply to the parent folder
  Set-Acl "C:\Windows\Temp" $acl
  
  # Apply to all child items
  Get-ChildItem "C:\Windows\Temp" -Recurse | ForEach-Object {
      Set-Acl -Path $_.FullName -AclObject $acl
  }

  if ((Get-ScheduledTask -TaskName "RecordSession")){Unregister-ScheduledTask -TaskName "RecordSession" -Confirm:$false}
# Create a new action that runs the PowerShell script with parameters
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-File C:\Windows\Temp\VTS\rc\start.ps1"
# Set the trigger to logon
$trigger = New-ScheduledTaskTrigger -AtLogon
# Register the scheduled task with highest privileges as SYSTEM user
Register-ScheduledTask -Action $action -Trigger $trigger -User "SYSTEM" -TaskName "RecordSession" -RunLevel Highest

@'
$script:source = @"
using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;

namespace RunAsUser
{
    internal class NativeHelpers
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct LUID
        {
            public int LowPart;
            public int HighPart;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct LUID_AND_ATTRIBUTES
        {
            public LUID Luid;
            public PrivilegeAttributes Attributes;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_INFORMATION
        {
            public IntPtr hProcess;
            public IntPtr hThread;
            public int dwProcessId;
            public int dwThreadId;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct STARTUPINFO
        {
            public int cb;
            public String lpReserved;
            public String lpDesktop;
            public String lpTitle;
            public uint dwX;
            public uint dwY;
            public uint dwXSize;
            public uint dwYSize;
            public uint dwXCountChars;
            public uint dwYCountChars;
            public uint dwFillAttribute;
            public uint dwFlags;
            public short wShowWindow;
            public short cbReserved2;
            public IntPtr lpReserved2;
            public IntPtr hStdInput;
            public IntPtr hStdOutput;
            public IntPtr hStdError;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct TOKEN_PRIVILEGES
        {
            public int PrivilegeCount;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
            public LUID_AND_ATTRIBUTES[] Privileges;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct WTS_SESSION_INFO
        {
            public readonly UInt32 SessionID;
            [MarshalAs(UnmanagedType.LPStr)]
            public readonly String pWinStationName;
            public readonly WTS_CONNECTSTATE_CLASS State;
        }
        public struct SECURITY_ATTRIBUTES
        {
            public Int32 nLength;
            public IntPtr lpSecurityDescriptor;
            public int bInheritHandle;
        }
    }
    internal class NativeMethods
    {
        [DllImport("kernel32", SetLastError = true)]
        public static extern int WaitForSingleObject(
          IntPtr hHandle,
          int dwMilliseconds);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool CloseHandle(
            IntPtr hSnapshot);
        [DllImport("userenv.dll", SetLastError = true)]
        public static extern bool CreateEnvironmentBlock(
            ref IntPtr lpEnvironment,
            SafeHandle hToken,
            bool bInherit);
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool CreateProcessAsUserW(
            SafeHandle hToken,
            String lpApplicationName,
            StringBuilder lpCommandLine,
            IntPtr lpProcessAttributes,
            IntPtr lpThreadAttributes,
            bool bInheritHandle,
            uint dwCreationFlags,
            IntPtr lpEnvironment,
            String lpCurrentDirectory,
            ref NativeHelpers.STARTUPINFO lpStartupInfo,
            out NativeHelpers.PROCESS_INFORMATION lpProcessInformation);
        [DllImport("userenv.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DestroyEnvironmentBlock(
            IntPtr lpEnvironment);
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool DuplicateTokenEx(
            SafeHandle ExistingTokenHandle,
            uint dwDesiredAccess,
            IntPtr lpThreadAttributes,
            SECURITY_IMPERSONATION_LEVEL ImpersonationLevel,
            TOKEN_TYPE TokenType,
            out SafeNativeHandle DuplicateTokenHandle);
        [DllImport("kernel32")]
        public static extern IntPtr GetCurrentProcess();
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool GetTokenInformation(
            SafeHandle TokenHandle,
            uint TokenInformationClass,
            SafeMemoryBuffer TokenInformation,
            int TokenInformationLength,
            out int ReturnLength);
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool LookupPrivilegeName(
            string lpSystemName,
            ref NativeHelpers.LUID lpLuid,
            StringBuilder lpName,
            ref Int32 cchName);
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool OpenProcessToken(
            IntPtr ProcessHandle,
            TokenAccessLevels DesiredAccess,
            out SafeNativeHandle TokenHandle);
        [DllImport("wtsapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool WTSEnumerateSessions(
            IntPtr hServer,
            int Reserved,
            int Version,
            ref IntPtr ppSessionInfo,
            ref int pCount);
        [DllImport("wtsapi32.dll")]
        public static extern void WTSFreeMemory(
            IntPtr pMemory);
        [DllImport("kernel32.dll")]
        public static extern uint WTSGetActiveConsoleSessionId();
        [DllImport("Wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSQueryUserToken(
            uint SessionId,
            out SafeNativeHandle phToken);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr CreatePipe(
            ref IntPtr hReadPipe,
            ref IntPtr hWritePipe,
            ref NativeHelpers.SECURITY_ATTRIBUTES lpPipeAttributes,
            Int32 nSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetHandleInformation(
            IntPtr hObject,
            int dwMask,
            int dwFlags);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool ReadFile(
            IntPtr hFile,
            byte[] lpBuffer,
            int nNumberOfBytesToRead,
            ref int lpNumberOfBytesRead,
            IntPtr lpOverlapped/*IntPtr.Zero*/);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool PeekNamedPipe(
            IntPtr handle,
            byte[] buffer,
            uint nBufferSize,
            ref uint bytesRead,
            ref uint bytesAvail,
            ref uint BytesLeftThisMessage);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DuplicateHandle(IntPtr hSourceProcessHandle,
           ushort hSourceHandle, IntPtr hTargetProcessHandle, out IntPtr lpTargetHandle,
           uint dwDesiredAccess, [MarshalAs(UnmanagedType.Bool)] bool bInheritHandle, uint dwOptions);

    }
    internal class SafeMemoryBuffer : SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafeMemoryBuffer(int cb) : base(true)
        {
            base.SetHandle(Marshal.AllocHGlobal(cb));
        }
        public SafeMemoryBuffer(IntPtr handle) : base(true)
        {
            base.SetHandle(handle);
        }
        protected override bool ReleaseHandle()
        {
            Marshal.FreeHGlobal(handle);
            return true;
        }
    }
    internal class SafeNativeHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafeNativeHandle() : base(true) { }
        public SafeNativeHandle(IntPtr handle) : base(true) { this.handle = handle; }
        protected override bool ReleaseHandle()
        {
            return NativeMethods.CloseHandle(handle);
        }
    }
    internal enum SECURITY_IMPERSONATION_LEVEL
    {
        SecurityAnonymous = 0,
        SecurityIdentification = 1,
        SecurityImpersonation = 2,
        SecurityDelegation = 3,
    }
    internal enum SW
    {
        SW_HIDE = 0,
        SW_SHOWNORMAL = 1,
        SW_NORMAL = 1,
        SW_SHOWMINIMIZED = 2,
        SW_SHOWMAXIMIZED = 3,
        SW_MAXIMIZE = 3,
        SW_SHOWNOACTIVATE = 4,
        SW_SHOW = 5,
        SW_MINIMIZE = 6,
        SW_SHOWMINNOACTIVE = 7,
        SW_SHOWNA = 8,
        SW_RESTORE = 9,
        SW_SHOWDEFAULT = 10,
        SW_MAX = 10
    }
    internal enum TokenElevationType
    {
        TokenElevationTypeDefault = 1,
        TokenElevationTypeFull,
        TokenElevationTypeLimited,
    }
    internal enum TOKEN_TYPE
    {
        TokenPrimary = 1,
        TokenImpersonation = 2
    }
    internal enum WTS_CONNECTSTATE_CLASS
    {
        WTSActive,
        WTSConnected,
        WTSConnectQuery,
        WTSShadow,
        WTSDisconnected,
        WTSIdle,
        WTSListen,
        WTSReset,
        WTSDown,
        WTSInit
    }
    [Flags]
    public enum PrivilegeAttributes : uint
    {
        Disabled = 0x00000000,
        EnabledByDefault = 0x00000001,
        Enabled = 0x00000002,
        Removed = 0x00000004,
        UsedForAccess = 0x80000000,
    }
    public class Win32Exception : System.ComponentModel.Win32Exception
    {
        private string _msg;
        public Win32Exception(string message) : this(Marshal.GetLastWin32Error(), message) { }
        public Win32Exception(int errorCode, string message) : base(errorCode)
        {
            _msg = String.Format("{0} ({1}, Win32ErrorCode {2} - 0x{2:X8})", message, base.Message, errorCode);
        }
        public override string Message { get { return _msg; } }
        public static explicit operator Win32Exception(string message) { return new Win32Exception(message); }
    }
    public static class ProcessExtensions
    {
        #region Win32 Constants
        private const int CREATE_UNICODE_ENVIRONMENT = 0x00000400;
        private const int CREATE_NO_WINDOW = 0x08000000;
        private const int CREATE_NEW_CONSOLE = 0x00000010;
        private const uint INVALID_SESSION_ID = 0xFFFFFFFF;
        private static readonly IntPtr WTS_CURRENT_SERVER_HANDLE = IntPtr.Zero;
        private const int HANDLE_FLAG_INHERIT = 0x00000001;
        private const int STARTF_USESTDHANDLES = 0x00000100;
        #endregion
        // Gets the user token from the currently active session
        private static SafeNativeHandle GetSessionUserToken(bool elevated)
        {
            var activeSessionId = INVALID_SESSION_ID;
            var pSessionInfo = IntPtr.Zero;
            var sessionCount = 0;
            // Get a handle to the user access token for the current active session.
            if (NativeMethods.WTSEnumerateSessions(WTS_CURRENT_SERVER_HANDLE, 0, 1, ref pSessionInfo, ref sessionCount))
            {
                try
                {
                    var arrayElementSize = Marshal.SizeOf(typeof(NativeHelpers.WTS_SESSION_INFO));
                    var current = pSessionInfo;
                    for (var i = 0; i < sessionCount; i++)
                    {
                        var si = (NativeHelpers.WTS_SESSION_INFO)Marshal.PtrToStructure(
                            current, typeof(NativeHelpers.WTS_SESSION_INFO));
                        current = IntPtr.Add(current, arrayElementSize);
                        if (si.State == WTS_CONNECTSTATE_CLASS.WTSActive)
                        {
                            activeSessionId = si.SessionID;
                            break;
                        }
                    }
                }
                finally
                {
                    NativeMethods.WTSFreeMemory(pSessionInfo);
                }
            }
            // If enumerating did not work, fall back to the old method
            if (activeSessionId == INVALID_SESSION_ID)
            {
                activeSessionId = NativeMethods.WTSGetActiveConsoleSessionId();
            }
            SafeNativeHandle hImpersonationToken;
            if (!NativeMethods.WTSQueryUserToken(activeSessionId, out hImpersonationToken))
            {
                throw new Win32Exception("WTSQueryUserToken failed to get access token.");
            }
            using (hImpersonationToken)
            {
                // First see if the token is the full token or not. If it is a limited token we need to get the
                // linked (full/elevated token) and use that for the CreateProcess task. If it is already the full or
                // default token then we already have the best token possible.
                TokenElevationType elevationType = GetTokenElevationType(hImpersonationToken);
                if (elevationType == TokenElevationType.TokenElevationTypeLimited && elevated == true)
                {
                    using (var linkedToken = GetTokenLinkedToken(hImpersonationToken))
                        return DuplicateTokenAsPrimary(linkedToken);
                }
                else
                {
                    return DuplicateTokenAsPrimary(hImpersonationToken);
                }
            }
        }

        private static IntPtr out_read;
        private static IntPtr out_write;
        private static IntPtr err_read;
        private static IntPtr err_write;
        private static int BUFSIZE = 4096;
        public static string StartProcessAsCurrentUser(string appPath, string cmdLine = null, string workDir = null, bool visible = true, int wait = -1, bool elevated = true, bool redirectOutput = true)
        {
            NativeHelpers.SECURITY_ATTRIBUTES saAttr = new NativeHelpers.SECURITY_ATTRIBUTES();
            saAttr.nLength = Marshal.SizeOf(typeof(NativeHelpers.SECURITY_ATTRIBUTES));
            saAttr.bInheritHandle = 0x1;
            saAttr.lpSecurityDescriptor = IntPtr.Zero;
            if (redirectOutput)
            {
                NativeMethods.CreatePipe(ref out_read, ref out_write, ref saAttr, 0);
                NativeMethods.CreatePipe(ref err_read, ref err_write, ref saAttr, 0);
                NativeMethods.SetHandleInformation(out_read, HANDLE_FLAG_INHERIT, 0);
                NativeMethods.SetHandleInformation(err_read, HANDLE_FLAG_INHERIT, 0);
            }

            var startInfo = new NativeHelpers.STARTUPINFO();
            startInfo.cb = Marshal.SizeOf(startInfo);
            uint dwCreationFlags = CREATE_UNICODE_ENVIRONMENT | (uint)(visible ? CREATE_NEW_CONSOLE : CREATE_NO_WINDOW);
            startInfo.wShowWindow = (short)(visible ? SW.SW_SHOW : SW.SW_HIDE);
            startInfo.hStdOutput = out_write;
            startInfo.hStdError = err_write;
            startInfo.dwFlags |= (uint)STARTF_USESTDHANDLES;

            StringBuilder commandLine = new StringBuilder(cmdLine);
            var procInfo = new NativeHelpers.PROCESS_INFORMATION();

            using (var hUserToken = GetSessionUserToken(elevated))
            {
                IntPtr pEnv = IntPtr.Zero;
                if (!NativeMethods.CreateEnvironmentBlock(ref pEnv, hUserToken, false))
                {
                    throw new Win32Exception("CreateEnvironmentBlock failed.");
                }

                try
                {
                    if (!NativeMethods.CreateProcessAsUserW(hUserToken,
                        appPath, // Application Name
                        commandLine, // Command Line
                        IntPtr.Zero,
                        IntPtr.Zero,
                        redirectOutput,
                        dwCreationFlags,
                        pEnv,
                        workDir, // Working directory
                        ref startInfo,
                        out procInfo))
                    {
                        throw new Win32Exception("CreateProcessAsUser failed.");
                    }
                    try
                    {
                        NativeMethods.WaitForSingleObject(procInfo.hProcess, wait);
                    }
                    finally
                    {
                        NativeMethods.CloseHandle(procInfo.hThread);
                        NativeMethods.CloseHandle(procInfo.hProcess);
                    }
                }
                finally
                {
                    NativeMethods.DestroyEnvironmentBlock(pEnv);
                }
            }
            if (redirectOutput)
            {
                var sb = new StringBuilder();
                byte[] buf = new byte[BUFSIZE];
                int dwRead = 0;
                while (true)
                {
                    if (Readable(out_read))
                    {
                        bool bSuccess = NativeMethods.ReadFile(out_read, buf, BUFSIZE, ref dwRead, IntPtr.Zero);
                        if (!bSuccess || dwRead == 0)
                            break;
                        sb.AppendLine(Encoding.Default.GetString(buf).TrimEnd(new char[] { (char)0 }));
                    }
                    else
                    {
                        break;
                    }
                }

                NativeMethods.CloseHandle(out_read);
                NativeMethods.CloseHandle(err_read);
                NativeMethods.CloseHandle(out_write);
                NativeMethods.CloseHandle(err_write);

                return sb.ToString();
            }
            else
            {
                return procInfo.dwProcessId.ToString();
            }
        }

        private static bool Readable(IntPtr streamHandle)
        {
            byte[] aPeekBuffer = new byte[1];
            uint aPeekedBytes = 0;
            uint aAvailBytes = 0;
            uint aLeftBytes = 0;

            bool aPeekedSuccess = NativeMethods.PeekNamedPipe(
                streamHandle,
                aPeekBuffer, 1,
                ref aPeekedBytes, ref aAvailBytes, ref aLeftBytes);

            if (aPeekedSuccess && aPeekBuffer[0] != 0)
                return true;
            else
                return false;
        }
        private static SafeNativeHandle DuplicateTokenAsPrimary(SafeHandle hToken)
        {
            SafeNativeHandle pDupToken;
            if (!NativeMethods.DuplicateTokenEx(hToken, 0, IntPtr.Zero, SECURITY_IMPERSONATION_LEVEL.SecurityImpersonation,
                TOKEN_TYPE.TokenPrimary, out pDupToken))
            {
                throw new Win32Exception("DuplicateTokenEx failed.");
            }
            return pDupToken;
        }
        public static Dictionary<String, PrivilegeAttributes> GetTokenPrivileges()
        {
            Dictionary<string, PrivilegeAttributes> privileges = new Dictionary<string, PrivilegeAttributes>();

            using (SafeNativeHandle hToken = OpenProcessToken(NativeMethods.GetCurrentProcess(), TokenAccessLevels.Query))
            using (SafeMemoryBuffer tokenInfo = GetTokenInformation(hToken, 3))
            {
                NativeHelpers.TOKEN_PRIVILEGES privilegeInfo = (NativeHelpers.TOKEN_PRIVILEGES)Marshal.PtrToStructure(
                    tokenInfo.DangerousGetHandle(), typeof(NativeHelpers.TOKEN_PRIVILEGES));

                IntPtr ptrOffset = IntPtr.Add(tokenInfo.DangerousGetHandle(), Marshal.SizeOf(privilegeInfo.PrivilegeCount));
                for (int i = 0; i < privilegeInfo.PrivilegeCount; i++)
                {
                    NativeHelpers.LUID_AND_ATTRIBUTES info = (NativeHelpers.LUID_AND_ATTRIBUTES)Marshal.PtrToStructure(ptrOffset,
                        typeof(NativeHelpers.LUID_AND_ATTRIBUTES));

                    int nameLen = 0;
                    NativeHelpers.LUID privLuid = info.Luid;
                    NativeMethods.LookupPrivilegeName(null, ref privLuid, null, ref nameLen);

                    StringBuilder name = new StringBuilder(nameLen + 1);
                    if (!NativeMethods.LookupPrivilegeName(null, ref privLuid, name, ref nameLen))
                    {
                        throw new Win32Exception("LookupPrivilegeName() failed");
                    }

                    privileges[name.ToString()] = info.Attributes;

                    ptrOffset = IntPtr.Add(ptrOffset, Marshal.SizeOf(typeof(NativeHelpers.LUID_AND_ATTRIBUTES)));
                }
            }

            return privileges;
        }
        private static TokenElevationType GetTokenElevationType(SafeHandle hToken)
        {
            using (SafeMemoryBuffer tokenInfo = GetTokenInformation(hToken, 18))
            {
                return (TokenElevationType)Marshal.ReadInt32(tokenInfo.DangerousGetHandle());
            }
        }
        private static SafeNativeHandle GetTokenLinkedToken(SafeHandle hToken)
        {
            using (SafeMemoryBuffer tokenInfo = GetTokenInformation(hToken, 19))
            {
                return new SafeNativeHandle(Marshal.ReadIntPtr(tokenInfo.DangerousGetHandle()));
            }
        }
        private static SafeMemoryBuffer GetTokenInformation(SafeHandle hToken, uint infoClass)
        {
            int returnLength;
            bool res = NativeMethods.GetTokenInformation(hToken, infoClass, new SafeMemoryBuffer(IntPtr.Zero), 0,
                out returnLength);
            int errCode = Marshal.GetLastWin32Error();
            if (!res && errCode != 24 && errCode != 122)  // ERROR_INSUFFICIENT_BUFFER, ERROR_BAD_LENGTH
            {
                throw new Win32Exception(errCode, String.Format("GetTokenInformation({0}) failed to get buffer length", infoClass));
            }
            SafeMemoryBuffer tokenInfo = new SafeMemoryBuffer(returnLength);
            if (!NativeMethods.GetTokenInformation(hToken, infoClass, tokenInfo, returnLength, out returnLength))
                throw new Win32Exception(String.Format("GetTokenInformation({0}) failed", infoClass));
            return tokenInfo;
        }
        private static SafeNativeHandle OpenProcessToken(IntPtr process, TokenAccessLevels access)
        {
            SafeNativeHandle hToken = null;
            if (!NativeMethods.OpenProcessToken(process, access, out hToken))
            {
                throw new Win32Exception("OpenProcessToken() failed");
            }
            return hToken;
        }
    }
}
"@

$Public = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue)
foreach ($import in @($Public)) {
    try {
        . $import.FullName
    }
    catch {
        Write-Error -message "$($env:ComputerName) - Failed to import function $($import.FullName): $_"
    }
}


function invoke-ascurrentuser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock,
        [Parameter(Mandatory = $false)]
        [switch]$NoWait,
        [Parameter(Mandatory = $false)]
        [switch]$UseWindowsPowerShell,
        [Parameter(Mandatory = $false)]
        [switch]$UseMicrosoftPowerShell,
        [Parameter(Mandatory = $false)]
        [switch]$NonElevatedSession,
        [Parameter(Mandatory = $false)]
        [switch]$Visible,
        [Parameter(Mandatory = $false)]
        [switch]$CacheToDisk,
        [Parameter(Mandatory = $false)]
        [switch]$CaptureOutput
    )
    if (!("RunAsUser.ProcessExtensions" -as [type])) {
        Add-Type -TypeDefinition $script:source -Language CSharp
    }
    if ($CacheToDisk) {
        $ScriptGuid = new-guid
        $null = New-item "$($ENV:TEMP)\$($ScriptGuid).ps1" -Value $ScriptBlock -Force
        $pwshcommand = "-ExecutionPolicy Bypass -Window Normal -file `"$($ENV:TEMP)\$($ScriptGuid).ps1`""
    }
    else {
        $encodedcommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($ScriptBlock))
        $pwshcommand = "-ExecutionPolicy Bypass -Window Normal -EncodedCommand $($encodedcommand)"
    }
    $OSLevel = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentVersion
    if ($OSLevel -lt 6.2) { $MaxLength = 8190 } else { $MaxLength = 32767 }
    if ($encodedcommand.length -gt $MaxLength -and $CacheToDisk -eq $false) {
        Write-Error -message "$($env:ComputerName) - The encoded script is longer than the command line parameter limit. Please execute the   
script with the -CacheToDisk option."
        return
    }
    if ($UseMicrosoftPowerShell -and -not (Test-Path -Path "$env:ProgramFiles\PowerShell\7\pwsh.exe")) {
        Write-Host -message "$($env:ComputerName) - Not able to find Microsoft PowerShell v7 (pwsh.exe). Ensure that it is installed on      
this system"
        return
    }
    $privs = [RunAsUser.ProcessExtensions]::GetTokenPrivileges()['SeDelegateSessionUserImpersonatePrivilege']
    if (-not $privs -or ($privs -band [RunAsUser.PrivilegeAttributes]::Disabled)) {
        Write-Host -message "$($env:ComputerName) - Not running with correct privilege. You must run this script as system or have the       
SeDelegateSessionUserImpersonatePrivilege token."
        return
    }
    else {
        try {
            # Use the same PowerShell executable as the one that invoked the function, Unless -UseWindowsPowerShell or -UseMicrosoftPowerShell is defined.
            $pwshPath = if ($UseWindowsPowerShell) { "$($ENV:windir)\system32\WindowsPowerShell\v1.0\powershell.exe" } 
            elseif ($UseMicrosoftPowerShell) { "$($env:ProgramFiles)\PowerShell\7\pwsh.exe" }
            else { (Get-Process -Id $pid).Path }

            if ($NoWait) { $ProcWaitTime = 1 } else { $ProcWaitTime = -1 }
            if ($NonElevatedSession) { $RunAsAdmin = $false } else { $RunAsAdmin = $true }
            [RunAsUser.ProcessExtensions]::StartProcessAsCurrentUser(
                $pwshPath, "`"$pwshPath`" $pwshcommand",
                (Split-Path $pwshPath -Parent), $Visible, $ProcWaitTime, $RunAsAdmin, $CaptureOutput)
            if ($CacheToDisk) { $null = remove-item "$($ENV:TEMP)\$($ScriptGuid).ps1" -Force }
        }
        catch {
            Write-Host -message "$($env:ComputerName) - Could not execute as currently logged on user: $($_.Exception.Message)"
            return
        }
    }
}


# SYSTEM
  if (-not (Test-Path "C:\ProgramData\chocolatey\choco.exe")) {
    Remove-item "C:\ProgramData\chocolatey" -Recurse -Force
  }
  if (-not (Test-path "C:\ProgramData\chocolatey\bin\aria2c.exe")) {
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    Choco install aria2 -y
  }

# USER
  invoke-ascurrentuser -scriptblock {
    if (-not (Test-Path "C:\Windows\Temp\VTS\rc\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe")) {
      Set-Location "C:\Windows\Temp\VTS\rc"
      aria2c -x16 -s16 -k1M -c -o ffmpeg.zip "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl-shared.zip" --file-allocation=none
      Expand-Archive -Path "C:\Windows\Temp\VTS\rc\ffmpeg.zip" -DestinationPath "C:\Windows\Temp\VTS\rc" -Force
      $acl = Get-Acl "C:\Windows\Temp"
      $permission = "Everyone","FullControl","Allow"
      $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
      $acl.SetAccessRule($accessRule)
      
      # Apply to the parent folder
      Set-Acl "C:\Windows\Temp" $acl
      
      # Apply to all child items
      Get-ChildItem "C:\Windows\Temp" -Recurse | ForEach-Object {
          Set-Acl -Path $_.FullName -AclObject $acl
      }
    }

    $script = {
      Start-Job -Name RecordScreen -ScriptBlock {
        while ($true){
          if (-not (Get-Process ffmpeg)){
            & "C:\Windows\Temp\VTS\rc\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe" -f gdigrab -framerate 5 -t 1800 -i desktop -vf "scale=1280:720" "C:\Windows\Temp\VTS\rc\T$(Get-Date -f hhmm-MM-dd-yyyy)-$($env:COMPUTERNAME)-$($env:USERNAME).mkv"
          }
        }
      }
    }

    Write-Host "Starting screen recording..."
    # Start-Process powershell -ArgumentList "-NoExit", "-Command & {$script}" -WindowStyle Hidden
    Start-Process powershell -ArgumentList "-NoExit", "-Command & {$script}" -WindowStyle Hidden -PassThru

}
'@ | Out-File -FilePath C:\Windows\Temp\VTS\rc\start.ps1 -Force -Encoding utf8
Start-Sleep 5
Start-ScheduledTask -TaskName "RecordSession"
}

<#
.SYNOPSIS
    The Stop-vtsScreenRecording function stops the ongoing screen recording.

.DESCRIPTION
    The Stop-vtsScreenRecording function stops the screen recording initiated by the Start-vtsScreenRecording function. It disables the scheduled task that was created to start the recording and stops the FFmpeg and PowerShell processes that were running the recording.

.PARAMETER No Parameters

.EXAMPLE
    PS C:\> Stop-vtsScreenRecording

    This command stops the ongoing screen recording.

.NOTES
    This function should be used to stop a screen recording that was started with the Start-vtsScreenRecording function.

.LINK
    Utilities
#>
function Stop-vtsScreenRecording {
  Disable-ScheduledTask -TaskName "RecordSession"
  Get-Process ffmpeg, powershell | Stop-Process -Force -Confirm:$false
}

<#
.SYNOPSIS
This function calculates the size of a directory in MB or GB.

.DESCRIPTION
The Get-vtsDirectorySize function calculates the size of a directory and returns the size in MB or GB. If the size is greater than 1024 MB, it will be converted to GB.

.PARAMETER Path
The path of the directory you want to calculate the size of. If no path is provided, the function will calculate the size of the current directory.

.EXAMPLE
Get-vtsDirectorySize -Path "C:\Windows"
This command will calculate the size of the Windows directory.

.LINK
File Management
#>
function Get-vtsDirectorySize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$Path = (Get-Location).Path
    )

    try {
        # Calculate the size of the directory
        $size = (Get-ChildItem $Path -Recurse -ErrorAction Stop | Measure-Object -Property Length -Sum).Sum
        $sizeInMB = "{0:N2}" -f ($size / 1MB)
        $sizeInGB = "{0:N2}" -f ($size / 1GB)

        # Output the size in MB or GB
        if ($sizeInMB -gt 1024) {
            Write-Output ("The size of $Path is " + $sizeInGB + " GB")
        } else {
            Write-Output ("The size of $Path is " + $sizeInMB + " MB")
        }
    }
    catch {
        Write-Error "An error occurred while calculating the size of the directory: $_"
    }
}

<#
.SYNOPSIS
    This function adds a specified user as an admin to all SharePoint sites.

.DESCRIPTION
    The Add-vtsSharePointAdminToAllSites function connects to SharePoint Online using the provided admin site URL, retrieves all site collections, and adds the specified user as an admin to each site.

.PARAMETER adminSiteUrl
    The URL of the SharePoint admin site. This parameter is mandatory.

.PARAMETER userToMakeOwner
    The username of the user to be made an admin. This parameter is mandatory.

.EXAMPLE
    Add-vtsSharePointAdminToAllSites -adminSiteUrl "https://contoso-admin.sharepoint.com" -userToMakeOwner "user@contoso.com"
    This example adds the user "user@contoso.com" as an admin to all SharePoint sites in the "contoso" tenant.

.NOTES
    This function requires the Microsoft.Online.SharePoint.PowerShell module. If the module is not installed, the function will install it.
    Requres PowerShell 5. PowerShell 7 doesn't work.

.LINK
    M365
#>
function Add-vtsSharePointAdminToAllSites {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Enter the SharePoint admin site URL in the format 'https://yourdomain-admin.sharepoint.com'")]
        [string]$adminSiteUrl,

        [Parameter(Mandatory=$true, HelpMessage="Enter the username of the user to be made an admin in the format 'user@yourdomain.com'")]
        [string]$userToMakeOwner
    )
    if (!(Get-InstalledModule Microsoft.Online.SharePoint.PowerShell)) {
        Install-Module Microsoft.Online.SharePoint.PowerShell
    }

    Import-Module Microsoft.Online.SharePoint.PowerShell

    Write-Host "Connecting to SharePoint Online..."
    Connect-SPOService -Url $adminSiteUrl

    Write-Host "Retrieving all site collections..."
    $sites = Get-SPOSite -Limit All

    foreach ($site in $sites) {
        Write-Host "Adding owner to site: " $site.Url
        Set-SPOUser -Site $site.Url -LoginName $userToMakeOwner -IsSiteCollectionAdmin $true

        Write-Host "Successfully added admin to site:" $site.Url
    }
}

<#
.SYNOPSIS
This function retrieves the disconnect time of a session.

.DESCRIPTION
The Get-vtsSessionDisconnectTime function uses the Get-WinEvent cmdlet to retrieve the event logs for a specific EventID from the Microsoft-Windows-TerminalServices-LocalSessionManager/Operational log. It then parses the XML of each event to get the username and the disconnect time. It returns an array of custom objects, each containing a username and a disconnect time.

.PARAMETER EventID
The ID of the event to filter the event logs. The default value is 24.

.PARAMETER LogName
The name of the log to get the event logs from. The default value is "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational".

.EXAMPLE
PS C:\> Get-vtsSessionDisconnectTime

This command retrieves the disconnect time of all VTS sessions.

.NOTES
Additional information about the function.

.LINK
Log Management
#>
function Get-vtsSessionDisconnectTime {
    param(
        [int]$EventID = 24,
        [string]$LogName = "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational"
    )
    
    $events = Get-WinEvent -FilterHashtable @{LogName = $LogName; ID = $EventID }
    $output = @()
    foreach ($event in $events) {
        $xml = [xml]$event.ToXml()
        $eventUsername = $xml.Event.UserData.EventXML.User
        $eventTime = $event.TimeCreated
    
        $output += [pscustomobject]@{
            Username       = $eventUsername
            DisconnectTime = $eventTime
        }
    }
    
    return $output
}

<#
.SYNOPSIS
This function continuously monitors for changes in the connected devices.

.DESCRIPTION
The Get-vtsDeviceChange function uses the Get-PnpDevice cmdlet to continuously scan for devices that are currently connected to the system. It compares the results of two consecutive scans to detect any changes in the connected devices. If a device is connected or removed between the two scans, it will output a message indicating the change.

.PARAMETER None
This function does not take any parameters.

.EXAMPLE
PS C:\> Get-vtsDeviceChange
This command will start the function and continuously monitor for any changes in the connected devices. If a device is connected or removed, it will output a message indicating the change.

.NOTES
To stop the function, use the keyboard shortcut for stopping a running command in your shell (usually Ctrl+C).

.LINK
Device Management
#>
function Get-vtsDeviceChange {
    while ($true) {
        $FirstScan = Get-PnpDevice | Where-Object Present -eq True | Select-Object -expand FriendlyName

        Start-Sleep -Seconds 1

        $SecondScan = Get-PnpDevice | Where-Object Present -eq True | Select-Object -expand FriendlyName

        $Changes = Compare-Object $FirstScan $SecondScan

        foreach ($Device in $Changes) {
            if ($Device.SideIndicator -eq "=>") {
                "`"$($Device.InputObject)`" connected."
            }
            if ($Device.SideIndicator -eq "<=") {
                "`"$($Device.InputObject)`" removed."
            }
        }

        $FirstScan = $SecondScan
    }
}



<#
.SYNOPSIS
Starts a packet capture using Wireshark's tshark.

.DESCRIPTION
The Start-vtsPacketCapture function starts a packet capture on the specified network interface. If Wireshark is not installed, it will install it along with Chocolatey and Npcap.

.PARAMETER interface
The network interface to capture packets from. Defaults to the first Ethernet interface that is up.

.PARAMETER output
The path to the output file. Defaults to a .pcap file in C:\temp with the computer name and current date and time.

.EXAMPLE
Start-vtsPacketCapture -interface "Ethernet" -output "C:\temp\capture.pcap"
Starts a packet capture on the Ethernet interface, with the output saved to C:\temp\capture.pcap.

.LINK
Network
#>
function Start-vtsPacketCapture {
    param (
        [string]$interface = (get-netadapter | Where-Object Name -like "*ethernet*" | Where-Object Status -eq Up | Select-Object -expand name),
        [string]$output = "C:\temp\$($env:COMPUTERNAME)-$(Get-Date -f hhmm-MM-dd-yyyy)-capture.pcap"
    )

    Write-Host "Starting packet capture..."

    if (!(Test-Path "C:\Program Files\Wireshark\tshark.exe")) {
        Write-Host "Wireshark not found. Installing necessary components..."

        #Install choco
        Write-Host "Installing Chocolatey..."
        Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

        #Install Wireshark
        Write-Host "Installing Wireshark..."
        choco install wireshark -y

        #Install npcap
        Write-Host "Installing Npcap..."
        mkdir C:\temp
        Invoke-WebRequest -Uri "https://npcap.com/dist/npcap-1.79.exe" -UseBasicParsing -OutFile "C:\temp\npcap-1.79.exe"

        & "C:\temp\npcap-1.79.exe"

        $wshell = New-Object -ComObject wscript.shell
        Start-sleep -Milliseconds 250
        $wshell.AppActivate('Npcap 1.79 Setup')
        $wshell.SendKeys('%a');
        Start-sleep -Milliseconds 250
        $wshell.AppActivate('Npcap 1.79 Setup')
        $wshell.SendKeys('%i');
        Start-sleep -Milliseconds 250
        $wshell.AppActivate('Npcap 1.79 Setup')
        $wshell.SendKeys('%n');
        Start-sleep -Milliseconds 250
        $wshell.AppActivate('Npcap 1.79 Setup')
        $wshell.SendKeys('{enter}');
    }
    
    #Start packet capture using tshark on ethernet NIC
    $tsharkPath = "C:\Program Files\Wireshark\tshark.exe"
    
    Write-Host "Starting packet capture on interface $interface, output to $output"
    Start-Process -FilePath $tsharkPath -ArgumentList "-i $interface -w $output" -WindowStyle Hidden
}

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
function Stop-vtsPacketCapture {
    Write-Host "Stopping packet capture..."
    Get-Process tshark | Stop-Process -Confirm:$false
    Write-Host "Packet capture stopped."
}
