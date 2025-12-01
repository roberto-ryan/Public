function Search-vtsEventLog {
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

function Get-vtsMappedDrive {
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

function Start-vtsPingReport {
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

function New-vtsRandomPassword {
  <#
  .Description
  Generates a random 12 character password and copies it to the clipboard.
  If the -Easy switch is specified, generates a password using two random words, a random number and a random symbol.
  Now, it also adds a random preposition between the two words for the Easy password.
  .EXAMPLE
  PS> New-vtsRandomPassword
  
  Output:
  Random Password Copied to Clipboard
  
  .EXAMPLE
  PS> New-vtsRandomPassword -Easy -WordListPath "C:\path\to\wordlist.csv"
  
  Output:
  Easy Random Password Copied to Clipboard
  
  .LINK
  Utilities
  #>
  param(
    [switch]$Easy,
    [string]$WordListPath = "$env:temp\wordlist.csv"
  )

  $numbers = 0..9
  $symbols = '!', '@', '#', '$', '%', '*', '?', '+', '='
  $prepositions = 'on', 'in', 'at', 'by', 'up', 'to', 'of', 'off', 'for', 'out', 'via'
  $number = $numbers | Get-Random
  $symbol = $symbols | Get-Random

  if ($Easy) {
    if (!(Test-Path $WordListPath)) {
      Invoke-WebRequest -uri "https://raw.githubusercontent.com/roberto-ryan/Public/main/wordlist.csv" -UseBasicParsing -OutFile $WordListPath
    }
    $words = Import-Csv -Path $WordListPath | ForEach-Object { $_.Word } 
    $randomPreposition = ($prepositions | Get-Random).ToUpper()
    $randomWord1 = $words | Get-Random
    $randomWord2 = $words | Get-Random
    $NewPW = $randomWord1 + $randomPreposition + $randomWord2 + $number + $symbol
  }
  else {
    $string = ( -join ((0x30..0x39) + ( 0x41..0x5A) + ( 0x61..0x7A) |
        Get-Random -Count 12  |
        ForEach-Object { [char]$_ }))
    $number = $numbers | Get-Random
    $symbol = $symbols | Get-Random
    $NewPW = $string + $number + $symbol
  }

  $NewPW | Set-Clipboard

  if ($Easy) {
    Write-Output "Easy Random Password Copied to Clipboard - $NewPW"
  }
  else {
    Write-Output "Random Password Copied to Clipboard - $NewPW"
  }
}

function Out-vtsPhoneticAlphabet {
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

function Get-vtsDisplayDetails {
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


function Install-vtsChromeExtension {
  <#
  .DESCRIPTION
  Forces the installation of a specified Google Chrome extension.
  .EXAMPLE
  PS> Install-vtsChromeExtension -extensionId "ddloeodolhdfbohkokiflfbacbfpjahp"
  
  .LINK
  Package Management
  #>
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

function Get-vtsTemperature {
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

function Get-vtsDiskStat {
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

function Start-vtsSpeedTest {
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
 
function Install-vtsWindowsUpdate {
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

function Search-vtsRDPGatewayLog {
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

function Uninstall-vtsNinja {
  <#
  .DESCRIPTION
  Uninstall ninja and remove keys, files and services.
  .EXAMPLE
  PS> Uninstall-vtsNinja
  
  .LINK
  Package Management
  #>
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

function Show-vtsToastNotification {
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

function New-vtsMappedDrive {
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

function Install-vtsPwsh {
  <#
  .DESCRIPTION
  Install PowerShell 7
  .EXAMPLE
  PS> Install-vtsPwsh
  
  .LINK
  Package Management
  #>
  param (
    [switch]$InstallLatestVersionWithGUI
  )

  if ($InstallLatestVersionWithGUI){
    iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI"
  } else {
    msiexec.exe /i "https://github.com/PowerShell/PowerShell/releases/download/v7.4.6/PowerShell-7.4.6-win-x64.msi" /qn
    Write-Host "Installing PowerShell 7... Please wait" -ForegroundColor Cyan
    While (-not (Test-Path "C:\Program Files\PowerShell\7\pwsh.exe" 2>$null)) { Start-Sleep 5 }
    & "C:\Program Files\PowerShell\7\pwsh.exe" 2>$null
  }
}

function Get-FTA {
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

function Get-PTA {
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

function Register-FTA {
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

function Remove-FTA {
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

function Set-FTA {
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

function Set-PTA {
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

function Add-vtsPrinterDriver {
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
  param(
    $WorkingDir = "C:\temp\PrinterDrivers",
    [Parameter(Mandatory = $true)]
    $DriverURL,
    $LogPath = "C:\temp\PrinterDrivers\DriverDownload.log"
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

    switch -Regex ($DriverURL) {
      "^http" { 
      Invoke-vtsFastDownload -DownloadPath "$WorkingDir" -URL "$DriverURL" -FileName "$FileName" 
      if ($?) { Write-Host "✔ Successfully Downloaded Driver`n" -ForegroundColor Green }
      }
      "^\\\\" { 
      Robocopy.exe "$DriverURL" "$WorkingDir\$FileName" /E /XO /MT:8 /R:3 /W:10 /LOG:"$Logpath" 
      if ($?) { Write-Host "✔ Successfully Downloaded Driver`n" -ForegroundColor Green }
      }
      Default { Write-Host "Link is invalid" }
    }
          
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
    #Remove-Item "$WorkingDir" -Force -Recurse -confirm:$false
  }
}

function Get-vts365UserLicense {
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
  param (
    [Parameter(Mandatory = $false, HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  if (-not $UserList) {
    $UserList = get-user | 
    Where UserPrincipalName -notlike "*.onmicrosoft.com*" | 
    sort DisplayName | 
    select DisplayName, UserPrincipalName | 
    Out-GridView -Title "Select users to manage licenses for" -OutputMode Multiple | 
    Select -expand UserPrincipalName
  }

  if (-not $UserList) {
    Write-Host "No users selected." -ForegroundColor Red
    return
  }

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

  $LicenseDetails | Out-Host

  # Ask user if they want to export a report
  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"
  
  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
    # Check if PSWriteHTML module is installed, if not, install it
    if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
      Install-Module -Name PSWriteHTML -Force -Confirm:$false
    }
        
    # Export the results to an HTML file using the PSWriteHTML module
    $LicenseDetails | Out-HtmlView
  }
}

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

function Copy-vts365MailToMailbox {
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
  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }
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

function Add-vts365UserLicense {
  <#
  .SYNOPSIS
  This function adds licenses to a list of users.
  
  .DESCRIPTION
  The function Add-vts365UserLicense connects to the Graph API and adds licenses to a list of users. If no user list is specified, it connects to Exchange Online and allows the user to select mailboxes dynamically. The function also allows the user to select which licenses to add.
  
  .PARAMETER UserList
  An optional parameter. A single string or comma separated email addresses to which the licenses are to be added. If not specified, the user will be prompted to select from a list of all mailboxes.
  
  .EXAMPLE
  Add-vts365UserLicense -UserList "user1@domain.com, user2@domain.com"
  
  .EXAMPLE
  Add-vts365UserLicense
  
  This example prompts the user to select from a list of all mailboxes and then adds licenses to the selected users.
  
  .LINK
  M365
  #>
  param (
    [Parameter(HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  if (-not $UserList) {
    Write-Host "Connecting to Exchange Online to retrieve mailboxes..."
    Connect-ExchangeOnline | Out-Null
    $UserList = get-user | Where UserPrincipalName -notlike "*.onmicrosoft.com*" | sort DisplayName | select DisplayName, UserPrincipalName | Out-GridView -OutputMode Multiple | Select -expand UserPrincipalName
    if (-not $UserList) {
      Write-Host "No users selected or no mailboxes available." -ForegroundColor Red
      return
    }
  } else {
    # Split the user list and trim whitespace
    $UserList = ($UserList -split ",").Trim()
  }

  Write-Host "Connecting to Graph API..."
  Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All -Device

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
  
  Write-Host "Adding licenses to users..."
  foreach ($User in $UserList) {
    Write-Host "Adding licenses to $User..."
    Set-MgUserLicense -UserId $User -AddLicenses $FormattedSKUs -RemoveLicenses @()
  }
  Write-Host "Operation completed."
}

function Remove-vtsALL365UserLicense {
  <#
  .SYNOPSIS
  This function removes all licenses from a list of users.
  
  .DESCRIPTION
  The function Remove-vtsALL365UserLicense connects to the Graph API and removes all licenses from a list of users. The user list is passed as a parameter to the function.
  
  .PARAMETER UserList
  A single string or comma separated email addresses from which the licenses are to be removed.
  
  .EXAMPLE
  Remove-vtsALL365UserLicense -UserList "user1@domain.com, user2@domain.com"
  
  .LINK
  M365
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  $UserList = ($UserList -split ",").Trim()

  Write-Host "Connecting to Graph API..."
  Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All -Device

  Write-Host "Removing licenses from users..."
  foreach ($User in $UserList) {
    Write-Host "Removing licenses from $User..."
    $UserLicenses = Get-MgUserLicenseDetail -UserId $User | Select-Object -ExpandProperty SkuId
    Set-MgUserLicense -UserId $User -AddLicenses @() -RemoveLicenses $UserLicenses
  }
  Write-Host "Operation completed."
}

function Search-vtsAllLogs {
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

function Get-vtsNICThroughput {
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

function Get-vts365MailboxStatistics {
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
  param(
    $EmailAddress
  )

  # Connect to Exchange Online
  Write-Host "Connecting to Exchange Online..."
  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }

  if ($null -eq $EmailAddress) { $EmailAddress = $(get-mailbox | Select-Object -expand UserPrincipalName) }

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

function Start-vtsScreenRecording {
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
  Set-ExecutionPolicy Unrestricted
  if ((Test-Path "C:\Windows\Temp\VTS\rc\start.ps1")) {
    Remove-item "C:\Windows\Temp\VTS\rc\start.ps1" -Force -Confirm:$false
  }
  if (-not (Test-Path "C:\Windows\Temp\VTS\rc")) {
    mkdir "C:\Windows\Temp\VTS\rc"
  }

  # Set ACL and NTFS permissions for everyone to have full control
  $acl = Get-Acl "C:\Windows\Temp"
  $permission = "Everyone", "FullControl", "Allow"
  $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
  $acl.SetAccessRule($accessRule)
  
  # Apply to the parent folder
  Set-Acl "C:\Windows\Temp" $acl
  
  # Apply to all child items
  Get-ChildItem "C:\Windows\Temp" -Recurse | ForEach-Object {
    Set-Acl -Path $_.FullName -AclObject $acl
  }

  if ((Get-ScheduledTask -TaskName "RecordSession")) { Unregister-ScheduledTask -TaskName "RecordSession" -Confirm:$false }
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

function Stop-vtsScreenRecording {
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
  Disable-ScheduledTask -TaskName "RecordSession"
  Get-Process ffmpeg, powershell | Stop-Process -Force -Confirm:$false
}

function Get-vtsDirectorySize {
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
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
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
    }
    else {
      Write-Output ("The size of $Path is " + $sizeInMB + " MB")
    }
  }
  catch {
    Write-Error "An error occurred while calculating the size of the directory: $_"
  }
}

function Add-vtsSharePointAdminToAllSites {
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
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the SharePoint admin site URL in the format 'https://yourdomain-admin.sharepoint.com'")]
    [string]$adminSiteUrl,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the username of the user to be made an admin in the format 'user@yourdomain.com'")]
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

function Get-vtsSessionDisconnectTime {
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

function Get-vtsDeviceChange {
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

function Start-vtsPacketCapture {
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

function Stop-vtsPacketCapture {
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
  Write-Host "Stopping packet capture..."
  Get-Process tshark | Stop-Process -Confirm:$false
  Write-Host "Packet capture stopped."
}

function Set-vts365MailboxArchive {
  <#
  .SYNOPSIS
  This function sets up mailbox archiving for a specified user in Office 365.
  
  .DESCRIPTION
  The Set-vts365MailboxArchive function connects to Exchange Online and performs several operations related to mailbox archiving. It allows you to view user retention policies, view user archive details, enable archive, and setup auto-archiving.
  
  .PARAMETER UserEmail
  The email address of the user for whom the mailbox archiving will be set up. The email address should be in 'user@domain.com' format.
  
  .EXAMPLE
  Set-vts365MailboxArchive -UserEmail "user@domain.com"
  
  This example sets up mailbox archiving for the user with the email address "user@domain.com".
  
  .LINK
  M365
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Please enter the user in 'user@domain.com' format")]
    $script:UserEmail
  )

  # Attempt to connect to Exchange Online
  Write-Host "Attempting to connect to Exchange Online..."
  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }

  # Function to view user retention policies
  function ViewUserRetentionPolicy {
    Write-Host "Preparing to view user retention policies..."
    $ViewUserRetentionPolicy = Read-Host "View user retention policies? (y/N)"
    switch ($ViewUserRetentionPolicy) {
      "y" { 
        try {
          get-mailbox -ResultSize Unlimited | Select-Object displayname, UserPrincipalName, RetentionPolicy | Out-Host
          Write-Host "User retention policies viewed successfully."
        }
        catch {
          Write-Error "Failed to view user retention policies: $_"
        }
      }
      Default {}
    }
  }

  # Function to view user archive details
  function ViewUserArchiveDetails {
    Write-Host "Preparing to view user archive details..."
    $ViewUserArchiveDetails = Read-Host "View user archive details? (y/N)"
    switch ($ViewUserArchiveDetails) {
      "y" { 
        try {
          Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, ArchiveStatus, ArchiveDatabase, ArchiveDomain, ArchiveGuid, ArchiveName, ArchiveQuota, ArchiveRelease, ArchiveState, ArchiveStatusDescription, ArchiveWarningQuota | Out-Host
          Write-Host "User archive details viewed successfully."
        }
        catch {
          Write-Error "Failed to view user archive details: $_"
        }
      }
      Default {}
    }
  }

  # Function to enable archive
  function EnableArchive {
    Write-Host "Preparing to enable archive..."
    $EnableUserArchive = Read-Host "Enable Archive? (y/N)"
    switch ($EnableUserArchive) {
      "y" { 
        foreach ($user in $UserEmail) {
          try {
            # Set the archive settings for the new user
            Set-Mailbox $user -ArchiveName "Archive"
            Enable-Mailbox $user -Archive
            Write-Host "Archive enabled successfully for $user."
          }
          catch {
            Write-Error "Failed to enable archive for $user`: $_"
          }
        }
      }
      Default {}
    }
  }

  # Function to enable auto archive
  function EnableAutoArchive {
    Write-Host "Preparing to setup auto-archiving..."
    $SetupAutoArchiving = Read-Host "Setup auto-archiving? (y/N)"

    switch ($SetupAutoArchiving) {
      "y" { 
        $RetentionLimit = Read-Host "Move to archive after this many days"

        # Check if the retention tag exists
        Write-Host "Checking if retention tag exists..."
        try {
          $RetentionTag = Get-RetentionPolicyTag -Identity "Archive after $RetentionLimit Days" -ErrorAction SilentlyContinue
          if ($null -eq $RetentionTag) {
            # Create a retention tag
            New-RetentionPolicyTag -Name "Archive after $RetentionLimit Days" -Type All -AgeLimitForRetention $RetentionLimit -RetentionAction MoveToArchive
            Write-Host "Retention tag created successfully."
          }
        }
        catch {
          Write-Error "Failed to create retention tag: $_"
        }

        # Check if the retention policy exists
        Write-Host "Checking if retention policy exists..."
        try {
          $RetentionPolicy = Get-RetentionPolicy -Identity "$RetentionLimit Day Retention Policy" -ErrorAction SilentlyContinue
          if ($null -eq $RetentionPolicy) {
            # Create a retention policy
            New-RetentionPolicy -Name "$RetentionLimit Day Retention Policy" -RetentionPolicyTagLinks "Archive after $RetentionLimit Days"
            Write-Host "Retention policy created successfully."
          }
        }
        catch {
          Write-Error "Failed to create retention policy: $_"
        }

        foreach ($user in $UserEmail) {
          try {
            # Apply the retention policy to a mailbox
            Set-Mailbox -Identity $user -RetentionPolicy "$RetentionLimit Day Retention Policy"
            Write-Host "Retention policy applied successfully to $user."
          }
          catch {
            Write-Error "Failed to apply retention policy to $user`: $_"
          }
        }
      }
      Default {}
    }
  }

  # Call functions
  ViewUserRetentionPolicy
  ViewUserArchiveDetails

  EnableArchive
  EnableAutoArchive

  ViewUserRetentionPolicy
  ViewUserArchiveDetails
}

function Get-vtsLockoutSource {
  <#
  .SYNOPSIS
  This function retrieves and returns the source of a user's lockout event from the Security log. If no lockout events are detected, it offers to enable logging for account lockouts.
  
  .DESCRIPTION
  The Get-vtsLockoutSource function employs the Get-WinEvent cmdlet to search for a lockout event for a specified user in the Security log. 
  If a lockout event is identified, the function generates a custom object that includes the time of the lockout event, the locked user's username, and the source of the lockout. 
  In the absence of a lockout event, the function prompts the user to enable logging for account lockouts. If the user agrees, the function enables logging for account lockouts and informs the user to wait for another account lockout to retry.
  
  .PARAMETER user
  This optional parameter specifies the username for which the lockout source is to be retrieved. If not provided, the function will return all lockout events.
  
  .EXAMPLE
  PS C:\> Get-vtsLockoutSource -user "jdoe"
  
  This command initiates the retrieval of the source of the lockout event for the user "jdoe". If no lockout events are detected for "jdoe", it will offer to enable logging for account lockouts.
  
  .INPUTS
  System.String
  
  .OUTPUTS
  PSCustomObject
  
  .LINK
  Log Management
  #>
  param(
    [Parameter(Mandatory = $false)]
    [string]$user
  )

  $logs = Get-WinEvent -FilterHashtable @{LogName = 'Security'; Id = 4740 } 2>$null

  if ($null -eq $logs) {
    Write-Host "$($env:COMPUTERNAME) - No lockout events have been detected in the logs. It's possible that logging for account lockouts is currently disabled. Would you like to activate this feature now? (y/n)"
    $EnableLogs = Read-Host
    if ($EnableLogs -eq "y") {
      Auditpol /set /category:"Account Logon" /success:enable /failure:enable | Out-Null
      Auditpol /set /category:"Logon/Logoff" /success:enable /failure:enable | Out-Null
      Auditpol /set /category:"Account Management" /success:enable /failure:enable | Out-Null
      if ($?) {
        Write-Host "Logging has been successfully enabled. Please wait for the occurrence of another account lockout to retry." -ForegroundColor Yellow
      }
    }
    return
  }
  # Return all lockout events in a formatted PowerShell table
  $logs | Where-Object Message -like "*$user*" | Select-Object TimeCreated, @{n = 'LockedUser'; e = { $_.Properties[0].Value } }, @{n = 'LockoutSource'; e = { $_.Properties[1].Value } }, @{n = 'LogSource'; e = { $_.Properties[4].Value } } | Format-Table -AutoSize
}

function Set-vts365CalendarPermissions {
  <#
  .SYNOPSIS
  This function sets the calendar permissions for a specified user in Office 365.
  
  .DESCRIPTION
  The Set-vts365CalendarPermissions function connects to Exchange Online and modifies the calendar permissions for a specified user. It first identifies the calendar path and backs up the existing permissions. Then, it adds or modifies the permissions for the specified access user. Finally, it verifies the updated permissions.
  
  .PARAMETER user
  The email address of the user whose calendar permissions you want to modify. The format should be 'user@domain.com'.
  
  .PARAMETER accessUser
  The email address of the user to whom you want to grant or modify access to the calendar. The format should be 'user@domain.com'.
  
  .PARAMETER accessRights
  The level of access rights you want to grant to the access user. Acceptable values are 'Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author', 'NonEditingAuthor', 'Reviewer', 'Contributor'.
  
  .EXAMPLE
  Set-vts365CalendarPermissions -user 'user1@domain.com' -accessUser 'user2@domain.com' -accessRights 'Editor'
  
  This example modifies the calendar permissions for user1@domain.com, granting 'Editor' access to user2@domain.com.
  
  .NOTES
  The function will backup the existing permissions to C:\temp\CalendarPermissionsBackup.txt before making any changes.
  
  .LINK
  M365
  #>
  param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the email address of the user whose calendar permissions you want to modify. The format should be 'user@domain.com'")]
    $user,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the email address of the user to whom you want to grant or modify access to the calendar. The format should be 'user@domain.com'")]
    $accessUser,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the level of access rights you want to grant to the access user. Acceptable values are 'Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author', 'NonEditingAuthor', 'Reviewer', 'Contributor'")]
    [ValidateSet('Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author', 'NonEditingAuthor', 'Reviewer', 'Contributor')]
    [string]$accessRights
  )

  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }

  # Identify the Calendar Path
  $calendarPath = $user + ":\Calendar"
  Write-Host "Identified Calendar Path: $calendarPath"

  if (-not(Test-Path -Path C:\temp)) {
    New-Item -ItemType Directory -Path C:\temp
  }

  # Backup Calendar Permissions
  Write-Host "Backing up Calendar Permissions to C:\temp\CalendarPermissionsBackup.txt"
  "Original Permissions" | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
  Get-Date | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
  $Permissions = Get-MailboxFolderPermission $calendarPath | Format-Table -AutoSize
  $Permissions | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
    
  # View Existing Calendar Permissions
  Write-Host "Viewing Existing Calendar Permissions..."
  $Permissions
    
  foreach ($member in $accessUser) {
    # Add New Permissions
    Write-Host "Adding New Permissions for $member with access rights $accessRights..."
    try {
      Add-MailboxFolderPermission $calendarPath -User $member -AccessRights $accessRights -ErrorAction Stop
      Write-Host "New permissions for $member with access rights $accessRights added successfully."
    }
    catch {
      Write-Host "Error adding permissions for $member with access rights $accessRights`: $_"
      # Modify Existing Permissions (Optional)
      Write-Host "Attempting to Modify Existing Permissions for $member with access rights $accessRights..."
      try {
        Set-MailboxFolderPermission $calendarPath -User $member -AccessRights $accessRights -ErrorAction Stop
        Write-Host "Permissions for $member with access rights $accessRights modified successfully."
      }
      catch {
        Write-Host "Error setting permissions for $member with access rights $accessRights`: $_"
      }
    }
  }
    
  # Backup Calendar Permissions
  Write-Host "Backing up Calendar Permissions to C:\temp\CalendarPermissionsBackup.txt"
  "Modified Permissions" | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
  Get-Date | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
  $Permissions = Get-MailboxFolderPermission $calendarPath | Format-Table -AutoSize
  $Permissions | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append

  # Verify the Updated Permissions
  Write-Host "Verifying the Updated Permissions for $calendarPath"
  $Permissions
}

function Get-vtsAutotaskTicketDetails {
  <#
  .SYNOPSIS
  This function retrieves the details of a specific Autotask ticket.
  
  .DESCRIPTION
  The Get-vtsAutotaskTicketDetails function makes a REST API call to Autotask's REST API to retrieve the details of a specific ticket. The function requires the ticket number, API integration code, username, and secret as parameters.
  
  .PARAMETER TicketNumber
  The number of the ticket for which details are to be retrieved.
  
  .PARAMETER ApiIntegrationCode
  The API integration code required for the API call.
  
  .PARAMETER UserName
  The username required for the API call.
  
  .PARAMETER Secret
  The secret required for the API call.
  
  .EXAMPLE
  Get-vtsAutotaskTicketDetails -TicketNumber "T20240214.0040" -ApiIntegrationCode '3SzZgr4XTFKavqT59YscgQA7!gr4XTFKI5*' -UserName 'avqT59YscgQA7!@EXAMPLE.COM' -Secret 'ZWqtUYKzPoJv0!'
  
  .LINK
  AutoTask API
  #>
  param (
    [Parameter(Mandatory = $true)]
    [string]$TicketNumber,
    [Parameter(Mandatory = $true)]
    [string]$ApiIntegrationCode,
    [Parameter(Mandatory = $true)]
    [string]$UserName,
    [Parameter(Mandatory = $true)]
    [string]$Secret
  )

  # Define the base URI for Autotask's REST API
  $baseUri = "https://webservices14.autotask.net/ATServicesRest/V1.0"

  # Set the necessary headers for the API call
  $headers = @{
    "ApiIntegrationCode" = $ApiIntegrationCode
    "UserName"           = $UserName
    "Secret"             = $Secret
  }

  # Define the endpoint for retrieving ticket details
  $endpoint = "/Tickets/query"

  # Define the body for the API call
  $body = @{
    "Filter" = @(
      @{
        "field" = "ticketNumber"
        "op"    = "eq"
        "value" = $TicketNumber
      }
    )
  } | ConvertTo-Json

  # Make the API call to Autotask using the POST method
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Post' -Headers $headers -Body $body -ContentType "application/json"

  # Output the response
  $response | Select-Object -expand items
}

function Get-vtsAutotaskTicketNotes {
  <#
  .SYNOPSIS
  This function retrieves all notes of a specific Autotask ticket using the ticket number.
  
  .DESCRIPTION
  The Get-vtsAutotaskTicketNotes function makes a REST API call to Autotask's REST API to retrieve the ticket ID of a specific ticket using the ticket number, and then retrieves all notes of that ticket. The function requires the ticket number, API integration code, username, and secret as parameters.
  
  .PARAMETER TicketNumber
  The number of the ticket for which notes are to be retrieved.
  
  .PARAMETER ApiIntegrationCode
  The API integration code required for the API call.
  
  .PARAMETER UserName
  The username required for the API call.
  
  .PARAMETER Secret
  The secret required for the API call.
  
  .EXAMPLE
  Get-vtsAutotaskTicketNotes -TicketNumber "T20240214.0040" -ApiIntegrationCode 'tOi5y2bp7j8U7=T59YscgQA7!gr4XTFKI5*' -UserName 'avqTtOi5y2bp7j8U7=gQA7!@EXAMPLE.COM' -Secret 'ZWqtUYKztOi5y2bp7j8U7=oJv0!'
  
  .LINK
  AutoTask API
  #>
  param (
    [Parameter(Mandatory = $true)]
    [string]$TicketNumber,
    [Parameter(Mandatory = $true)]
    [string]$ApiIntegrationCode,
    [Parameter(Mandatory = $true)]
    [string]$UserName,
    [Parameter(Mandatory = $true)]
    [string]$Secret
  )

  # Define the base URI for Autotask's REST API
  $baseUri = "https://webservices14.autotask.net/ATServicesRest/V1.0"

  # Set the necessary headers for the API call
  $headers = @{
    "ApiIntegrationCode" = $ApiIntegrationCode
    "UserName"           = $UserName
    "Secret"             = $Secret
  }

  # Define the endpoint for retrieving ticket details
  $endpoint = "/Tickets/query"

  # Define the body for the API call
  $body = @{
    "Filter" = @(
      @{
        "field" = "TicketNumber"
        "op"    = "eq"
        "value" = $TicketNumber
      }
    )
  } | ConvertTo-Json

  # Make the API call to Autotask using the POST method to get the ticket ID
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Post' -Headers $headers -Body $body -ContentType "application/json"

  # Get the ID of the ticket
  $TicketId = $response.items.id

  # Define the endpoint for retrieving ticket notes
  $endpoint = "/Tickets/$TicketId/notes"

  # Make the API call to Autotask using the GET method to get the ticket notes
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Get' -Headers $headers -ContentType "application/json"

  # Output the response
  $response | Select-Object -expand items
}

function Get-vtsAutoTaskContactNamebyID {
  <#
  .SYNOPSIS
  This function retrieves the first and last name of a contact from Autotask's REST API using the contact's ID.
  
  .DESCRIPTION
  The Get-vtsAutoTaskContactNamebyID function makes a GET request to Autotask's REST API to retrieve the details of a contact. It requires the contact's ID, API integration code, username, and secret as parameters. The function returns the first and last name of the contact.
  
  .PARAMETER ContactID
  The ID of the contact whose details are to be retrieved.
  
  .PARAMETER ApiIntegrationCode
  The API integration code required for the API call.
  
  .PARAMETER UserName
  The username required for the API call.
  
  .PARAMETER Secret
  The secret required for the API call.
  
  .EXAMPLE
  Get-vtsAutoTaskContactNamebyID -ContactID "30683060" -ApiIntegrationCode "YourApiIntegrationCode" -UserName "YourUserName" -Secret "YourSecret"
  This example shows how to call the function with all required parameters.
  
  .LINK
  AutoTask API
  #>
  param(
    [Parameter(Mandatory = $true)]
    [string]$ContactID,
    [Parameter(Mandatory = $true)]
    [string]$ApiIntegrationCode,
    [Parameter(Mandatory = $true)]
    [string]$UserName,
    [Parameter(Mandatory = $true)]
    [string]$Secret
  )

  # Define the base URI for Autotask's REST API
  $baseUri = "https://webservices14.autotask.net/ATServicesRest/V1.0"
        
  # Set the necessary headers for the API call
  $headers = @{
    "ApiIntegrationCode" = $ApiIntegrationCode
    "UserName"           = $UserName
    "Secret"             = $Secret
  }
        
  # Define the endpoint for contact
  $endpoint = "/Contacts/$ContactID"
        
  # Make the API call to Autotask using the GET method to get the contact
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Get' -Headers $headers -ContentType "application/json"
        
  # Return the first and last name of the contact
  "$($response.item.firstName) $($response.item.lastName)"
}

function Get-vtsAutoTaskResourceNamebyID {
  <#
  .SYNOPSIS
  This function retrieves the first and last name of a resource from Autotask's REST API using the resource's ID.
  
  .DESCRIPTION
  The Get-vtsAutoTaskResourceNamebyID function makes a GET request to Autotask's REST API to retrieve the details of a resource. It requires the resource's ID, API integration code, username, and secret as parameters. The function returns the first and last name of the resource.
  
  .PARAMETER ResourceID
  The ID of the resource whose details are to be retrieved.
  
  .PARAMETER ApiIntegrationCode
  The API integration code required for the API call.
  
  .PARAMETER UserName
  The username required for the API call.
  
  .PARAMETER Secret
  The secret required for the API call.
  
  .EXAMPLE
  Get-vtsAutoTaskResourceNamebyID -ResourceID "30683060" -ApiIntegrationCode "YourApiIntegrationCode" -UserName "YourUserName" -Secret "YourSecret"
  This example shows how to call the function with all required parameters.
  
  .LINK
  AutoTask API
  #>
  param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceID,
    [Parameter(Mandatory = $true)]
    [string]$ApiIntegrationCode,
    [Parameter(Mandatory = $true)]
    [string]$UserName,
    [Parameter(Mandatory = $true)]
    [string]$Secret
  )

  # Define the base URI for Autotask's REST API
  $baseUri = "https://webservices14.autotask.net/ATServicesRest/V1.0"
        
  # Set the necessary headers for the API call
  $headers = @{
    "ApiIntegrationCode" = $ApiIntegrationCode
    "UserName"           = $UserName
    "Secret"             = $Secret
  }
        
  # Define the endpoint for resource
  $endpoint = "/Resources/$ResourceID"
        
  # Make the API call to Autotask using the GET method to get the resource
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Get' -Headers $headers -ContentType "application/json"
        
  # Return the first and last name of the resource
  "$($response.item.firstName) $($response.item.lastName)"
}

function Get-vtsFileContentMatch {
  <#
  .SYNOPSIS
  This function searches for a specific pattern in the content of files in a given directory and optionally exports the results to a CSV file.
  
  .DESCRIPTION
  The Get-vtsFileContentMatch function takes a directory path, a pattern to match, and an optional array of file types to exclude. It recursively searches through all files in the specified directory (excluding the specified file types), and returns a custom object for each line in each file that matches the specified pattern. The custom object contains the full path of the file and the line that matched the pattern. The function can also export the results to a CSV file if the path to the CSV file is provided.
  
  .PARAMETER Path
  The path to the directory to search. This parameter is mandatory.
  
  .PARAMETER Pattern
  The pattern or word to match in the file content. This parameter is mandatory.
  
  .PARAMETER Exclude
  An array of file types to exclude from the search. The default value is '*.exe', '*.dll'. This parameter is optional.
  
  .PARAMETER ExportToCsv
  A boolean value indicating whether to export the results to a CSV file. This parameter is optional.
  
  .PARAMETER CsvPath
  The path to the CSV file to export the results to. This parameter is optional.
  
  .EXAMPLE
  Get-vtsFileContentMatch -Path "C:\Users\Username\Documents" -Pattern "error" -ExportToCsv $true -CsvPath "C:\Users\Username\Documents\results.csv"
  
  This example searches for the word "error" in all files in the "C:\Users\Username\Documents" directory, excluding .exe and .dll files, and exports the results to a CSV file.
  
  .LINK
  Utilities
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Please provide the path to the directory.")]
    [string]$Path,
    [Parameter(Mandatory = $true, HelpMessage = "Please provide the pattern or word to match.")]
    [string]$Pattern,
    [Parameter(Mandatory = $false, HelpMessage = "Please provide the file types to exclude. Default is '*.exe', '*.dll'.")]
    [string[]]$Exclude = @("*.exe", "*.dll"),
    [Parameter(Mandatory = $false, HelpMessage = "Please indicate if you want to export the results to a CSV file.")]
    [bool]$ExportToCsv = $false,
    [Parameter(Mandatory = $false, HelpMessage = "Please provide the path to the CSV file.")]
    [string]$CsvPath = "C:\temp\$(Get-Date -f yyyy-MM-dd-HH-mm)-$($env:COMPUTERNAME)-FilesIncluding-$Pattern.csv"
  )

  $Results = @()
  Get-ChildItem -Path $Path -Recurse -File -Exclude $Exclude | ForEach-Object {
    $filePath = $_.FullName
    Get-Content -Path $filePath | ForEach-Object {
      if ($_ -match $Pattern) {
        $Result = [pscustomobject]@{
          FilePath = $filePath
          Match    = $_
        }
        $Result | Format-List
        $Results += $Result
      }
    }
  }

  if ($ExportToCsv) {
    $Results | Export-Csv -Path $CsvPath -NoTypeInformation -Force
    if ($?) { Write-Host "Results exported to $CsvPath" -ForegroundColor Yellow }
  }
}

function Revoke-vts365EmailMessage {
  <#
  .SYNOPSIS
  This function revokes a specific email message in Office 365.
  
  .DESCRIPTION
  The Revoke-vts365EmailMessage function connects to Exchange Online and IPPS Session, gets the message trace, starts a compliance search, waits for the search to complete, purges and deletes the email instances, and checks the status of the search action.
  
  .PARAMETER TicketNumber
  This is a mandatory parameter. It is the unique identifier for the compliance search.
  
  .PARAMETER From
  This is an optional parameter. It is the sender's email address.
  
  .PARAMETER To
  This is an optional parameter. It is the recipient's email address.
  
  .EXAMPLE
  Revoke-vts365EmailMessage -TicketNumber "12345" -From "sender@example.com" -To "recipient@example.com"
  
  This example shows how to revoke an email message sent from "sender@example.com" to "recipient@example.com" with the ticket number "12345".
  
  .LINK
  Still in Development
  
  #>
  param(
    # [Parameter(Mandatory = $true)]
    [string]$TicketNumber = "TEST4",
    [array]$From,
    [array]$To = "Zach.Koscoe@completehealth.com"
  )
    
  Write-Host "Connecting to Exchange Online and IPPS Session..."
  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }
  # Connect-IPPSSession

  Write-Host "Getting message trace..."
  if (($From -ne $null) -and ($To -ne $null)) {
    $Message = Get-MessageTrace -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date) -RecipientAddress $To -SenderAddress $From | Out-GridView -OutputMode Multiple
  }
    
  if (($To -ne $null) -and ($From -eq $null)) {
    $Message = Get-MessageTrace -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date) -RecipientAddress $To | Out-GridView -OutputMode Multiple
  }
    
  if (($To -eq $null) -and ($From -ne $null)) {
    $Message = Get-MessageTrace -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date) -SenderAddress $From | Out-GridView -OutputMode Multiple
  }
    
  Write-Host "Starting compliance search..."
  New-ComplianceSearch -Name $TicketNumber -ExchangeLocation $($Message.RecipientAddress) -ContentMatchQuery "(c:c)(from=$($Message.SenderAddress))(subjecttitle=""$($Message.subject)"")" | Start-ComplianceSearch
    
  Write-Host "Waiting for the search to complete..."
  do {
    Start-Sleep -Seconds 60
    $searchStatus = (Get-ComplianceSearch -Identity $TicketNumber).Status
  } while ($searchStatus -ne 'Completed')
    
  Write-Host "Purging and deleting the email instances..."
  New-ComplianceSearchAction -SearchName $TicketNumber -Purge -PurgeType HardDelete
    
  Write-Host "Checking the status of the search action..."
  do {
    Start-Sleep -Seconds 60
    $actionStatus = Get-ComplianceSearchAction | Where-Object Name -eq "$($TicketNumber)_Purge" | Select-Object -expand Status
  } while ($actionStatus -ne 'Completed')

  Write-Host "Operation completed."
}

function Start-vtsRepair {
  <#
  .SYNOPSIS
  This function invokes a system check using DISM and System File Checker.
  
  .DESCRIPTION
  The Start-vtsRepair function initiates a system check by first running the DISM restore health process, followed by a System File Checker scan. It provides status updates at each stage of the process.
  
  .EXAMPLE
  Start-vtsRepair
  This command will initiate the system check process.
  
  .NOTES
  The DISM restore health process can help fix Windows corruption errors. The System File Checker scan will scan all protected system files, and replace corrupted files with a cached copy.
  
  .LINK
  Utilities
  #>
  Write-Host "Starting DISM restore health process..."
  dism /online /cleanup-image /restorehealth
  Write-Host "DISM restore health process completed."

  Write-Host "Starting System File Checker scan now..."
  sfc /scannow
  Write-Host "System File Checker scan completed."
}

function New-vtsSPOnlineDocumentLibrary {
  <#
  .SYNOPSIS
  This function creates a new SharePoint Online document library.
  
  .DESCRIPTION
  The New-vtsSPOnlineDocumentLibrary function creates a new document library in a specified SharePoint Online site. 
  It requires the organization name, site URL, and the name of the new document library as parameters. 
  If the PnP.PowerShell module is not installed, the function will install it for the current user.
  
  .PARAMETER orgName
  The name of the organization. This parameter is mandatory. Example: contoso
  
  .PARAMETER siteUrl
  The URL of the SharePoint Online site where the new document library will be created. This parameter is mandatory. Example: https://contoso.sharepoint.com/sites/test
  
  .PARAMETER libraryName
  The name of the new document library to be created. This parameter is mandatory.
  
  .EXAMPLE
  New-vtsSPOnlineDocumentLibrary -orgName "contoso" -siteUrl "https://contoso.sharepoint.com/sites/test" -libraryName "NewLibrary"
  This example creates a new document library named "NewLibrary" in the "test" site of the "contoso" organization.
  
  .LINK
  SharePoint Online
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the organization name. Example: contoso")]
    [string]$orgName,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the site name. Example: https://contoso.sharepoint.com/sites/test")]
    [string]$siteUrl,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the name of the new document library.")]
    [string]$libraryName
  )

  $PnPConnected = Get-PnPConnection

  if (-not($PnPConnected)){
    if (-not(Get-Module PnP.PowerShell -ListAvailable)) {
      Install-Module -Name PnP.PowerShell -Scope CurrentUser
    }
  
    Import-Module PnP.PowerShell
  
    # Connect to SharePoint Online
    # Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Connect-PnPOnline -Url $siteUrl -Interactive
  }

  # Create a new document library
  New-PnPList -Title $libraryName -Template DocumentLibrary -Url $libraryName
}

function Get-vtsSPOnlineDocumentLibraryFolders {
  <#
  .SYNOPSIS
  This function retrieves all folders in a specified SharePoint Online document library.
  
  .DESCRIPTION
  The Get-vtsSPOnlineDocumentLibraryFolders function retrieves all folders in a specified SharePoint Online document library. 
  It requires the site URL and the name of the document library as parameters. 
  If the PnP.PowerShell module is not installed, the function will install it for the current user.
  
  .PARAMETER siteUrl
  The URL of the SharePoint Online site where the document library is located. This parameter is mandatory. Example: https://contoso.sharepoint.com/sites/test
  
  .PARAMETER libraryName
  The name of the document library from which to retrieve folders. This parameter is mandatory.
  
  .EXAMPLE
  Get-vtsSPOnlineDocumentLibraryFolders -siteUrl "https://contoso.sharepoint.com/sites/test" -libraryName "LibraryName"
  This example retrieves all folders in the "LibraryName" document library on the "test" site.
  
  .LINK
  SharePoint Online
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the site URL. Example: https://contoso.sharepoint.com/sites/test")]
    [string]$siteUrl,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the document library name.")]
    [string]$libraryName,
    [switch]$Recursive
  )

  $PnPConnected = Get-PnPConnection

  if (-not($PnPConnected)){
    if (-not(Get-Module PnP.PowerShell -ListAvailable)) {
      Install-Module -Name PnP.PowerShell -Scope CurrentUser
    }
  
    Import-Module PnP.PowerShell
  
    # Connect to SharePoint Online
    # Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Connect-PnPOnline -Url $siteUrl -Interactive
  }

  # Get all folders in the document library
  if ($Recursive){
    $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $libraryName -ItemType Folder -Recursive
  } else {
    $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $libraryName -ItemType Folder
  }
  # Return the folder names
  return $folders.Name
}

function Get-vtsDomainLockouts {
  <#
  .SYNOPSIS
  This function retrieves all domain controllers and invokes the Get-vtsLockoutSource function on each of them.
  
  .DESCRIPTION
  The Get-vtsDomainLockouts function retrieves a list of all domain controllers in the current Active Directory domain. It then invokes the Get-vtsLockoutSource function on each domain controller to retrieve the source of account lockout events.
  
  .PARAMETER None
  This function does not accept any parameters.
  
  .EXAMPLE
  PS C:\> Get-vtsDomainLockouts
  
  This command retrieves all domain controllers and invokes the Get-vtsLockoutSource function on each of them.
  
  .INPUTS
  None
  
  .OUTPUTS
  None
  
  .LINK
  Log Management
  #>
  Write-Host "Retrieving domain controllers...`n"
  $DCs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty Hostname
  Write-Host "Domain Controllers:`n $($DCs -join ""`n "")`n"
  Write-Host "Invoking Get-vtsLockoutSource on each domain controller...`n"
  $DCs | ForEach-Object { Invoke-Command -ComputerName $_ -ScriptBlock { irm rrwo.us | iex *>$null ; Write-Host "$($env:COMPUTERNAME) Results:" -ForegroundColor Yellow ; Get-vtsLockoutSource } }
}


function Suspend-vtsADUser {
  <#
  .SYNOPSIS
  This function suspends an Active Directory user and optionally schedules a task to reactivate the user after a specified number of days.
  
  .DESCRIPTION
  The Suspend-vtsADUser function fetches all Active Directory users and displays them in an Out-GridView for selection. After a user is selected, the function prompts for the number of days after which to reactivate the user. If no input is entered, the user is disabled. If a number of days is entered, the function disables the user and schedules a task to re-enable the user after the specified number of days.
  
  .PARAMETER None
  This function does not take any parameters.
  
  .EXAMPLE
  Suspend-vtsADUser
  
  This example shows how to run the function without any parameters. It will fetch all Active Directory users, allow you to select a user, and then prompt for the number of days after which to reactivate the user.
  
  .NOTES
  This function requires the Active Directory module for Windows PowerShell and the Task Scheduler cmdlets.
  
  .LINK
  Active Directory
  #>

  $users = Get-ADUser -Filter * -Property DisplayName, PhysicalDeliveryOfficeName, Manager | Select-Object DisplayName, PhysicalDeliveryOfficeName, @{Name = 'Manager'; Expression = { (Get-ADUser $_.Manager).Name } }, SamAccountName | Sort-Object DisplayName
    
  $selectedUser = $users | Out-GridView -Title "Select a User to Manage" -PassThru
    
  if ($null -ne $selectedUser) {
    $days = Read-Host "Enter the number of days after which to reactivate the user (Leave empty to disable the user)"
        
    $verificationUser = Read-Host "Please type $($selectedUser.SamAccountName) to confirm suspension."

    if ($verificationUser -ne $selectedUser.SamAccountName) {
      Write-Host "Verification failed. Exiting..." -ForegroundColor Red
      break
    }

    if ([string]::IsNullOrWhiteSpace($days)) {
      Disable-ADAccount -Identity $selectedUser.SamAccountName
      Write-Host "User $($selectedUser.SamAccountName) has been disabled."
    }
    else {
      try {
        $days = [int]$days
        Disable-ADAccount -Identity $selectedUser.SamAccountName
                
        $action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-NoProfile -WindowStyle Hidden -Command `"Enable-ADAccount -Identity $($selectedUser.SamAccountName)`""
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddDays($days).Date
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
            
        $taskName = "EnableADUser_" + $selectedUser.SamAccountName
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Re-enable AD user $($selectedUser.SamAccountName) after $days days. Task created $(Get-Date) by $($env:USERNAME)" | Out-Null
                
        if ($?) {
          Write-Host "`nUser $($selectedUser.SamAccountName) has been disabled. A task has been scheduled to re-enable the account after $days days at:"
          Write-Host "`n$((Get-Date).AddDays($days).Date)`n" -ForegroundColor Yellow
        }
      }
      catch {
        Write-Host "An error occurred: $_"
      }
    }
        
  }
  else {
    Write-Host "No user was selected."
  }

}

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

function Compare-vtsFiles {
  <#
  .SYNOPSIS
  This script compares the files in two directories and generates a report of their SHA1 hashes.
  
  .DESCRIPTION
  The Compare-vtsFiles function takes in two parameters, the source folder and the destination folder. It calculates the SHA1 hash of each file in both folders and compares them. If the hashes match, it means the files are identical. If they don't, the files are different. The function generates a report of the comparison results, including the file paths and their hashes. The report is saved in a CSV file in the TEMP directory.
  
  .PARAMETER SourceFolder
  The path of the source folder.
  
  .PARAMETER DestinationFolder
  The path of the destination folder.
  
  .PARAMETER ReportPath
  The path where the report will be saved. By default, it is saved in the TEMP directory.
  
  .EXAMPLE
  Compare-vtsFiles -SourceFolder "C:\Source" -DestinationFolder "C:\Destination"
  
  This will compare the files in the Source and Destination folders and generate a report in the TEMP directory.
  
  .EXAMPLE
  Compare-vtsFiles -SourceFolder "C:\Source" -DestinationFolder "C:\Destination" -ReportPath "C:\Reports\Hashes.csv"
  
  This will compare the files in the Source and Destination folders and generate a report in the specified path.
  
  .LINK
  File Management
  #>
  param (
    [string]$SourceFolder,
    [string]$DestinationFolder,
    [string]$ReportPath = "$env:TEMP\VTS\$(Get-Date -f yyyy-MM-dd-hhmmss)_Hashes.csv"
  )

  # Initialize a new list to store the results
  $result = [System.Collections.Generic.List[object]]::new()

  # Define a script block to process each file
  $sb = {
    process {
      # Ignore 'Thumbs.db' files
      if ($_.Name -eq 'Thumbs.db') { return }

      # Create a custom object with file properties
      [PSCustomObject]@{
        h  = (Get-FileHash $_.FullName -Algorithm SHA1).Hash # File hash
        n  = $_.Name # File name
        fn = $_.fullname # Full file path
      }
    }
  }

  # Get all files from the source and destination folders
  $sourceFiles = Get-ChildItem $SourceFolder -Recurse -File | & $sb
  $destinationFiles = Get-ChildItem $DestinationFolder -Recurse -File | & $sb

  # Process each file in the source folder
  foreach ($file in $sourceFiles) {
    # If the file exists in the destination folder
    if ($destinationFile = $destinationFiles | Where-Object { $_.n -eq $file.n }) {
      # Create a custom object with source and target file properties
      $comparisonResult = [PSCustomObject]@{
        SourceFilePath = $file.fn
        SourceFileHash = $file.h
        TargetFilePath = $destinationFile.fn
        TargetFileHash = $destinationFile.h
        Status         = if ($file.h -eq $destinationFile.h) { 'Hashes Match' } else { 'Hashes Do Not Match' }
      }
    }
    else {
      # If the file does not exist in the destination folder
      $comparisonResult = [PSCustomObject]@{
        SourceFilePath = $file.fn
        SourceFileHash = $file.h
        TargetFilePath = $null
        TargetFileHash = $null
        Status         = 'File not found in destination'
      }
    }

    # Add the comparison result to the result list
    $result.Add($comparisonResult)
  }

  $ReportDirectory = Split-Path -Path $ReportPath -Parent
  if (!(Test-Path -Path $ReportDirectory)) {
    New-Item -ItemType Directory -Path $ReportDirectory -Force | Out-Null
  }

  # Output the result list in a table format
  $result | Export-Csv -NoTypeInformation -Path $ReportPath -Force

  Write-Host "Report exported to $ReportPath"
}

function Connect-vtsWiFi {
  <#
  .SYNOPSIS
  This script connects to a WiFi network using the provided SSID and password.
  
  .DESCRIPTION
  The Connect-vtsWiFi function connects to a WiFi network using the provided SSID and password. If no SSID is provided, the function will list all available networks and prompt the user to select one. If no password is provided, the function will prompt the user to enter one. If a WiFi profile for the selected SSID already exists, the function will ask the user if they want to remove the old profile and replace it.
  
  .PARAMETER SSID
  The SSID of the WiFi network to connect to. If not provided, the function will list all available networks and prompt the user to select one.
  
  .PARAMETER Password
  The password for the WiFi network to connect to. If not provided, the function will prompt the user to enter one.
  
  .EXAMPLE
  Connect-vtsWiFi -SSID "MyWiFiNetwork" -Password "MyPassword"
  
  This example connects to the WiFi network "MyWiFiNetwork" using the password "MyPassword".
  
  .LINK
  Network
  #>
  param (
    [Parameter(Mandatory = $false)]
    [string]$SSID,
      
    [Parameter(Mandatory = $false)]
    [SecureString]$Password
  )

  Write-Host "`nChecking for wifiprofilemanagement module..."
  if (-not (Get-Module -ListAvailable -Name wifiprofilemanagement)) {
    Write-Host "`nInstalling wifiprofilemanagement module..."
    Install-Module -Name wifiprofilemanagement -Force -Scope CurrentUser
  }

  if (-not $SSID) {
    Write-Host "`nFetching available networks...`n"

    $networks = (netsh wlan show networks | sls ^SSID) -replace "SSID . : " | sort

    $i = 1
    $networkList = @()
    foreach ($network in $networks) {
      Write-Host "$i`: $network"
      $networkList += $network
      $i++
    }
    $selection = Read-Host -Prompt "`nPlease select a network by number"
    $SSID = ($networkList[$selection - 1]).Trim()
  }

  if (-not $Password) {
    $Password = Read-Host -Prompt "Please enter the password for $SSID" -AsSecureString
  }

  Write-Host "`nChecking for existing WiFi profile..."
  if (Get-WiFiProfile -ProfileName $SSID 2>$null) {
    $userChoice = Read-Host -Prompt "WiFi profile for $SSID already exists. Do you want to remove the old profile and replace it? (y/n)"
    if ($userChoice -eq "y") {
      Write-Host "`nRemoving old WiFi profile..."
      Remove-WiFiProfile -ProfileName $SSID
      Write-Host "`nCreating new WiFi profile..."
      New-WiFiProfile -ProfileName $SSID -ConnectionMode auto -Authentication WPA2PSK -Password $Password -Encryption AES
    }
  }
  else {
    Write-Host "`nCreating new WiFi profile..."
    New-WiFiProfile -ProfileName $SSID -ConnectionMode auto -Authentication WPA2PSK -Password $Password -Encryption AES
  }
  Write-Host "`nConnecting to WiFi profile..."
  Connect-WiFiProfile -ProfileName $SSID

}

function ai3 {
  <#
  .SYNOPSIS
  This script uses OpenAI's GPT-4 model to generate IT support ticket notes and follow-up questions.
  
  .DESCRIPTION
  The script consists of three main functions: LineAcrossScreen, Invoke-OpenAIAPI, and Generate-Questions. 
  
  LineAcrossScreen creates a line across the console screen with a specified color. 
  
  Invoke-OpenAIAPI sends a request to OpenAI's API with a given prompt and API key, and optionally a previous response for context. It returns the AI's response.
  
  Generate-Questions uses the OpenAI API to generate a set of follow-up questions based on the provided ticket notes.
  
  The main loop of the script prompts the user to enter an issue description and ticket notes, generates follow-up questions, and finally generates the ticket notes.
  
  .PARAMETER Prompt
  The prompt to be sent to the OpenAI API.
  
  .PARAMETER OpenAIAPIKey
  The API key for OpenAI.
  
  .PARAMETER PreviousResponse
  The previous response from the AI, used for context in the next API call.
  
  .PARAMETER TicketNotes
  The ticket notes to be used as context for generating follow-up questions.
  
  .PARAMETER Color
  The color of the line to be drawn across the console screen.
  
  .EXAMPLE
  PS> ai3
  
  .LINK
  AI
  #>

  Start-Transcript | Out-Null

  if (![string]::IsNullOrEmpty($response)) {
    $continueSession = Read-Host -Prompt "Would you like to continue from where you left off? (y/n)"
    if ($continueSession -eq "n") {
      $context = ""
      $userInput = ""
      $clarifyingQuestions = ""
      $global:response = ""
    }
  }

  $KeyPath = "$env:LOCALAPPDATA\VTS\SecureKeyFile.txt"

  function Encrypt-SecureString {
    param(
      [Parameter(Mandatory = $true)]
      [string]$InputString,
      [Parameter(Mandatory = $true)]
      [string]$FilePath
    )
  
    $secureString = ConvertTo-SecureString -String $InputString -AsPlainText -Force
    $secureString | Export-Clixml -Path $FilePath
  }

  function Decrypt-SecureString {
    param(
      [Parameter(Mandatory = $true)]
      [string]$FilePath
    )

    $secureString = Import-Clixml -Path $FilePath
    $decryptedString = [System.Net.NetworkCredential]::new("", $secureString).Password
    return $decryptedString
  }
    
  if (Test-Path $KeyPath) {
    $OpenAIAPIKey = Decrypt-SecureString -FilePath $KeyPath
  }
    
  if ([string]::IsNullOrEmpty($OpenAIAPIKey)) {
    $OpenAIAPIKey = Read-Host -Prompt "Please enter your OpenAI API Key" -AsSecureString
    $OpenAIAPIKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($OpenAIAPIKey))
    $saveKey = Read-Host -Prompt "Would you like to save the API key for future use? (y/n)"
    if ($saveKey -eq "y") {
      $KeyDirectory = Split-Path -Path $KeyPath -Parent
      # Check if the directory exists, if not, create it
      if (!(Test-Path -Path $KeyDirectory)) {
        New-Item -ItemType Directory -Path $KeyDirectory | Out-Null
      }

      Encrypt-SecureString -InputString $OpenAIAPIKey -FilePath $KeyPath

      Write-Host "Your API key has been saved securely."
    }
  }

  function LineAcrossScreen {
    param (
      [Parameter(Mandatory = $false)]
      [string]$Color = "Green"
    )
    $script:windowWidth = (Get-Host).UI.RawUI.WindowSize.Width
    Write-Host ('-' * $windowWidth) -ForegroundColor $Color
  }

  function Invoke-OpenAIAPI {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Prompt,
      [Parameter(Mandatory = $true)]
      [string]$OpenAIAPIKey,
      [Parameter(Mandatory = $false)]
      [string]$PreviousResponse
    )

    $Headers = @{
      "Content-Type"  = "application/json"
      "Authorization" = "Bearer $OpenAIAPIKey"
    }
    $Body = @{
      "model"             = "gpt-4-0125-preview"
      "messages"          = @(
        @{
          "role"    = "system"
          "content" = "Act as a helpful IT technician to create detailed ticket notes for IT support issues."
        },
        @{
          "role"    = "system"
          "content" = "Always use the plural first person (e.g., 'we') and refer to items impersonally (e.g., 'the computer'). ALWAYS structure responses in the following format:\n\nComputer Name: <detect the computername from notes entered.>\n\nReported Issue:<text describing issue>\n\nTroubleshooting Methods:\n- <bulletted troubleshooting steps here>\n\nResolution:<resolution here>\n\nComments & Misc. info:<miscellaneous info here>\n\nMessage to End User:\n<email to end user here using non-technical, straight-forward, common wording>"
        },
        @{
          "role"    = "system"
          "content" = "Include all troubleshooting steps in the 'Troubleshooting Methods' section. Don't exclude ANY details."
        },
        @{
          "role"    = "system"
          "content" = "Imitate the following writing examples while avoiding using adjectives and uncommon words: Hi Janine,\n\nI understand this has been such a turbulent issue for Dr. Harris. Our ability to support personal non-windows devices is limited as our remote team does not have a way to access this device to provide immediate assistance.\n\nI am working to have an on-site technician deployed to your location to have this resolved as quickly as possible. Please let me know if you have any questions or concerns.\n\nRespectfully,\n\n\nHi Summer,\n\nI have modified the policy to remove gaming, now instead of just running on every computer, it will run for every user. This should catch any stragglers that may have still been out there.\n\nI am going to let the policy deploy across workstations today and through the weekend and check in on Monday to see if any computers still have the Solitaire application.\n\nRespectfully,"
        },
        @{
          "role"    = "user"
          "content" = "Here's an example of the output I want:\n\nComputer Name: SD-PC20\n\nIssue Reported: Screen flickering\n\nTroubleshooting Methods:\n- Checked for Windows Updates.\n- Navigated to the Device Manager, located Display Adapters and right-clicked on the NVIDIA GeForce GTX 1050, selecting Update Driver.\n- Clicked on Search Automatically for Drivers, followed by Search for Updated Drivers on Windows Update.\n- Searched for 'gtx 1050 drivers' and clicked on the first result.\n- Clicked on the Official Drivers link and downloaded the driver.\n- Updated the graphics driver, resolving the issue.\n\nResolution: Updating the graphics driver resolved the issue.\n\nAdditional Comments: None\n\n\nMessage to End User: \n\n[User Name],\n\nWe have successfully resolved the screen flickering issue you were experiencing by updating the graphics driver. At your earliest convenience, please test your system to confirm that the issue with your screen has been rectified. Should you encounter any additional issues or require further assistance, do not hesitate to reach out to us.\n\nRespectfully,"
        },
        @{
          "role"    = "user"
          "content" = "Document only the steps explicitly stated. Ensure accuracy and quality of the ticket notes, and make them sound good."
        },
        @{
          "role"    = "user"
          "content" = "Refer to previous responses for context when updating ticket notes: $PreviousResponse"
        },
        @{
          "role"    = "user"
          "content" = "Update the ticket notes, taking the following into account: $prompt"
        },
        @{
          "role"    = "assistant"
          "content" = ""
        }
      )
      "temperature"       = 0
      'top_p'             = 1.0
      'frequency_penalty' = 0.0
      'presence_penalty'  = 0.0
      'stop'              = @('"""')
    } | ConvertTo-Json
    
    try {
        
      $global:response = (Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Method Post -Body $Body -Headers $Headers).choices.message.content
      Return $global:response
    }
    catch {
      Write-Error "$($_.Exception.Message)"
    }
  }

  function Generate-Questions {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Prompt,
      [Parameter(Mandatory = $true)]
      [string]$OpenAIAPIKey,
      [Parameter(Mandatory = $false)]
      [string]$TicketNotes
    )

    $Headers = @{
      "Content-Type"  = "application/json"
      "Authorization" = "Bearer $OpenAIAPIKey"
    }
    $Body = @{
      "model"             = "gpt-4-0125-preview"
      "messages"          = @( @{
          "role"    = "system"
          "content" = "You are a helpful IT technician assistant that helps generate followup questions, and outputs them one per line."
        },
        @{
          "role"    = "user"
          "content" = "Ticket Notes: $TicketNotes"
        },
        @{
          "role"    = "user"
          "content" = "$prompt"
        },
        @{
          "role"    = "assistant"
          "content" = ""
        })
      "temperature"       = 0
      'top_p'             = 1.0
      'frequency_penalty' = 0.0
      'presence_penalty'  = 0.0
      'stop'              = @('"""')
    } | ConvertTo-Json
    
    try {
        
      $global:response = (Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Method Post -Body $Body -Headers $Headers).choices.message.content
      Return $global:response
    }
    catch {
      Write-Error "$($_.Exception.Message)"
    }
  }

  # Main loop
  $context = "$($global:response)"

  if ( ($null -eq $context) -or ("" -eq $context) ) { $context = Read-Host "`nEnter an issue description`n" ; LineAcrossScreen }
  
  Write-Host "`nEnter ticket note (or 'done' to finish or 'skip' to skip follow up questions)`n`n"
  $userInput = ""
  while ($userInput.ToLower() -ne "done" -and $userInput.ToLower() -ne "skip") {
    $userInput = Read-Host
    LineAcrossScreen
    ""
    if ($userInput.ToLower() -eq "done") {
      break
    }
    if ($userInput.ToLower() -eq "skip") {
      $skip = $true
      break
    }

    $context += "Human: $userInput`n"
  }

  if (-not($skip)) {
    Write-Host "Generating follow up questions..."
  
    LineAcrossScreen -Color Yellow
  
    # Generate clarifying questions
    $clarifyingQuestions = (Generate-Questions -prompt "Please generate a set of clarifying questions for the IT technician to answer. These questions should be based on the ticket notes and aim to address any missing information or unclear details in the original notes: $context" -OpenAIAPIKey $OpenAIAPIKey -TicketNotes $context) -split "`n"
  
    foreach ($question in $clarifyingQuestions) {
      if (![string]::IsNullOrEmpty($question)) {
        $answer = Read-Host -prompt "`n$question`n`n"
        if (![string]::IsNullOrEmpty($answer)) {
          $context += "assistant: $question`nHuman: $answer`n"
        }
        LineAcrossScreen -Color Yellow
      }
    }
  }

  Write-Host "Generating Ticket Notes..."

  LineAcrossScreen

  $response = Invoke-OpenAIAPI -prompt $context -OpenAIAPIKey $OpenAIAPIKey -PreviousResponse $context
  Write-Host ""
  Write-Host "$response"
  Write-Host ""

  LineAcrossScreen

  Stop-Transcript
}

function Invoke-vtsFastDownload {
  <#
  .SYNOPSIS
  This function downloads a file from a given URL using the fast download utility aria2.
  
  .DESCRIPTION
  Invoke-vtsFastDownload is a function that downloads a file from a specified URL to a specified path on the local system. 
  It first checks if the download path exists, if not, it creates it. 
  Then it sets the current location to the download path and sets the execution policy and security protocol. 
  It installs Chocolatey if not already installed and then installs aria2 using Chocolatey. 
  Finally, it downloads the file using aria2 and saves it at the specified location.
  
  .PARAMETER DownloadPath
  The path where the downloaded file will be saved. Default is "C:\temp".
  
  .PARAMETER URL
  The URL of the file to be downloaded. This is a mandatory parameter.
  
  .PARAMETER FileName
  The name of the file to be saved on the local system. This is a mandatory parameter.
  
  .EXAMPLE
  Invoke-vtsFastDownload -URL "http://example.com/file.zip" -FileName "file.zip"
  This will download the file from the specified URL and save it as "file.zip" in the default download path "C:\temp".
  
  .EXAMPLE
  Invoke-vtsFastDownload -DownloadPath "D:\downloads" -URL "http://example.com/file.zip" -FileName "file.zip"
  This will download the file from the specified URL and save it as "file.zip" in the specified download path "D:\downloads".
  
  .LINK
  Utilities
  #>
  param (
    [string]$DownloadPath = "C:\temp",
    [Parameter(Mandatory = $true)]
    [string]$URL,
    [Parameter(Mandatory = $true)]
    [string]$FileName
  )

  Write-Host "Checking if the download path $DownloadPath exists..."
  if (!(Test-Path $DownloadPath)) {
    Write-Host "Download path does not exist. Creating it now..."
    New-Item -ItemType Directory -Force -Path $DownloadPath
  }
  else {
    Write-Host "Download path exists. Proceeding with the download..."
  }

  Write-Host "Setting the current location to $DownloadPath..."
  Set-Location -Path $DownloadPath

  Write-Host "Setting execution policy and security protocol..."
  Set-ExecutionPolicy Bypass -Scope Process -Force
  [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072

  Write-Host "Installing Chocolatey..."
  Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

  Write-Host "Installing aria2 with Chocolatey..."
  & choco install aria2 -y

  Write-Host "`nDownloading file from $URL to $(Join-Path $DownloadPath $FileName)" -ForegroundColor Green
  & aria2c -x16 -s16 -k1M -c -o "$FileName" "$URL" --file-allocation=none

  Write-Host "`nDownload complete. File saved at $(Join-Path $DownloadPath $FileName)" -ForegroundColor Green
}

function Get-vtsFilePathCharacterCount {
  <#
  .SYNOPSIS
      This script gets the character count of file names in a directory and exports the data to a CSV file.
  
  .DESCRIPTION
      The Get-vtsFilePathCharacterCount function takes a directory path and an output file path as parameters. It then retrieves all the items in the directory, counts the number of characters in each item's full name, and stores this information in an array. The array is then exported to a CSV file at the specified output file path.
  
  .PARAMETER directoryPath
      The path of the directory to get the file character count from. This parameter is mandatory.
  
  .PARAMETER outputFilePath
      The path of the CSV file to output the results to. This parameter is mandatory.
  
  .EXAMPLE
      Get-vtsFilePathCharacterCount -directoryPath "C:\path\to\directory" -outputFilePath "C:\path\to\outputfile.csv"
      This command gets the character count of all file names in the directory "C:\path\to\directory" and exports the data to the CSV file at "C:\path\to\outputfile.csv".
  
  .LINK
      File Management
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Please enter the directory path in the format 'C:\path\to\directory'")]
    [string]$directoryPath,
    [Parameter(Mandatory = $true, HelpMessage = "Please enter the output file path in the format 'C:\path\to\outputfile.csv'")]
    [string]$outputFilePath
  )

  Write-Host "Starting to get file character count from directory: $directoryPath"

  $items = Get-ChildItem -Path $directoryPath -Recurse

  Write-Host "Found $($items.Count) items in the directory"

  $output = @()

  foreach ($item in $items) {
    $output += New-Object PSObject -Property @{
      "Fullname"                         = $item.FullName
      "Number of characters in Fullname" = $item.FullName.Length
    }
  }

  Write-Host "Processed all items, now exporting to CSV file: $outputFilePath"

  $output | Export-Csv -Path $outputFilePath -NoTypeInformation

  Write-Host "Export completed successfully"
}

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

function Get-vtsScreenshot {
  <#
  .SYNOPSIS
     A script to take a screenshot of the current screen and save it to a specified path.
  
  .DESCRIPTION
     This script uses the System.Windows.Forms and System.Drawing assemblies to capture a screenshot of the current screen. 
     The screenshot is saved as a .png file at the path specified by the $Path parameter. 
     If no path is specified, the screenshot will be saved in the temp folder with a timestamp in the filename.
  
  .PARAMETER Path
     The path where the screenshot will be saved. If not specified, the screenshot will be saved in the temp folder with a timestamp in the filename.
  
  .EXAMPLE
     Get-vtsScreenshot -Path "C:\Users\Username\Pictures\Screenshot.png"
     This command will take a screenshot and save it as Screenshot.png in the Pictures folder of the user Username.
  
  .LINK
      Utilities
  #>
  param (
    [string]$Path = "$env:temp\$(Get-Date -f yyyy-MM-dd-HH-mm)-Screenshot.png"
  )

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing

  $Screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
  $Width = $Screen.Width
  $Height = $Screen.Height
  $Left = $Screen.Left
  $Top = $Screen.Top

  $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
  $graphic = [System.Drawing.Graphics]::FromImage($bitmap)
  $graphic.CopyFromScreen($Left, $Top, 0, 0, $bitmap.Size)

  $bitmap.Save($Path)

  Write-Host "Screenshot saved at $Path"
}

function Convert-vtsScreenToAscii {
  <#
  .SYNOPSIS
  This script takes a screenshot of the current display and converts it to ASCII art to be displayed directly in the console.
  
  .DESCRIPTION
  The script uses the Get-vtsScreenshot function to capture a screenshot and save it to a specified directory. It then uses the ImageToAscii class to convert the screenshot into ASCII art. The ASCII art is then printed to the console.
  
  .PARAMETER ImageDirectory
  The directory where the screenshot will be saved. Default is "C:\temp\".
  
  .PARAMETER NumOfImages
  The number of screenshots to take and convert to ASCII art. Default is 3.
  
  .PARAMETER SleepInterval
  The interval in seconds between each screenshot. Default is 1.
  
  .EXAMPLE
  PS C:\> .\screen-to-ascii.ps1 -ImageDirectory "C:\screenshots\" -NumOfImages 5 -SleepInterval 2
  This example will take 5 screenshots at an interval of 2 seconds each, save them to the "C:\screenshots\" directory, and convert each one to ASCII art.
  
  .NOTES
  The ASCII art is generated with a fixed width of 100 characters. This is to maintain the aspect ratio of the ASCII characters.
  
  .LINK
  Utilities
  
  #>
  param (
    [string]$ImageDirectory = "$env:temp\",
    [int]$NumOfImages = 100,
    [int]$SleepInterval = 0
  )

  if ((whoami) -eq "nt authority\system") {
    Write-Error "Must run script as logged in user. Running as system doesn't work."
  }
  else {
    # Define the ImageToAscii class
    Add-Type -TypeDefinition @"
using System;
using System.Drawing;
public class ImageToAscii {
  public static string ConvertImageToAscii(string imagePath, int width) {
      Bitmap image = new Bitmap(imagePath, true);
      image = GetResizedImage(image, width);
      return ConvertToAscii(image);
  }
  private static Bitmap GetResizedImage(Bitmap original, int width) {
      int height = (int)(original.Height * ((double)width / original.Width) / 2); // Adjust for aspect ratio of ASCII characters
      var resized = new Bitmap(original, new Size(width, height));
      return resized;
  }
  private static string ConvertToAscii(Bitmap image) {
      string ascii = "";
      for (int h = 0; h < image.Height; h++) {
          for (int w = 0; w < image.Width; w++) {
              Color pixelColor = image.GetPixel(w, h);
              int grayScale = (pixelColor.R + pixelColor.G + pixelColor.B) / 3;
              ascii += GetAsciiCharForGrayscale(grayScale);
          }
          ascii += "\\n";
      }
      return ascii;
  }
  private static char GetAsciiCharForGrayscale(int grayScale) {
      string asciiChars = "@%#*+=-:. ";
      return asciiChars[grayScale * asciiChars.Length / 256];
  }
}
"@ -ReferencedAssemblies System.Drawing *>$null -ErrorAction SilentlyContinue

    try {
      if (!(Test-Path -Path $ImageDirectory)) {
        New-Item -ItemType Directory -Force -Path $ImageDirectory
      }
    
    
      for ($i = 1; $i -le $NumOfImages; $i++) {
        $imagePath = Join-Path -Path $ImageDirectory -ChildPath "image_$i.png"
        Get-vtsScreenshot -Path $imagePath *>$null
    
          
        # Convert the image to ASCII art
        $asciiArt = [ImageToAscii]::ConvertImageToAscii($imagePath, 100) # Adjust the width for ASCII character aspect ratio
          
        # Print the ASCII art to the host a chunk at a time
        $lines = $asciiArt -split "\\n"
          
        foreach ($line in $lines) {
          if ($line.Length -gt 100) {
            # If the line is longer than 100 characters, split it into chunks
            $chunks = $line -split "(.{100})", -1, 'RegexMatch'
            Clear-Host
            foreach ($chunk in $chunks) {
              if ($chunk -ne "") {
                Write-Host $chunk
              }
            }
          }
          else {
            # If the line is not longer than 100 characters, just print it
            Write-Host $line
          }
    
        }
        Start-Sleep $SleepInterval
      }
    
    }
    finally {
      <#Do this after the try block regardless of whether an exception occurred or not#>
      Remove-Item "$ImageDirectory\image_*.png" -Recurse -Force -Confirm:$false
    }

  }

}

function Get-vts365TeamsMembershipReport {
  <#
  .SYNOPSIS
      This script generates a report of Microsoft Teams' memberships.
  
  .DESCRIPTION
      The Get-vts365TeamsMembershipReport function generates a report of all Microsoft Teams' memberships. 
      It lists all the teams along with their owners, members, and guests. 
      If a team does not have any members or guests, it will be noted in the report. 
      The report is then copied to the clipboard.
  
  .PARAMETER None
      This function does not take any parameters.
  
  .EXAMPLE
      PS C:\> .\Get-vts365TeamsMembershipReport.ps1
      This command runs the script and generates the report.
  
  .NOTES
      You need to have the MicrosoftTeams module installed and be connected to Microsoft Teams for this script to work.
  
  .LINK
      M365
  #>
  if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Install-Module -Name MicrosoftTeams
  }
  import-module MicrosoftTeams
  Connect-MicrosoftTeams

  $teamsToAudit = get-team | Out-GridView -OutputMode Multiple -Title "Select one or more Teams, then click OK."

  $report = @()

  foreach ($team in $teamsToAudit) {
    $owners = $team | Get-TeamUser -Role Owner | Select-Object -ExpandProperty User
    $members = $team | Get-TeamUser -Role Member | Select-Object -ExpandProperty User
    $guests = $team | Get-TeamUser -Role Guest | Select-Object -ExpandProperty User
    $report += "====================`n"
    $report += "Team Name: $($team.Displayname)`n"
    $report += "`nOwners:`n"
    foreach ($owner in $owners) {
      $report += "`t• $owner`n"
    }
    $report += "`nMembers:`n"
    if ($members) {
      foreach ($member in $members) {
        $report += "`t• $member`n"
      }
    }
    else {
      $report += "`tNo Members`n"
    }
    $report += "`nGuests:`n"
    if ($guests) {
      foreach ($guest in $guests) {
        $report += "`t• $guest`n"
      }
    }
    else {
      $report += "`tNo Guests`n"
    }
  }

  $report | Set-Clipboard

  Write-Host "Results copied to clipboard." -f Yellow
}

function Ping-vtsList {
  <#
  .SYNOPSIS
      This function performs a ping operation on a list of IP addresses or hostnames and compiles a report.
  
  .DESCRIPTION
      The Ping-vtsList function conducts a ping test on a list of IP addresses or hostnames that are provided in a file. 
      It then compiles a comprehensive report detailing the status and response time for each target. 
      Additionally, the function provides an option to export this report to an HTML file for easy viewing and sharing.
  
  .PARAMETER TargetIPAddressFile
      This parameter requires the full path to the file that contains the target IP addresses or hostnames.
  
  .PARAMETER ReportTitle
      This parameter allows you to set the title of the report. If not specified, the default title is "Ping Report".
  
  .EXAMPLE
      Ping-vtsList -TargetIPAddressFile "C:\temp\IPList.txt" -ReportTitle "Server Ping Report"
  
      In this example, the function pings the IP addresses or hostnames listed in the file "C:\temp\IPList.txt" and generates a report with the title "Server Ping Report".
  
  .LINK
      Network
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the full path to the file containing the target IP addresses / hostnames.")]
    $TargetIPAddressFile,
    $ReportTitle = "Ping Report"
  )

  $PowerShellVersion = $PSVersionTable.PSVersion.Major

  if ($PowerShellVersion -lt 7) {
    Write-Warning "For enhanced performance through parallel pinging, please consider upgrading to PowerShell 7+."
  }

  $global:IPList = Get-Content $TargetIPAddressFile

  $global:Report = @()

  switch ($PowerShellVersion) {
    7 { 
      $global:Report = $global:IPList | ForEach-Object -Parallel  {
        Clear-Variable Ping -ErrorAction SilentlyContinue
        try {
          $Ping = Test-Connection $_ -Count 1 -ErrorAction Stop
            
        }
        catch {
          $PingError = $_.Exception.Message
        }
        if ((($Ping).StatusCode -eq 0) -or (($Ping).Status -eq "Success")) {
          [pscustomobject]@{
            Target       = $_
            Status       = "OK"
            ResponseTime = if ($Ping.ResponseTime) { $Ping.ResponseTime } elseif ($Ping.Latency) { $Ping.Latency }
          }
        } else {
          [pscustomobject]@{
            Target       = $_
            Status       = if ($PingError) { $PingError } else { "Failed" }
            ResponseTime = "n/a"
          }
            
        }
      }
            
    }
    Default {
      foreach ($IP in $global:IPList) {
        Clear-Variable Ping
        try {
          $Ping = Test-Connection $IP -Count 1 -ErrorAction Stop
            
        }
        catch {
          $PingError = $_.Exception.Message
        }
        if ((($Ping).StatusCode -eq 0) -or (($Ping).Status -eq "Success")) {
          $global:Report += [pscustomobject]@{
            Target       = $IP
            Status       = "OK"
            ResponseTime = if ($Ping.ResponseTime) { $Ping.ResponseTime } elseif ($Ping.Latency) { $Ping.Latency }
          }
        } else {
          $global:Report += [pscustomobject]@{
            Target       = $IP
            Status       = if ($PingError) { $PingError } else { "Failed" }
            ResponseTime = "n/a"
          }
            
        }
      }

    }
  }

  $global:Report | Out-Host

  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"
  
  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
    # Check if PSWriteHTML module is installed, if not, install it
    if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
      Install-Module -Name PSWriteHTML -Force -Confirm:$false
    }
        
    # Export the results to an HTML file using the PSWriteHTML module
    $global:Report | Out-HtmlView -Title $ReportTitle
  }
}

function Start-vtsPathPing {
  <#
  .SYNOPSIS
  This script performs a pathping test on a list of servers and logs the results.
  
  .DESCRIPTION
  The Start-vtsPathPing function in this script performs a pathping test on a list of servers. The results of the test are logged in a text file in a specified directory. If the directory does not exist, it is created. The progress of the test is displayed in the console.
  
  .PARAMETER Servers
  An array of server names or IP addresses to test. The default values are "uvwss.ubervoip.net" and "ubervoice.ubervoip.net".
  
  .PARAMETER LogDir
  The directory where the log file will be saved. The default value is "$env:temp\Network Testing\".
  
  .EXAMPLE
  Start-vtsPathPing -Servers "192.168.1.1", "192.168.1.2" -LogDir "C:\Logs\"
  
  This example performs a pathping test on the servers "192.168.1.1" and "192.168.1.2". The results are saved in the "C:\Logs\" directory.
  
  .LINK
  Network
  #>
  param(
      [Parameter(Mandatory=$false)]
      $Servers = @(
          "uvwss.ubervoip.net",
          "ubervoice.ubervoip.net"
      ),
      [Parameter(Mandatory=$false)]
      $LogDir = "$env:temp\Network Testing\"
  )

  if (!(Test-Path -Path $LogDir)) {
      New-Item -ItemType Directory -Path $LogDir | Out-Null
  }

  $Timestamp = Get-Date -f yyyy-MM-dd_HH_mm_ss

  # Loop through each server
  $ServerCount = $Servers.Count
  
  for ($i=0; $i -lt $ServerCount; $i++) {
    $StartTime = Get-Date
    $Server = $Servers[$i]
    $Progress = [math]::Round((($i / $ServerCount) * 100), 2)
    $EstimatedFinishTime = $StartTime.AddMinutes(9 * ($ServerCount - $i)).ToShortTimeString()
    Write-Progress -Activity "Running PATHPING on Target: $Server ($(($i+1)) of $ServerCount)" -Status "Estimated Completion Time: $EstimatedFinishTime (~$(9 * ($ServerCount - $i)) min)" -PercentComplete $Progress
    (Get-Date) | Out-File -Append -FilePath "$($Logdir)\$($Timestamp)-pathping.txt"
    PATHPING.EXE $Server >> "$($Logdir)\$($Timestamp)-pathping.txt"
  }

  # Clear the progress bar
  Write-Progress -Activity "Pathping process completed" -Completed

  & notepad.exe "$($Logdir)\$($Timestamp)-pathping.txt"
}

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

function New-vts365User {
  <#
  .SYNOPSIS
  Creates new Microsoft 365 users from a CSV file.
  
  .DESCRIPTION
  The New-vts365User function creates new Microsoft 365 users with the specified domain and usage location. It requires a CSV file path as input, which should contain the user details. The function also allows setting a password length.
  
  The CSV file must contain the following columns:
  "FirstName","LastName","PrimaryEmail","RecoveryPhone","RecoveryEmail"
  
  Example CSV content:
  "FirstName","LastName","PrimaryEmail","RecoveryPhone","RecoveryEmail"
  "John","Doe","john.doe@contoso.com","+1 234567890","john.recovery@email.com"
  "Jane","Smith","jane.smith@contoso.com","+1 987654321","jane.recovery@email.com"
  
  .PARAMETER CsvPath
  The path to the CSV file containing user information to be imported.
  
  .PARAMETER UsageLocation
  The usage location for the new user. Default is 'US'.
  
  .PARAMETER PasswordLength
  The length of the password to be generated for the new user. Default is 16 characters.
  
  .PARAMETER CustomPassword
  Specifies a custom password for the new user. If not provided, a random password will be generated.
  
  .PARAMETER CreateTemplate
  Switch to create a CSV template for user input.
  
  .EXAMPLE
  New-vts365User -CsvPath "C:\Users\example\userlist.csv" -Domain "contoso.com"
  
  This command creates new users from the specified CSV file with the domain 'contoso.com' and the default usage location and password length.
  
  .EXAMPLE
  New-vts365User -CsvPath "C:\Users\example\userlist.csv" -Domain "contoso.com" -UsageLocation "GB" -PasswordLength 20
  
  This command creates new users with the domain 'contoso.com', sets the usage location to 'GB', and generates passwords with a length of 20 characters.
  
  .NOTES
  Requires the MSOnline and AzureAD modules.
  
  .LINK
  M365
  
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({
      if (Test-Path $_ -PathType Leaf) {
        $true
      }
      else {
        throw "File not found or not a valid file: $_"
      }
    })]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("US", "CA", "MX", "GB", "DE", "FR", "JP", "AU", "BR", "IN")]
    [string]$UsageLocation = "US",

    [Parameter(Mandatory = $false)]
    [ValidateRange(12, 32)]
    [int]$PasswordLength = 16,

    [Parameter(Mandatory = $false)]
    [string]$CustomPassword,

    [switch]$CreateTemplate,

    [switch]$UseLast4PhoneDigitsAsPassword,

    $keyPath = "$($env:USERPROFILE)\Documents\$(Get-Date -f MMddyymmss)_temp.txt"
  )

  if($UseLast4PhoneDigitsAsPassword){
    $codeword = Read-Host "Enter a code word to use for generating passwords"
  }

  # Function to retry operations
  function Invoke-WithRetry {
    param(
      [ScriptBlock]$Action,
      [int]$MaxAttempts = 3,
      [int]$DelaySeconds = 2
    )

    $attempts = 0
    do {
      $attempts++
      try {
        return & $Action
      }
      catch {
        if ($attempts -eq $MaxAttempts) { throw }
        Start-Sleep -Seconds $DelaySeconds
      }
    } while ($attempts -lt $MaxAttempts)
  }

  # Function to create a random password
  function Get-RandomPassword {
    param ([int]$Length = 16)
    $charSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_-+=[]{}|;:,.<>?'.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($Length)
    $rng.GetBytes($bytes)

    $result = New-Object char[]($Length)
    for ($i = 0 ; $i -lt $Length ; $i++) {
      $result[$i] = $charSet[$bytes[$i] % $charSet.Length]
    }
    return (-join $result)
  }

  # Check and install required modules
  function Install-RequiredModules {
    $requiredModules = @("MSOnline", "Microsoft.Graph")
    foreach ($module in $requiredModules) {
      if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Verbose "Installing $module module..."
        Install-Module -Name $module -Force -Scope CurrentUser
      }
    }
  }

  # Connect to Microsoft 365 and Microsoft Graph
  function Connect-Services {
    try {
      if ((Get-MsolCompanyInformation).DisplayName) {
        Write-Verbose "Already connected to $((Get-MsolCompanyInformation).DisplayName)'s Microsoft 365 tenant."
      }
      else {
        Connect-MsolService
      }
      Connect-MgGraph -Scopes "User.ReadWrite.All", "UserAuthenticationMethod.ReadWrite.All", "Directory.AccessAsUser.All", "Directory.ReadWrite.All" -UseDeviceAuthentication -ForceRefresh
    }
    catch {
      Write-Error "Failed to connect to Microsoft 365 or Microsoft Graph. Error: $_"
      throw
    }
  }

  # Create CSV Template if requested
  function Create-CSVTemplate {
    $templatePath = Join-Path $env:USERPROFILE "Documents\$(Get-Date -f MMddyymmss)_M365NewUser.csv"
    $templateData = @(
      [PSCustomObject]@{
        FirstName     = ""
        LastName      = ""
        PrimaryEmail  = ""
        RecoveryPhone = ""
        RecoveryEmail = ""
      }
    )

    try {
      $templateData | Export-Csv -Path $templatePath -NoTypeInformation
      while(-not(Test-Path $templatePath)) {
        Start-Sleep -Seconds 1
      }
      Invoke-Item $templatePath
      Write-Host "The csv template is being created, and will be automatically opened momentarily. Enter the user details in the new csv template. Once finished, save the template, then type 'next' to proceed."
      do {
        $response = Read-Host
      } while ($response -ne "next")
      return $templatePath
    }
    catch {
      Write-Error "Failed to create template CSV file: $_"
      throw
    }
  }

  # Import and validate CSV data
  function Import-CSVData {
    param(
      [string]$Path
    )
    try {
      $users = Import-Csv $Path
      $requiredColumns = @("FirstName", "LastName", "PrimaryEmail", "RecoveryPhone", "RecoveryEmail")
      $missingColumns = $requiredColumns | Where-Object { $_ -notin $users[0].PSObject.Properties.Name }
      if ($missingColumns) {
        throw "CSV is missing required columns: $($missingColumns -join ', ')"
      }
      return $users
    }
    catch {
      Write-Error "Failed to parse CSV data: $_"
      throw
    }
  }

  # Process each user
  function Process-Users {
    param(
      [array]$Users
    )

    foreach ($user in $Users) {
      if ($user.LastName) {
        $displayName = "$($user.FirstName) $($user.LastName)"
      }
      else {
        $displayName = $user.FirstName
      }

      $userPrincipalName = $user.PrimaryEmail

      if ($CustomPassword) {
        $password = $CustomPassword
        "$($user.PrimaryEmail) $password" | out-file -filepath $keypath -Append -Force
      } elseif ($UseLast4PhoneDigitsAsPassword) {
        do {
          $last4 = ($user.RecoveryPhone -replace "[^0-9]", "")[-4..-1] -join ""
          if (-not($last4)){
            $last4 = 1234
          }
          $password = $codeword+$last4

        } until (
          $password.Length -ge 12
        )
        "$($user.PrimaryEmail) $password" | out-file -filepath $keypath -Append -Force
      }
      else {
        $password = Get-RandomPassword -Length $PasswordLength
        "$($user.PrimaryEmail) $password" | out-file -filepath $keypath -Append -Force
      }
      $PasswordProfile = @{
        Password                      = $password
        ForceChangePasswordNextSignIn = $true
      }

      try {
        # Check if user exists in Microsoft Graph
        $existingUser = Get-MgUser -Filter "userPrincipalName eq '$($user.PrimaryEmail)'" -ErrorAction SilentlyContinue

        if ($existingUser) {
          Write-Verbose "User $($user.PrimaryEmail) already exists. Updating details..."

          if ($CustomPassword) {
            if ($user.LastName) {
              Update-MgUser -UserId $existingUser.Id -DisplayName "$displayName" -GivenName $user.FirstName -Surname $user.LastName -UsageLocation $UsageLocation -PasswordProfile $PasswordProfile
            }
            else {
              Update-MgUser -UserId $existingUser.Id -DisplayName "$displayName" -GivenName $user.FirstName -UsageLocation $UsageLocation -PasswordProfile $PasswordProfile
            }
          }
          else {
            if ($user.LastName) {
              Update-MgUser -UserId $existingUser.Id -DisplayName "$displayName" -GivenName $user.FirstName -Surname $user.LastName -UsageLocation $UsageLocation
            }
            else {
              Update-MgUser -UserId $existingUser.Id -DisplayName "$displayName" -GivenName $user.FirstName -UsageLocation $UsageLocation
            }
          }

          # Update aliases (still using MSOnline as Graph doesn't have a direct equivalent)
          $currentAliases = Get-MsolUser -UserPrincipalName $user.PrimaryEmail | Select-Object -ExpandProperty ProxyAddresses
          if ($currentAliases -notcontains "smtp:$userPrincipalName") {
            # Set-MsolUser -UserPrincipalName $user.PrimaryEmail -EmailAddresses @($currentAliases + "smtp:$userPrincipalName")
          }

          Write-Output "User updated: $displayName ($($user.PrimaryEmail))"
        }
        else {
          Write-Verbose "Creating new user: $displayName"

          try {
            if ($user.LastName) {
              # Create new user in Microsoft Graph
              $newUser = New-MgUser -DisplayName "$displayName" `
                -GivenName $user.FirstName `
                -Surname $user.LastName `
                -UserPrincipalName $user.PrimaryEmail `
                -UsageLocation $UsageLocation `
                -PasswordProfile $PasswordProfile `
                -AccountEnabled:$true `
                -MailNickname ($user.FirstName.ToLower())
            }
            else {
              # Create new user in Microsoft Graph
              $newUser = New-MgUser -DisplayName "$displayName" `
                -GivenName $user.FirstName `
                -UserPrincipalName $user.PrimaryEmail `
                -UsageLocation $UsageLocation `
                -PasswordProfile $PasswordProfile `
                -AccountEnabled:$true `
                -MailNickname ($user.FirstName.ToLower())
            }

            Write-Output "User created: $displayName ($($user.PrimaryEmail))"
          }
          catch {
            Write-Error "Failed to create or update user $displayName. Error: $_"
          }
        }
      }
      catch {
        Write-Error "Failed to create or update user $displayName. Error: $_"
      }
    }

    Write-Output "All users have been processed. Waiting for user sync to complete..."

    do {
      Start-Sleep -Seconds 30
      Get-MgUser -Filter "userPrincipalName eq '$($user.PrimaryEmail)'" -ErrorAction SilentlyContinue
    } until (
      $?
    )

    foreach ($user in $Users) {
      $newlyCreatedUser = Get-MgUser -Filter "userPrincipalName eq '$($user.PrimaryEmail)'"
      if ($user.RecoveryPhone) {
        Invoke-WithRetry -Action {
          # Set phone authentication method
          $phoneParams = @{
            PhoneNumber = "$($user.RecoveryPhone)"
            PhoneType   = "mobile"
          }
          New-MgUserAuthenticationPhoneMethod -UserId $newlyCreatedUser.Id -BodyParameter $phoneParams
        }
      }

      if ($user.RecoveryEmail) {
        Invoke-WithRetry -Action {
          # Set email authentication method
          $emailParams = @{
            EmailAddress = $user.RecoveryEmail
          }
          New-MgUserAuthenticationEmailMethod -UserId $newlyCreatedUser.Id -BodyParameter $emailParams
        }
      }
    }
  }

  # Disconnect from services
  function Disconnect-Services {
    # Disconnect-MgGraph
    Write-Verbose "Disconnected from Microsoft 365 and Microsoft Graph services."
  }

  # Main logic starts here
  if (-not($CsvPath) -and -not($CreateTemplate)) {
    Write-Error "Please provide a CSV file path or use the -CreateTemplate switch to generate a template CSV file."
    return
  }

  if ($CreateTemplate) {
    $CsvPath = Create-CSVTemplate
  }

  Install-RequiredModules
  Connect-Services
  $users = Import-CSVData -Path $CsvPath
  Process-Users -Users $users
  Disconnect-Services
  Write-Verbose "All users have been processed."
  ii $keyPath
}

function Get-vtsWlanProfilesAndKeys {
  <#
  .SYNOPSIS
  Retrieves the network profiles and their associated keys on the local machine.
  
  .DESCRIPTION
  This function lists all wireless network profiles stored on the local machine along with their clear text keys (passwords). It uses the 'netsh' command-line utility to query the profiles and extract the information.
  
  .EXAMPLE
  PS C:\> Get-vtsWlanProfilesAndKeys
  
  This command will display a list of all wireless network profiles and their associated keys.
  
  .NOTES
  This function requires administrative privileges to reveal the keys for the network profiles.
  
  .LINK
  Network
  #>
    netsh wlan show profiles | 
    Where-Object {$_ -match ' : '} | 
    ForEach-Object {$_.split(':')[1].trim()} | 
    ForEach-Object {
        $networkName = $_
        netsh wlan show profile name="$_" key=clear
    } | 
    Where-Object {$_ -match 'key content'} | 
    Select-Object @{Name='Network'; Expression={$networkName}}, @{Name='Key'; Expression={$_.split(':')[1].trim()}}
}

function Get-vts365LastSignIn {
  <#
  .SYNOPSIS
  Retrieves the last sign-in information for all Microsoft 365 users.
  
  .DESCRIPTION
  The Get-vts365LastSignIn function connects to Microsoft Graph, retrieves all users, and gathers the most recent sign-in log for each user. It outputs the user's display name, user principal name, last sign-in time, application used, and the client application used for the last sign-in.
  
  .EXAMPLE
  PS C:\> Get-vts365LastSignIn
  
  This command retrieves the last sign-in information for all Microsoft 365 users and displays it in the console.
  
  .EXAMPLE
  PS C:\> Get-vts365LastSignIn
  Do you want to export a report? (Y/N): Y
  
  This command retrieves the last sign-in information for all Microsoft 365 users, displays it in the console, and prompts the user to export the information to a CSV file.
  
  .NOTES
  Requires the Microsoft.Graph module.
  
  .LINK
  M365
  #>

  Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All"

  # Get all users
  $users = Get-MgUser -All

  # Initialize an array to store results
  $results = @()

  Write-Host "Processing users..."
  foreach ($user in $users) {
      # Get sign-in logs for the user
      $signInLogs = Get-MgAuditLogSignIn -Filter "userId eq '$($user.Id)'" -Top 1 -OrderBy "createdDateTime desc"
  
      if ($signInLogs) {
          $lastSignIn = $signInLogs.CreatedDateTime
          $appDisplayName = $signInLogs.AppDisplayName
          $clientAppUsed = $signInLogs.ClientAppUsed
      }
      else {
          $lastSignIn = "No sign-in record found"
          $appDisplayName = "N/A"
          $clientAppUsed = "N/A"
      }
  
      # Create a custom object with user info and last sign-in details
      $userInfo = [PSCustomObject]@{
          UserPrincipalName = $user.UserPrincipalName
          DisplayName       = $user.DisplayName
          LastSignInTime    = $lastSignIn
          AppDisplayName    = $appDisplayName
          ClientAppUsed     = $clientAppUsed
      }
  
      # Add the user info to the results array
      $results += $userInfo
  
      # Display info in the console
      Write-Host "User: $($userInfo.DisplayName) ($($userInfo.UserPrincipalName))"
      Write-Host "  Last Sign-In: $($userInfo.LastSignInTime)"
      Write-Host "  App: $($userInfo.AppDisplayName)"
      Write-Host "  Client: $($userInfo.ClientAppUsed)"
      Write-Host "------------------------"
  }

  # Ask user if they want to export a report
  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"

  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
      # Check if PSWriteHTML module is installed, if not, install it
      if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
          Install-Module -Name PSWriteHTML -Force -Confirm:$false
      }
    
      # Export the results to an HTML file using the PSWriteHTML module
      $results | Out-HtmlView
  }

  # Disconnect from Microsoft Graph
  Disconnect-MgGraph
}

function Get-vtsReadablePermissions {
  <#
  .SYNOPSIS
      Retrieves readable permissions for a given directory path.
  
  .DESCRIPTION
      The Get-vtsReadablePermissions function takes a file system path as input and returns a custom object array with the permissions of each subfolder. It includes details such as the user or group with access, the type of access, and the inheritance and propagation of the permissions.
  
  .EXAMPLE
      PS C:\> Get-vtsReadablePermissions -Path "C:\MyFolder"
      This example retrieves the permissions for all the immediate child folders of "C:\MyFolder".
  
  .NOTES
      This function requires at least PowerShell version 3.0 to run properly due to the usage of the Get-ChildItem -Directory parameter.
  
  .LINK
      File Management
  #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Check if the path exists
    if (!(Test-Path $Path)) {
        Write-Error "The specified path does not exist."
        return
    }

    # Initialize an empty array to store the permission information
    $permissions = @()

    $ChildFolders = Get-ChildItem -Path $Path -Directory -Depth 1 | Select-Object -expand fullname

    foreach ($Folder in $ChildFolders) {
        # Retrieve the ACL for the specified path
        $acl = Get-Acl -Path $Folder

        # Loop through each Access Control Entry (ACE)
        foreach ($ace in $acl.Access) {
            # Translate the IdentityReference (user or group)
            $user = $ace.IdentityReference
            $accessRights = $ace.FileSystemRights
            $accessType = $ace.AccessControlType

            # Translate Inheritance Flags
            $inheritance = switch ($ace.InheritanceFlags) {
                "None" { "This Folder Only" }
                "ContainerInherit" { "This Folder and Subfolders" }
                "ObjectInherit" { "This Folder and Files" }
                "ContainerInherit, ObjectInherit" { "This Folder, Subfolders, and Files" }
                default { "Unknown type" }
            }

            # Translate Propagation Flags
            $propagation = switch ($ace.PropagationFlags) {
                "None" { "None" }
                "NoPropagateInherit" { "Does not pass down" }
                "InheritOnly" { "Only affects children" }
                default { "Unknown propagation type" }
            }

            # Add to the permissions array
            $permissions += [PSCustomObject]@{
                Folder      = $Folder
                UserOrGroup = $user
                AccessType  = $accessType
                Rights      = $accessRights
                Inheritance = $inheritance
                Propagation = $propagation
            }
        }
    }

    # Output the permission information in a readable format
    $FilteredPermissions = $permissions |
    Where-Object UserOrGroup -ne "NT AUTHORITY\SYSTEM" |
    Where-Object UserOrGroup -ne "BUILTIN\Administrators" |
    Where-Object UserOrGroup -ne "CREATOR OWNER" |
    Where-Object UserOrGroup -notlike "*S-1-5*" |
    Sort-Object Folder, Rights, AccessType, Inheritance, UserOrGroup 
    
    $FilteredPermissions |
    Format-Table -AutoSize

    # Ask user if they want to export a report
    $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"
  
    if ($exportReport -eq "Y" -or $exportReport -eq "y") {
        # Set TLS1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        # Prompt for filepath
        $Filepath = Read-Host "Enter a file path to save the report (e.g., C:\Reports)"

        # Set alternate $Filepath if null or whitepace
        if ([string]::IsNullOrWhiteSpace($Filepath)) {
            Write-Host "`$Filepath is empty. Defaulting to C:\Windows\TEMP"
            $Filepath = "C:\Windows\TEMP" 
        }

        # Set alternate $Filepath if not valid
        if (-not(Test-Path $Filepath)) {
            Write-Host "$Filepath is not a valid path. Defaulting to C:\Windows\TEMP"
            $Filepath = "C:\Windows\TEMP" 
        }
        # Check if PSWriteHTML module is installed, if not, install it
        if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
            Install-Module -Name PSWriteHTML -Force -Confirm:$false
        }
        
        # Export the results to an HTML file using the PSWriteHTML module
        $FilteredPermissions | Out-HtmlView -Title "Permission Report - $(Get-Date -f "dddd MM-dd-yyyy HHmm")" -Filepath "$Filepath\Permission Report - $(Get-Date -f MM_dd_yyyy).html"
    }
}

function Set-vtsDirectoryOwnership {
  <#
  .SYNOPSIS
  This function takes ownership and grants full control permissions to a specified directory and its contents.
  
  .DESCRIPTION
  The Set-vtsDirectoryOwnership function takes ownership of a specified directory and grants full control permissions to the current user. It uses the takeown and icacls commands to achieve this. The function can be used to take ownership and set permissions on any directory.
  
  .PARAMETER DirectoryPath
  The path of the directory to take ownership of and grant permissions to. This parameter is mandatory.
  
  .EXAMPLE
  Set-vtsDirectoryOwnership -DirectoryPath "E:\Users\tmays"
  
  This example takes ownership of the "E:\Users\tmays" directory and grants full control permissions to the current user.
  
  .LINK
  File Management
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the path of the directory to take ownership of and grant permissions to.")]
    [string]$DirectoryPath
  )

  if (!(Test-Path -Path $DirectoryPath)) {
    Write-Error "The specified path does not exist."
    return
  }

  try {
    Write-Host "Taking ownership of $DirectoryPath..."
    takeown /F "$DirectoryPath" /R /D Y

    Write-Host "Granting full control permissions to $($env:USERNAME) on $DirectoryPath..."
    icacls "$DirectoryPath" /grant "$($env:USERNAME):F" /T /C /Q

    Write-Host "Ownership and permissions have been successfully updated for $DirectoryPath."
  }
  catch {
    Write-Error "An error occurred while updating ownership and permissions: $_"
  }
}

function Get-vtsShortcut {
  <#
  .SYNOPSIS
  Retrieves shortcut (.lnk) file information from specified paths.
  
  .DESCRIPTION
  The Get-vtsShortcut function retrieves detailed information about Windows shortcut files (.lnk), including their target paths, hotkeys, arguments, and icon locations. If no path is specified, it searches both the current user's and all users' Start Menu folders.
  
  .PARAMETER path
  Optional. The path to search for shortcut files. If not specified, searches Start Menu folders.
  
  .EXAMPLE
  PS> Get-vtsShortcut
  Returns all shortcuts from user and system Start Menu folders.
  
  .EXAMPLE
  PS> Get-vtsShortcut -path "C:\Users\Username\Desktop"
  Returns all shortcuts from the specified desktop folder.
  
  .LINK
  File Management
  #>
  param(
    $path = $null
  )
  
  $obj = New-Object -ComObject WScript.Shell

  if ($path -eq $null) {
    $pathUser = [System.Environment]::GetFolderPath('StartMenu')
    $pathCommon = $obj.SpecialFolders.Item('AllUsersStartMenu')
    $path = Get-ChildItem $pathUser, $pathCommon -Filter *.lnk -Recurse 
  }
  if ($path -is [string]) {
    $path = Get-ChildItem $path -Filter *.lnk
  }
  $path | ForEach-Object { 
    if ($_ -is [string]) {
      $_ = Get-ChildItem $_ -Filter *.lnk
    }
    if ($_) {
      $link = $obj.CreateShortcut($_.FullName)

      $info = @{}
      $info.Hotkey = $link.Hotkey
      $info.TargetPath = $link.TargetPath
      $info.LinkPath = $link.FullName
      $info.Arguments = $link.Arguments
      $info.Target = try { Split-Path $info.TargetPath -Leaf } catch { 'n/a' }
      $info.Link = try { Split-Path $info.LinkPath -Leaf } catch { 'n/a' }
      $info.WindowStyle = $link.WindowStyle
      $info.IconLocation = $link.IconLocation

      New-Object PSObject -Property $info
    }
  }
}

function Set-vtsShortcut {
  <#
  .SYNOPSIS
  Modifies properties of an existing Windows shortcut file.
  
  .DESCRIPTION
  The Set-vtsShortcut function allows modification of Windows shortcut (.lnk) file properties including the target path, hotkey, arguments, and icon location. It accepts input from the pipeline or direct parameters.
  
  .PARAMETER LinkPath
  The full path to the shortcut file to modify.
  
  .PARAMETER Hotkey
  The keyboard shortcut to assign to the shortcut file.
  
  .PARAMETER IconLocation 
  The path to the icon file and icon index to use.
  
  .PARAMETER Arguments
  The command-line arguments to pass to the target application.
  
  .PARAMETER TargetPath
  The path to the target file that the shortcut will launch.
  
  .EXAMPLE
  PS> Get-vtsShortcut "C:\shortcut.lnk" | Set-vtsShortcut -TargetPath "C:\NewTarget.exe"
  Modifies the target path of an existing shortcut.
  
  .EXAMPLE
  PS> Set-vtsShortcut -LinkPath "C:\shortcut.lnk" -Hotkey "CTRL+ALT+F"
  Sets a keyboard shortcut for an existing shortcut file.
  
  .EXAMPLE
  Get-ChildItem -Path "C:\Path\To\Your\Directory" -Include *.lnk -Recurse -file | 
  Select-Object -expand fullname |
  ForEach-Object { 
    Get-vtsShortcut $_ | 
    Where-Object TargetPath -like "*192.168.1.220*" | 
    ForEach-Object {
      Set-vtsShortcut -LinkPath "$($_.LinkPath)" -IconLocation "$($_.IconLocation)" -TargetPath "$(($_.TargetPath) -replace '192.168.1.220','192.168.5.220')"
    }
  }
  Recursively searches for all `.lnk` (shortcut) files in a specified directory. It then filters these shortcuts to find those whose target path contains the IP address `192.168.1.220`. For each matching shortcut, it updates the target path to replace `192.168.1.220` with `192.168.5.220`, while preserving the original link path and icon location.
  
  .LINK
  File Management
  #>
  param(
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    $LinkPath,
    $Hotkey,
    $IconLocation,
    $Arguments,
    $TargetPath
  )
  begin {
    $shell = New-Object -ComObject WScript.Shell
  }
  
  process {
    $link = $shell.CreateShortcut($LinkPath)

    $PSCmdlet.MyInvocation.BoundParameters.GetEnumerator() |
    Where-Object { $_.key -ne 'LinkPath' } |
    ForEach-Object { $link.$($_.key) = $_.value }
    $link.Save()
  }
}

function Set-vts365UserLicense {
  <#
  .SYNOPSIS
  This function manages licenses for a list of users by adding and removing licenses in a single operation.
  
  .DESCRIPTION
  The function Set-vts365UserLicense connects to the Graph API and manages licenses for selected users. 
  It allows selecting users from Exchange Online and choosing which licenses to add and remove.
  
  .EXAMPLE
  Set-vts365UserLicense
  
  .LINK
  M365
  #>
  param (
    [Parameter(Mandatory = $false, HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  if (-not $UserList) {
    Write-Host "Connecting to Exchange Online to retrieve mailboxes..."
    Connect-ExchangeOnline | Out-Null
    $UserList = get-user | 
    Where UserPrincipalName -notlike "*.onmicrosoft.com*" | 
    sort DisplayName | 
    select DisplayName, UserPrincipalName | 
    Out-GridView -Title "Select users to manage licenses for" -OutputMode Multiple | 
    Select -expand UserPrincipalName
  }

  if (-not $UserList) {
    Write-Host "No users selected." -ForegroundColor Red
    return
  }
  
  Write-Host "Connecting to Graph API..."
  Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All -Device

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

  Write-Host "`nSelect licenses to ADD:"
  $AddLicenses = $LicenseDisplay | 
      Out-GridView -Title "Select licenses to ADD (or cancel for none)" -OutputMode Multiple | 
      Select-Object -ExpandProperty SkuID

  Write-Host "`nSelect licenses to REMOVE:"
  $RemoveLicenses = $LicenseDisplay | 
      Out-GridView -Title "Select licenses to REMOVE (or cancel for none)" -OutputMode Multiple | 
      Select-Object -ExpandProperty SkuID

  $FormattedAddSKUs = @()
  foreach ($Sku in $AddLicenses) {
      $FormattedAddSKUs += @{SkuId = $Sku }
  }

  Write-Host "Processing license changes for users..."
  foreach ($User in $UserList) {
      Write-Host "Managing licenses for $User..."
      Set-MgUserLicense -UserId $User -AddLicenses $FormattedAddSKUs -RemoveLicenses $RemoveLicenses
  }
  Write-Host "Operation completed."
}

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

function New-vtsBulkADUsers {
  <#
  .SYNOPSIS
      Creates multiple Active Directory users from a CSV file.
  
  .DESCRIPTION
      The New-vtsBulkADUsers function automates the creation of multiple Active Directory users
      using data from a CSV file. It supports creating users in either a new or existing OU,
      sets up email properties, and triggers an AD sync after completion.
  
  .PARAMETER CsvPath
      Path to the CSV file containing user information. Required columns:
      - FirstName
      - LastName
      - PrimaryEmail
  
  .PARAMETER OUOption
      Specifies whether to create a new OU ('CreateNew') or select an existing one ('Select').
  
  .PARAMETER NewOUName
      Name of the new OU to create. Required when OUOption is 'CreateNew'.
  
  .PARAMETER NewOUPath
      Distinguished path where the new OU will be created. Required when OUOption is 'CreateNew'.
  
  .PARAMETER Password
      SecureString containing the default password for new users.
      Defaults to "WelcomeRocky123!".
  
  .PARAMETER LogPath
      Path where the operation log will be saved.
      Defaults to "C:\temp\User_Creation-<current-date>.csv".
  
  .EXAMPLE
      PS> New-vtsBulkADUsers -CsvPath "C:\Users.csv" -OUOption Select
      Creates users from Users.csv in an existing OU selected via GUI.
  
  .EXAMPLE
      PS> $params = @{
          CsvPath = "C:\Users.csv"
          OUOption = "CreateNew"
          NewOUName = "NewEmployees"
          NewOUPath = "DC=contoso,DC=com"
      }
      PS> New-vtsBulkADUsers @params
      Creates users from Users.csv in a new OU named "NewEmployees".
  
  .NOTES
      Author: VTS Systems
      Required Modules: ActiveDirectory
      CSV Format Example:
      FirstName,LastName,PrimaryEmail
      John,Doe,john.doe@contoso.com
  
  .LINK
      M365
  #>
  [CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string]$CsvPath,

  [Parameter(Mandatory)]
  [ValidateSet('CreateNew', 'Select')]
  [string]$OUOption,

  [Parameter()]
  [string]$NewOUName,

  [Parameter()]
  [string]$NewOUPath,

  [Parameter()]
  [SecureString]$Password = ("WelcomeRocky123!" | ConvertTo-SecureString -AsPlainText -Force),

  [Parameter()]
  [string]$LogPath = "C:\temp\User_Creation-$(get-date -Format 'yyyy-MM-dd').csv"
)

  function Write-Log {
    [CmdletBinding()]
    param(
      [Parameter()]
      [ValidateNotNullOrEmpty()]
      [string]$Message,

      [Parameter()]
      [ValidateNotNullOrEmpty()]
      [ValidateSet('Information', 'Warning', 'Error')]
      [string]$Severity = 'Information'
    )

    [pscustomobject]@{
      Time     = (Get-Date -f g)
      Message  = $Message
      Severity = $Severity
    } | Export-Csv -Path $LogPath -Append -NoTypeInformation

    Write-Host -Object $Message
  }

  # Create log directory if it doesn't exist
  if (-not (Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType File -Force
  }

  # Handle OU Selection
  if ($OUOption -eq 'CreateNew') {
    if (-not $NewOUName -or -not $NewOUPath) {
      throw "NewOUName and NewOUPath are required when creating a new OU"
    }
    try {
      $SelectedOU = New-ADOrganizationalUnit -Name $NewOUName -Path $NewOUPath -PassThru
      Write-Log -Message "Created new OU: $($SelectedOU.DistinguishedName)" -Severity Information
    }
    catch {
      throw "Failed to create new OU: $_"
    }
  }
  else {
    $SelectedOU = Get-ADOrganizationalUnit -Filter * | 
    Out-GridView -Title "Select Target OU" -OutputMode Single
    if (-not $SelectedOU) {
      throw "No OU selected"
    }
  }

  # Import user list
  $UserList = Import-Csv $CsvPath

  # Create users
  foreach ($user in $UserList) {
    $splat = @{
      Path              = $SelectedOU.DistinguishedName
      Name              = "$($user.FirstName) $($user.LastName)"
      UserPrincipalName = $user.PrimaryEmail
      SamAccountName    = "$($user.FirstName).$($user.LastName)"
      GivenName         = $user.FirstName
      Surname           = $user.LastName
      AccountPassword   = $Password
      Enabled           = $true 
    }

    try {
      New-ADUser @splat -Verbose -Confirm:$false
      Set-ADUser "$($user.FirstName).$($user.LastName)" -Add @{ProxyAddresses = "SMTP:$($user.PrimaryEmail)" }
      Set-ADUser "$($user.FirstName).$($user.LastName)" -EmailAddress $user.PrimaryEmail
      Write-Log -Message "Created user: $($user.FirstName) $($user.LastName)" -Severity Information
    }
    catch {
      Write-Log -Message "Failed to create user $($user.FirstName) $($user.LastName): $_" -Severity Error
    }
  }

  # Run delta sync
  try {
    Start-ADSyncSyncCycle -PolicyType Delta
    while (-not $?) { 
      Start-Sleep 5
      Start-ADSyncSyncCycle -PolicyType Delta 
    }
    Write-Log -Message "AD Sync completed successfully" -Severity Information
  }
  catch {
    Write-Log -Message "AD Sync failed: $_" -Severity Error
  }
  
}

function Manage-vtsFileRetention {
  <#
  .SYNOPSIS
      Manages file retention by resetting, marking, moving, and deleting files based on access times.
  
  .DESCRIPTION
      This script manages file retention in a specified directory. It performs the following actions:
      - Resets the LastWriteTime for recently accessed files.
      - Marks old files by setting an artificial LastWriteTime.
      - Moves old files to a "ToBeDeleted" directory.
      - Deletes files from the "ToBeDeleted" directory after a specified grace period.
  
  .PARAMETER BasePath
      The base path where the files are located. Default is "C:\qdstm*\out".
  
  .PARAMETER DaysBeforeSoftDelete
      The number of days before a file is considered old and marked for deletion. Default is 30 days.
  
  .PARAMETER DaysBeforeHardDelete
      The number of days before a file is permanently deleted from the "ToBeDeleted" directory. Default is 45 days.
  
  .PARAMETER ArtificialLastWriteTimeYears
      The number of years to subtract from the LastWriteTime to mark a file as old. Default is 100 years.
  
  .PARAMETER Extension
      The file extension to filter files. Default is "*.pdf".
  
  .EXAMPLE
      PS> .\Manage-vtsFileRetention.ps1 -BasePath "C:\example\path" -DaysBeforeSoftDelete 60 -DaysBeforeHardDelete 90 -Extension "*.txt"
      This example manages file retention for .txt files in the specified path, marking files as old after 60 days and deleting them after 90 days.
  
  .LINK
      File Management
  
  #>
  param (
      [string]$BasePath = "C:\qdstm*\out",
      [int]$DaysBeforeSoftDelete = 30,
      [int]$DaysBeforeHardDelete = 45,
      [int]$ArtificialLastWriteTimeYears = 100,
      [string]$Extension = ".pdf"
  )

  $OUT = Get-Item -Path $BasePath | Select-Object -ExpandProperty FullName

  $ToBeDeleted = Join-Path $OUT "ToBeDeleted"

  function Write-ActionLog {
      param(
          [string]$Message,
          [string]$Action,
    [string]$ToBeDeleted = $ToBeDeleted
      )
      $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
      $logMessage = "[$timestamp] $Action : $Message"
      $logPath = Join-Path $ToBeDeleted "file_management.log"
      Add-Content -Path $logPath -Value $logMessage
      Write-Host $logMessage
  }

  if ($OUT) {
      $ToBeDeleted = Join-Path $OUT "ToBeDeleted"
  
      if (-not (Test-Path $ToBeDeleted)) {
          try {
              New-Item -ItemType Directory -Path $ToBeDeleted
              Write-ActionLog -Message "Created ToBeDeleted directory at $ToBeDeleted" -Action "INIT"
          }
          catch {
              Write-ActionLog -Message "Failed to create directory: $ToBeDeleted. Error: $_" -Action "ERROR"
              exit
          }
      }
  
      # Reset LastWriteTime for recently accessed files
      try {
          Get-ChildItem -Path $OUT | 
          Where-Object { 
      $_.Extension -eq $Extension -and
      $_.LastAccessTime -gt (Get-Date).AddDays(-1) 
    } | 
          ForEach-Object { 
              $_.LastWriteTime = Get-Date
              Write-ActionLog -Message "Reset LastWriteTime for recently accessed file: $($_.FullName)" -Action "RESET"
          }
      }
      catch {
          Write-ActionLog -Message "Failed to update LastWriteTime for accessed files in $OUT. Error: $_" -Action "ERROR"
          exit
      }
  
      # Process old files
      try {
          Get-ChildItem -Path $OUT | 
          Where-Object {
      $_.Extension -eq $Extension -and
      $_.LastAccessTime -lt (Get-Date).AddDays( - ($DaysBeforeSoftDelete))
    } | 
          ForEach-Object { 
              $_.LastWriteTime = ($_.LastWriteTime).AddYears( - ($ArtificialLastWriteTimeYears))
              Write-ActionLog -Message "Marked file as old: $($_.FullName)" -Action "MARK"
          }
      }
      catch {
          Write-ActionLog -Message "Failed to update LastWriteTime for old files in $OUT. Error: $_" -Action "ERROR"
          exit
      }
  
      # Move old files block - modified to set new CreationTime
      try {
          Get-ChildItem -Path $OUT | 
          Where-Object { 
      $_.Extension -eq $Extension -and
      $_.LastAccessTime -lt (Get-Date).AddDays( - ($DaysBeforeSoftDelete)) 
    } | 
          ForEach-Object {
              $fileName = $_.FullName
              # Set CreationTime to now for the moved file
              $_.CreationTime = Get-Date
              Move-Item -Path $_.FullName -Destination $ToBeDeleted
              Write-ActionLog -Message "Moved to ToBeDeleted: $fileName (not accessed for $DaysBeforeSoftDelete days)" -Action "MOVE"
          }
      }
      catch {
          Write-ActionLog -Message "Failed to move files to $ToBeDeleted. Error: $_" -Action "ERROR"
          exit
      }
  
      # Delete expired files - modified block
      try {
          Get-ChildItem -Path $ToBeDeleted -Recurse | 
    Where-Object Extension -eq ".pdf" |
          Where-Object { 
              $movedToDeletedDate = $_.CreationTime
              $daysInToBeDeleted = ((Get-Date) - $movedToDeletedDate).Days
              $isOldFile = $_.LastWriteTime -lt (Get-Date).AddYears( - ($ArtificialLastWriteTimeYears))
              $daysInToBeDeleted -ge ($DaysBeforeHardDelete - $DaysBeforeSoftDelete) -and $isOldFile
          } | 
          ForEach-Object {
              $fileName = $_.FullName
              $daysInToBeDeleted = ((Get-Date) - $_.CreationTime).Days
              Remove-Item -Path $_.FullName -Force
              Write-ActionLog -Message "Permanently deleted: $fileName (was in ToBeDeleted for $daysInToBeDeleted days)" -Action "DELETE"
          }
      }
      catch {
          Write-Error "Failed to remove files from $ToBeDeleted. Error: $_"
          exit
      }
  }
}

function Enable-vtsTLS12AndStrongCrypto {
  <#
  .SYNOPSIS
  Enables TLS 1.2 for server and client, sets default secure protocols for WinHTTP, and enables strong cryptography for .NET Framework.
  
  .DESCRIPTION
  This function configures the necessary registry settings to enable TLS 1.2 for both server and client, sets the default secure protocols for WinHTTP, and enables strong cryptography for the .NET Framework.
  
  .EXAMPLE
  Enable-vtsTLS12AndStrongCrypto
  
  .LINK
  Network
  
  #>

  # Enable TLS 1.2 for Server and Client
  $TLSPaths = @(
      "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server",
      "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"
  )

  foreach ($path in $TLSPaths) {
      if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
      Set-ItemProperty -Path $path -Name "Enabled" -Value 1 -Type DWord
      Set-ItemProperty -Path $path -Name "DisabledByDefault" -Value 0 -Type DWord
  }

  # Set Default Secure Protocols for WinHTTP
  $WinHttpPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp"
  if (-not (Test-Path $WinHttpPath)) { New-Item -Path $WinHttpPath -Force | Out-Null }
  Set-ItemProperty -Path $WinHttpPath -Name "DefaultSecureProtocols" -Value 0xA00 -Type DWord

  # Enable Strong Cryptography for .NET Framework
  $DotNetPaths = @(
      "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
      "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319"
  )

  foreach ($path in $DotNetPaths) {
      if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
      Set-ItemProperty -Path $path -Name "SchUseStrongCrypto" -Value 1 -Type DWord
  }

  Write-Output "TLS 1.2 and .NET Strong Cryptography settings have been configured. Please restart the server for the changes to take effect."
}

function Get-vts365DynamicDistributionListRecipients {
  <#
  .SYNOPSIS
  This function retrieves the recipients of a given set of distribution lists from Exchange Online.
  
  .DESCRIPTION
  The Get-vts365DynamicDistributionListRecipients function connects to Exchange Online and retrieves the recipients of the specified dynamic distribution lists. The results are stored in an array and outputted at the end. It also provides an option to export the results to an HTML report.
  
  .PARAMETER DistributionList
  An array of distribution list names for which to retrieve recipients.
  
  .EXAMPLE
  PS C:\> Get-vts365DynamicDistributionListRecipients -DistributionList "Group1", "Group2"
  
  .LINK
  M365
  #>

  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }

  if((Get-ConnectionInformation).TenantID -like "7*4*b*5*-*4*3-*8*f-*f1*-6*"){
    $IncludeOfficeInfo = $true
  }

  # Get all dynamic distribution groups and suppress error output
  $distributionGroups = Get-DynamicDistributionGroup | Select-Object -ExpandProperty Name 2>$null
  
  # Initialize an empty array for the group table
  $GroupTable = @()
  $key = 1

  # Populate the group table with group names and corresponding keys
  foreach ($group in $distributionGroups) {
    $GroupTable += [pscustomobject]@{
      Key   = $key
      Group = $group
    }
    $key++
  }
  
  # Check if the current user is system authority
  if ($(whoami) -eq "nt authority\system") {
    $GroupTable | Out-Host
    $userInput = Read-Host "Please input the group numbers you wish to query, separated by commas. Alternatively, input '*' to search all groups."
    if ("$userInput" -eq '*') {
      Write-Host "Searching all available groups..."
      $SelectedGroups = $GroupTable.Group
    } else {
      Write-Host "Searching selected groups..."
      $SelectedGroups = $GroupTable | Where-Object Key -in ("$userInput" -split ",") | Select-Object -ExpandProperty Group
    }
  } else {
    $SelectedGroups = $GroupTable | Out-GridView -OutputMode Multiple | Select-Object -ExpandProperty Group
  }

  # Get recipients based on the selected groups and output their details
  $Results = foreach ($group in $SelectedGroups) {
    $DDLGroup = Get-DynamicDistributionGroup -Identity $group
    if($IncludeOfficeInfo){
    Get-Recipient -ResultSize Unlimited -RecipientPreviewFilter ($DDLGroup.RecipientFilter) | ForEach-Object {
      [pscustomobject]@{
        GroupName   = $DDLGroup
        DisplayName = $_.DisplayName
        Email       = $_.PrimarySmtpAddress
        PositionID  = $_.notes
        Title       = $_.title
        Office      = $_.office
      }
    } | Sort-Object DisplayName, Office, Title} else{
      Get-Recipient -ResultSize Unlimited -RecipientPreviewFilter ($DDLGroup.RecipientFilter) | ForEach-Object {
        [pscustomobject]@{
          GroupName   = $DDLGroup
          DisplayName = $_.DisplayName
          Email       = $_.PrimarySmtpAddress
        }
      } | Sort-Object DisplayName
    }
  }

  if($IncludeOfficeInfo){
    # Filter out results where PositionID, Title, or Office properties are null
    $Results = $Results | Where-Object { ![string]::IsNullOrWhiteSpace($_.PositionID) -and ![string]::IsNullOrWhiteSpace($_.Title) -and ![string]::IsNullOrWhiteSpace($_.Office) }
  }

  $Results | Format-Table -AutoSize | Out-Host

  # Ask user if they want to export a report
  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"
  
  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
    # Check if PSWriteHTML module is installed, if not, install it
    if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
      Install-Module -Name PSWriteHTML -Force -Confirm:$false
    }
        
    # Export the results to an HTML file using the PSWriteHTML module
    $Results | Out-HtmlView -Title "Dynamic Distribution List Recipients Report - $(Get-Date -f "dddd MM-dd-yyyy HHmm")" -Filepath "C:\Reports\Dynamic Distribution List Recipients Report - $(Get-Date -f MM_dd_yyyy).html"
  }
}

function Get-vts365DistributionListRecipients {
  <#
  .SYNOPSIS
  This function retrieves the recipients of a given set of distribution lists from Exchange Online.
  
  .DESCRIPTION
  The Get-vts365DistributionListRecipients function connects to Exchange Online and retrieves the recipients of the specified distribution lists. The results are stored in an array and outputted at the end. It also provides an option to export the results to an HTML report.
  
  .EXAMPLE
  PS C:\> Get-vts365DistributionListRecipients
  
  .LINK
  M365
  #>
  if (-not(Get-ConnectionInformation)) {
      Connect-ExchangeOnline -ShowBanner:$false
  }

  if((Get-ConnectionInformation).TenantID -like "7*4*b*5*-*4*3-*8*f-*f1*-6*"){
      $IncludeOfficeInfo = $true
  }

  # Get all regular distribution groups and suppress error output
  $distributionGroups = Get-DistributionGroup | Select-Object -ExpandProperty Name 2>$null

  # Initialize an empty array for the group table
  $GroupTable = @()
  $key = 1

  # Populate the group table with group names and corresponding keys
  foreach ($group in $distributionGroups) {
      $GroupTable += [pscustomobject]@{
          Key   = $key
          Group = $group
      }
      $key++
  }

  # Check if the current user is system authority
  if ($(whoami) -eq "nt authority\system") {
      $GroupTable | Out-Host
      $userInput = Read-Host "Please input the group numbers you wish to query, separated by commas. Alternatively, input '*' to search all groups."
      if ("$userInput" -eq '*') {
          Write-Host "Searching all available groups..."
          $SelectedGroups = $GroupTable.Group
      } else {
          Write-Host "Searching selected groups..."
          $SelectedGroups = $GroupTable | Where-Object Key -in ("$userInput" -split ",") | Select-Object -ExpandProperty Group
      }
  } else {
      $SelectedGroups = $GroupTable | Out-GridView -OutputMode Multiple | Select-Object -ExpandProperty Group
  }

  # Get members based on the selected groups and output their details
  $Results = foreach ($group in $SelectedGroups) {
      $DLGroup = Get-DistributionGroup -Identity $group
      
      # Get owners first
      $Owners = $DLGroup.ManagedBy | ForEach-Object {
          $owner = Get-Recipient -Identity $_
          if($IncludeOfficeInfo){
              [pscustomobject]@{
                  GroupName         = $DLGroup
                  GroupEmail       = $DLGroup.PrimarySmtpAddress
                  DisplayName      = $owner.DisplayName
                  Email           = $owner.PrimarySmtpAddress
                  PositionID      = $owner.notes
                  Title           = $owner.title
                  Office          = $owner.office
                  Role            = "Owner"
              }
          } else {
              [pscustomobject]@{
                  GroupName         = $DLGroup
                  GroupEmail       = $DLGroup.PrimarySmtpAddress
                  DisplayName      = $owner.DisplayName
                  Email           = $owner.PrimarySmtpAddress
                  Role            = "Owner"
              }
          }
      }

      # Then get members
      if($IncludeOfficeInfo){
          $Members = Get-DistributionGroupMember -Identity $group | ForEach-Object {
              [pscustomobject]@{
                  GroupName         = $DLGroup
                  GroupEmail       = $DLGroup.PrimarySmtpAddress
                  DisplayName      = $_.DisplayName
                  Email           = $_.PrimarySmtpAddress
                  PositionID      = $_.notes
                  Title           = $_.title
                  Office          = $_.office
                  Role            = "Member"
              }
          }
      } else {
          $Members = Get-DistributionGroupMember -Identity $group | ForEach-Object {
              [pscustomobject]@{
                  GroupName         = $DLGroup
                  GroupEmail       = $DLGroup.PrimarySmtpAddress
                  DisplayName      = $_.DisplayName
                  Email           = $_.PrimarySmtpAddress
                  Role            = "Member"
              }
          }
      }
      
      # Combine owners and members
      $Owners + $Members | Sort-Object DisplayName, Office, Title
  }

  if($IncludeOfficeInfo){
      # Filter out results where PositionID, Title, or Office properties are null
      $Results = $Results | Where-Object { ![string]::IsNullOrWhiteSpace($_.PositionID) -and ![string]::IsNullOrWhiteSpace($_.Title) -and ![string]::IsNullOrWhiteSpace($_.Office) }
  }

  $Results | Format-Table -AutoSize | Out-Host

  # Ask user if they want to export a report
  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"

  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
      if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
          Install-Module -Name PSWriteHTML -Force -Confirm:$false
      }
      $Results | Out-HtmlView -Title "Distribution List Recipients Report - $(Get-Date -f "dddd MM-dd-yyyy HHmm")" -Filepath "C:\Reports\Distribution List Recipients Report - $(Get-Date -f MM_dd_yyyy).html"
  }
}

function Export-vts365DistributionListConfig {
  <#
  .SYNOPSIS
  Exports the configuration of selected distribution lists to a CSV file.
  
  .DESCRIPTION
  The Export-vts365DistributionListConfig function connects to Exchange Online and exports detailed configuration information about selected distribution lists to a CSV file. It includes information such as managed by users, members, moderators, and various distribution list settings.
  
  .PARAMETER FilePath
  The path where the CSV file will be saved. If not specified, defaults to the Downloads folder with filename format 'yyyy-MM-dd_HHmm_DistributionListConfig.csv'.
  
  .EXAMPLE
  Export-vts365DistributionListConfig
  
  This example prompts for distribution list selection and exports their configuration to the default location.
  
  .EXAMPLE
  Export-vts365DistributionListConfig -FilePath "C:\Reports\DLConfig.csv"
  
  This example exports the selected distribution lists' configuration to the specified file path.
  
  .NOTES
  Requires connection to Exchange Online PowerShell.
  
  .LINK
  M365 Migration Scripts
  #>
  [CmdletBinding()]
  param (
      [Parameter()]
      [String]
      $FilePath = (Join-Path -Path $env:USERPROFILE -ChildPath "Downloads\$(Get-Date -f 'yyyy-MM-dd_HHmm')_DistributionListConfig.csv")
  )
  
  # Get all distribution groups
  $Groups = Get-DistributionGroup -ResultSize Unlimited |
  Out-GridView -OutputMode Multiple -Title "Select Distribution Groups to Export"
  
  # Count the total number of distribution groups
  $totalGroups = $Groups.Count
  $currentGroupIndex = 0
  
  # Initialize a List to store the data
  $Report = [System.Collections.Generic.List[Object]]::new()
  
  # Loop through distribution groups
  foreach ($Group in $Groups) {
      $currentGroupIndex++
      $GroupDN = $Group.DistinguishedName
  
      # Get ManagedBy names and SMTP addresses properly
      $ManagedByNames = @()
      $ManagedBySmtp = @()
      if ($Group.ManagedBy) {
          $Group.ManagedBy | ForEach-Object {
              try {
                  $owner = Get-Recipient $_ -ErrorAction Stop
                  $ManagedByNames += $owner.DisplayName
                  $ManagedBySmtp += $owner.PrimarySmtpAddress
              }
              catch {
                  $ManagedByNames += $_
                  $ManagedBySmtp += $_
              }
          }
      }

      # Update progress bar
      $progressParams = @{
          Activity        = "Processing Distribution Groups"
          Status          = "Processing Group $currentGroupIndex of $totalGroups"
          PercentComplete = ($currentGroupIndex / $totalGroups) * 100
      }
  
      Write-Progress @progressParams
  
      $GroupMembers = Get-DistributionGroupMember $GroupDN -ResultSize Unlimited
      
      # Get required attributes directly within the output object
      $ReportLine = [PSCustomObject]@{
          DisplayName                            = $Group.DisplayName
          Name                                   = $Group.Name
          PrimarySmtpAddress                     = $Group.PrimarySmtpAddress
          EmailAddresses                         = ($Group.EmailAddresses -join ',')
          Domain                                 = $Group.PrimarySmtpAddress.ToString().Split("@")[1]
          Alias                                  = $Group.Alias
          GroupType                              = $Group.GroupType
          RecipientTypeDetails                   = $Group.RecipientTypeDetails
          Members                                = $GroupMembers.Name -join ','
          MembersPrimarySmtpAddress              = $GroupMembers.PrimarySmtpAddress -join ','
          ManagedBy                              = ($ManagedByNames -join ',')
          ManagedBySmtpAddress                   = ($ManagedBySmtp -join ',')
          HiddenFromAddressLists                 = $Group.HiddenFromAddressListsEnabled
          MemberJoinRestriction                  = $Group.MemberJoinRestriction
          MemberDepartRestriction                = $Group.MemberDepartRestriction
          AcceptMessagesOnlyFrom                 = ($Group.AcceptMessagesOnlyFrom.Name -join ',')
          AcceptMessagesOnlyFromDLMembers        = ($Group.AcceptMessagesOnlyFromDLMembers -join ',')
          AcceptMessagesOnlyFromSendersOrMembers = ($Group.AcceptMessagesOnlyFromSendersOrMembers -join ',')
          ModeratedBy                            = ($Group.ModeratedBy -join ',')
          BypassModerationFromSendersOrMembers   = ($Group.BypassModerationFromSendersOrMembers -join ',')
          ModerationEnabled                      = $Group.ModerationEnabled
          SendModerationNotifications            = $Group.SendModerationNotifications
          GrantSendOnBehalfTo                    = ($Group.GrantSendOnBehalfTo.Name -join ',')
      }
      $Report.Add($ReportLine)
  }
  
  # Clear progress bar
  Write-Progress -Activity "Processing Distribution Groups" -Completed
  
  # Sort the output by DisplayName and export to CSV file
  $Report | Sort-Object DisplayName | Export-Csv -Path $FilePath -NoTypeInformation

  $Report | Out-Host

  Write-Host "`n`nCSV file has been created at: $FilePath `n`n"

  # Ask the user if they want to open the file after it's created
  $OpenFile = Read-Host "Do you want to open the CSV file? (Y/N)"
  if ($OpenFile -eq 'Y') {
      # Open the CSV file in the default application
      Invoke-Item $FilePath
  }
}

function Import-vts365DistributionListConfig {
  <#
  .SYNOPSIS
  Imports and creates distribution lists from a CSV configuration file.
  
  .DESCRIPTION
  The Import-vts365DistributionListConfig function creates new distribution lists in Exchange Online based on configuration data from a CSV file. It supports creating security groups, distribution groups, and room lists with their respective settings and members.
  
  The function will prefix all new distribution list names with "C-" and will automatically handle different types of groups including:
  - MailUniversalSecurityGroup
  - MailUniversalDistributionGroup
  - RoomList
  
  .PARAMETER FilePath
  The path to the CSV file containing the distribution list configurations to import. The CSV file should include columns for DisplayName, Name, Alias, PrimarySmtpAddress, and other distribution list properties.
  
  .EXAMPLE
  Import-vts365DistributionListConfig -FilePath "C:\DistributionLists.csv"
  
  This example imports distribution list configurations from the specified CSV file and creates the corresponding groups in Exchange Online.
  
  .NOTES
  Requires an active connection to Exchange Online PowerShell.
  Make sure the CSV file contains all required columns and properly formatted data.
  
  .LINK
  M365 Migration Scripts
  #>
  [CmdletBinding()]
  param (
      [Parameter(Mandatory)]
      [String]$FilePath,
      [Parameter()]
      [String]$AppendToDisplayName,
      [Parameter()]
      [String]$DefaultDomain
  )

  if (-not($AppendToDisplayName)) {
      $AppendToDisplayName = Read-Host "Enter Company Abbreviation to Append to DisplayName (so we can find the groups later.)"
  }

  # Import CSV and get default domain
  $GroupsData = Import-Csv $FilePath
  if(-not($DefaultDomain)){
      $DefaultDomain = Get-AcceptedDomain | Where-Object Default -eq $True | Select-Object -ExpandProperty DomainName
  }

  foreach ($GroupData in $GroupsData) {
      $LocalPart = ($GroupData.PrimarySmtpAddress) -replace '@.*$'
      $DisplayName = $GroupData.DisplayName + " - $AppendToDisplayName"
      
      # Check if group exists
      $ExistingGroup = Get-DistributionGroup -Identity $DisplayName -ErrorAction SilentlyContinue
      if ($ExistingGroup) {
          Write-Host "Group exists: $DisplayName" -ForegroundColor Yellow
          continue
      }

      try {
          Write-Host "Creating group: $DisplayName" -ForegroundColor Cyan
          
          # Create group first without owners
          $NewGroupParams = @{
              DisplayName = $DisplayName
              Name = $GroupData.Name
              Alias = $GroupData.Alias
              PrimarySMTPAddress = "$LocalPart@$DefaultDomain"
          }

          # Add type-specific parameters
          switch ($GroupData.RecipientTypeDetails) {
              "MailUniversalSecurityGroup" { 
                  $NewGroupParams.Type = "Security" 
              }
              "RoomList" { 
                  $NewGroupParams.Roomlist = $true 
              }
          }

          # Create the group
          $NewGroup = New-DistributionGroup @NewGroupParams

          # Set basic properties
          $SetGroupParams = @{
              Identity = $NewGroup.DisplayName
              # HiddenFromAddressListsEnabled = $True
              MemberJoinRestriction = $GroupData.MemberJoinRestriction
              MemberDepartRestriction = $GroupData.MemberDepartRestriction
              RequireSenderAuthenticationEnabled = [System.Convert]::ToBoolean($GroupData.RequireSenderAuthenticationEnabled)
          }

          if (-not [string]::IsNullOrWhiteSpace($GroupData.Notes)) {
              $SetGroupParams.Description = $GroupData.Notes
          }

          Set-DistributionGroup @SetGroupParams

          # Add owners separately
          if (-not [string]::IsNullOrWhiteSpace($GroupData.ManagedBy)) {
              $Owners = $GroupData.ManagedBy -split ',' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
              Write-Host "Adding owners..." -ForegroundColor Cyan
              
              foreach ($Owner in $Owners) {
                  try {
                      Set-DistributionGroup -Identity $NewGroup.DisplayName -ManagedBy @{Add=$Owner} -ErrorAction Stop
                      Write-Host "Added owner: $Owner" -ForegroundColor Green
                  }
                  catch {
                      Write-Host "Failed to add owner $Owner : $_" -ForegroundColor Yellow
                  }
              }
          }

          Start-Sleep -Seconds 5

          # Add members
          if (-not [string]::IsNullOrWhiteSpace($GroupData.MembersPrimarySmtpAddress)) {
              $Members = $GroupData.MembersPrimarySmtpAddress -split ',' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
              foreach ($Member in $Members) {
                  try {
                      Add-DistributionGroupMember -Identity $NewGroup.PrimarySmtpAddress -Member $Member -BypassSecurityGroupManagerCheck -ErrorAction Stop
                      Write-Host "Added member: $Member" -ForegroundColor Green
                  }
                  catch {
                      Write-Host "Failed to add member $Member : $_" -ForegroundColor Red
                  }
              }
          }

          Write-Host "Distribution group $DisplayName created successfully." -ForegroundColor Green
      }
      catch {
          Write-Host "Failed to create group $DisplayName : $_" -ForegroundColor Red
      }
  }
}

function Audit {
	<#
	.SYNOPSIS
		Performs a basic audit of a Windows computer.

	.DESCRIPTION
		This script will perform a basic audit of a Windows computer. It will collect information on the computer name, role, operating system, service pack, manufacturer, model, number of processors, memory, registered user, registered organisation, last system boot, hotfix information, logical disk configuration, NIC configuration, software, local shares, printers, services, regional settings, event log settings, event log errors, and event log warnings.

	.PARAMETER auditlist
		The path to a file containing a list of computer names to audit. If no list is specified, the script will audit the local computer.

	.EXAMPLE
		Audit -auditlist "C:\Computers.txt"
		This will audit the computers listed in the file "C:\Computers.txt".

	.LINK
		Utilities

	.NOTES
#####################################################
#				                                    #
#    Original Audit script by Alan Renouf           #
#    Blog: http://virtu-al.net/	                    #
#	     		                                    #
#    Usage: Audit.ps1 'pathtolistofservers'         #
# 			                                        #
#    The file is optional and needs to be a 	    #
#	 plain text list of computers to be audited     #
#	 one on each line, if no list is specified      #
#	 the local machine will be audited.             #
#                                                   #
#####################################################

	#>
	param( [string] $auditlist)

	Function Get-CustomHTML ($Header) {
		$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($Header)</title>
<META http-equiv=Content-Type content='text/html; charset=windows-1252'>

<meta name="save" content="history">

<style type="text/css">
DIV .expando {DISPLAY: block; FONT-WEIGHT: normal; FONT-SIZE: 8pt; RIGHT: 8px; COLOR: #e0e0e0; FONT-FAMILY: Arial; POSITION: absolute; TEXT-DECORATION: underline}
TABLE {TABLE-LAYOUT: fixed; FONT-SIZE: 100%; WIDTH: 100%}
*{margin:0}
body {background-color: #121212; color: #e0e0e0;}
.dspcont { display:none; BORDER-RIGHT: #555555 1px solid; BORDER-TOP: #555555 1px solid; PADDING-LEFT: 16px; FONT-SIZE: 8pt;MARGIN-BOTTOM: -1px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 0px; BORDER-LEFT: #555555 1px solid; WIDTH: 95%; COLOR: #e0e0e0; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #555555 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; BACKGROUND-COLOR: #1e1e1e}
.filler {BORDER-RIGHT: medium none; BORDER-TOP: medium none; DISPLAY: block; BACKGROUND: none transparent scroll repeat 0% 0%; MARGIN-BOTTOM: -1px; FONT: 100%/8px Tahoma; MARGIN-LEFT: 43px; BORDER-LEFT: medium none; COLOR: #e0e0e0; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: medium none; POSITION: relative}
.save{behavior:url(#default#savehistory);}
.dspcont1{ display:none}
a.dsphead0 {BORDER-RIGHT: #555555 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #555555 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #555555 1px solid; CURSOR: hand; COLOR: #FFFFFF; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #555555 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #880000}
a.dsphead1 {BORDER-RIGHT: #555555 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #555555 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #555555 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #555555 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #2D5F87}
a.dsphead2 {BORDER-RIGHT: #555555 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #555555 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #555555 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #555555 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #2D5F87}
a.dsphead1 span.dspchar{font-family:monospace;font-weight:normal;}
td {VERTICAL-ALIGN: TOP; FONT-FAMILY: Tahoma}
th {VERTICAL-ALIGN: TOP; COLOR: #ff6666; TEXT-ALIGN: left}
BODY {margin-left: 4pt} 
BODY {margin-right: 4pt} 
BODY {margin-top: 6pt} 
a {color: #4d9eff;}
a:visited {color: #b388ff;}
</style>


<script type="text/javascript">
function dsp(loc){
   if(document.getElementById){
      var foc=loc.firstChild;
      foc=loc.firstChild.innerHTML?
         loc.firstChild:
         loc.firstChild.nextSibling;
      foc.innerHTML=foc.innerHTML=='hide'?'show':'hide';
      foc=loc.parentNode.nextSibling.style?
         loc.parentNode.nextSibling:
         loc.parentNode.nextSibling.nextSibling;
      foc.style.display=foc.style.display=='block'?'none':'block';}}  

if(!document.getElementById)
   document.write('<style type="text/css">\n'+'.dspcont{display:block;}\n'+ '</style>');
</script>

</head>
<body>
<b><font face="Arial" size="5">$($Header)</font></b><hr size="8" color="#CC0000">
<font face="Arial" size="1"><b>Version 3 by Alan Renouf virtu-al.net</b></font><br>
<font face="Arial" size="1">Report created on $(Get-Date)</font>
<div class="filler"></div>
<div class="filler"></div>
<div class="filler"></div>
<div class="save">
"@
		Return $Report
	}

	Function Get-CustomHeader0 ($Title) {
		$Report = @"
		<h1><a class="dsphead0">$($Title)</a></h1>
	<div class="filler"></div>
"@
		Return $Report
	}

	Function Get-CustomHeader ($Num, $Title) {
		$Report = @"
	<h2><a href="javascript:void(0)" class="dsphead$($Num)" onclick="dsp(this)">
	<span class="expando">show</span>$($Title)</a></h2>
	<div class="dspcont">
"@
		Return $Report
	}

	Function Get-CustomHeaderClose {

		$Report = @"
		</DIV>
		<div class="filler"></div>
"@
		Return $Report
	}

	Function Get-CustomHeader0Close {

		$Report = @"
</DIV>
"@
		Return $Report
	}

	Function Get-CustomHTMLClose {

		$Report = @"
</div>

</body>
</html>
"@
		Return $Report
	}

	Function Get-HTMLTable {
		param([array]$Content)
		$HTMLTable = $Content | ConvertTo-Html
		$HTMLTable = $HTMLTable -replace '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">', ""
		$HTMLTable = $HTMLTable -replace '<html xmlns="http://www.w3.org/1999/xhtml">', ""
		$HTMLTable = $HTMLTable -replace '<head>', ""
		$HTMLTable = $HTMLTable -replace '<title>HTML TABLE</title>', ""
		$HTMLTable = $HTMLTable -replace '</head><body>', ""
		$HTMLTable = $HTMLTable -replace '</body></html>', ""
		Return $HTMLTable
	}

	Function Get-HTMLDetail ($Heading, $Detail) {
		$Report = @"
<TABLE>
	<tr>
	<th width='25%'><b>$Heading</b></font></th>
	<td width='75%'>$($Detail)</td>
	</tr>
</TABLE>
"@
		Return $Report
	}

	if ($auditlist -eq "") {
		Write-Host "No list specified, using $env:computername"
		$targets = $env:computername
	}
	else {
		if ((Test-Path $auditlist) -eq $false) {
			Write-Host "Invalid audit path specified: $auditlist"
			exit
		}
		else {
			Write-Host "Using Audit list: $auditlist"
			$Targets = Get-Content $auditlist
		}
	}

	Foreach ($Target in $Targets) {

		Write-Output "Collating Detail for $Target"
		$ComputerSystem = Get-WmiObject -computername $Target Win32_ComputerSystem
		switch ($ComputerSystem.DomainRole) {
			0 { $ComputerRole = "Standalone Workstation" }
			1 { $ComputerRole = "Member Workstation" }
			2 { $ComputerRole = "Standalone Server" }
			3 { $ComputerRole = "Member Server" }
			4 { $ComputerRole = "Domain Controller" }
			5 { $ComputerRole = "Domain Controller" }
			default { $ComputerRole = "Information not available" }
		}
	
		$OperatingSystems = Get-WmiObject -computername $Target Win32_OperatingSystem
		$TimeZone = Get-WmiObject -computername $Target Win32_Timezone
		$Keyboards = Get-WmiObject -computername $Target Win32_Keyboard
		$SchedTasks = Get-WmiObject -computername $Target Win32_ScheduledJob
		$BootINI = $OperatingSystems.SystemDrive + "boot.ini"
		$RecoveryOptions = Get-WmiObject -computername $Target Win32_OSRecoveryConfiguration
	
		switch ($ComputerRole) {
			"Member Workstation" { $CompType = "Computer Domain"; break }
			"Domain Controller" { $CompType = "Computer Domain"; break }
			"Member Server" { $CompType = "Computer Domain"; break }
			default { $CompType = "Computer Workgroup"; break }
		}

		$LBTime = $OperatingSystems.ConvertToDateTime($OperatingSystems.Lastbootuptime)
		Write-Output "..Regional Options"
		$ObjKeyboards = Get-WmiObject -ComputerName $Target Win32_Keyboard
		$keyboardmap = @{
			"00000402" = "BG" 
			"00000404" = "CH" 
			"00000405" = "CZ" 
			"00000406" = "DK" 
			"00000407" = "GR" 
			"00000408" = "GK" 
			"00000409" = "US" 
			"0000040A" = "SP" 
			"0000040B" = "SU" 
			"0000040C" = "FR" 
			"0000040E" = "HU" 
			"0000040F" = "IS" 
			"00000410" = "IT" 
			"00000411" = "JP" 
			"00000412" = "KO" 
			"00000413" = "NL" 
			"00000414" = "NO" 
			"00000415" = "PL" 
			"00000416" = "BR" 
			"00000418" = "RO" 
			"00000419" = "RU" 
			"0000041A" = "YU" 
			"0000041B" = "SL" 
			"0000041C" = "US" 
			"0000041D" = "SV" 
			"0000041F" = "TR" 
			"00000422" = "US" 
			"00000423" = "US" 
			"00000424" = "YU" 
			"00000425" = "ET" 
			"00000426" = "US" 
			"00000427" = "US" 
			"00000804" = "CH" 
			"00000809" = "UK" 
			"0000080A" = "LA" 
			"0000080C" = "BE" 
			"00000813" = "BE" 
			"00000816" = "PO" 
			"00000C0C" = "CF" 
			"00000C1A" = "US" 
			"00001009" = "US" 
			"0000100C" = "SF" 
			"00001809" = "US" 
			"00010402" = "US" 
			"00010405" = "CZ" 
			"00010407" = "GR" 
			"00010408" = "GK" 
			"00010409" = "DV" 
			"0001040A" = "SP" 
			"0001040E" = "HU" 
			"00010410" = "IT" 
			"00010415" = "PL" 
			"00010419" = "RU" 
			"0001041B" = "SL" 
			"0001041F" = "TR" 
			"00010426" = "US" 
			"00010C0C" = "CF" 
			"00010C1A" = "US" 
			"00020408" = "GK" 
			"00020409" = "US" 
			"00030409" = "USL" 
			"00040409" = "USR" 
			"00050408" = "GK" 
		}
		$keyb = $keyboardmap.$($ObjKeyboards.Layout)
		if (!$keyb) {
			$keyb = "Unknown"
		}
		$MyReport = Get-CustomHTML "$Target Audit"
		$MyReport += Get-CustomHeader0  "$Target Details"
		$MyReport += Get-CustomHeader "2" "General"
		$MyReport += Get-HTMLDetail "Computer Name" ($ComputerSystem.Name)
		$MyReport += Get-HTMLDetail "Computer Role" ($ComputerRole)
		$MyReport += Get-HTMLDetail $CompType ($ComputerSystem.Domain)
		$MyReport += Get-HTMLDetail "Operating System" ($OperatingSystems.Caption)
		$MyReport += Get-HTMLDetail "Service Pack" ($OperatingSystems.CSDVersion)
		$MyReport += Get-HTMLDetail "System Root" ($OperatingSystems.SystemDrive)
		$MyReport += Get-HTMLDetail "Manufacturer" ($ComputerSystem.Manufacturer)
		$MyReport += Get-HTMLDetail "Model" ($ComputerSystem.Model)
		$MyReport += Get-HTMLDetail "Number of Processors" ($ComputerSystem.NumberOfProcessors)
		$MyReport += Get-HTMLDetail "Memory" ($ComputerSystem.TotalPhysicalMemory)
		$MyReport += Get-HTMLDetail "Registered User" ($ComputerSystem.PrimaryOwnerName)
		$MyReport += Get-HTMLDetail "Registered Organisation" ($OperatingSystems.Organization)
		$MyReport += Get-HTMLDetail "Last System Boot" ($LBTime)
		$MyReport += Get-CustomHeaderClose
		Write-Output "..Hotfix Information"
		$colQuickFixes = Get-WmiObject Win32_QuickFixEngineering
		$MyReport += Get-CustomHeader "2" "HotFixes"
		$MyReport += Get-HTMLTable ($colQuickFixes | Where { $_.HotFixID -ne "File 1" } | Select HotFixID, Description)
		$MyReport += Get-CustomHeaderClose
		Write-Output "..Logical Disks"
		$Disks = Get-WmiObject -ComputerName $Target Win32_LogicalDisk
		$MyReport += Get-CustomHeader "2" "Logical Disk Configuration"
		$LogicalDrives = @()
		Foreach ($LDrive in ($Disks | Where { $_.DriveType -eq 3 })) {
			$Details = "" | Select "Drive Letter", Label, "File System", "Disk Size (MB)", "Disk Free Space", "% Free Space"
			$Details."Drive Letter" = $LDrive.DeviceID
			$Details.Label = $LDrive.VolumeName
			$Details."File System" = $LDrive.FileSystem
			$Details."Disk Size (MB)" = [math]::round(($LDrive.size / 1MB))
			$Details."Disk Free Space" = [math]::round(($LDrive.FreeSpace / 1MB))
			$Details."% Free Space" = [Math]::Round(($LDrive.FreeSpace / 1MB) / ($LDrive.Size / 1MB) * 100)
			$LogicalDrives += $Details
		}
		$MyReport += Get-HTMLTable ($LogicalDrives)
		$MyReport += Get-CustomHeaderClose
		Write-Output "..Network Configuration"
		$Adapters = Get-WmiObject -ComputerName $Target Win32_NetworkAdapterConfiguration
		$MyReport += Get-CustomHeader "2" "NIC Configuration"
		$IPInfo = @()
		Foreach ($Adapter in ($Adapters | Where { $_.IPEnabled -eq $True })) {
			$Details = "" | Select Description, "Physical address", "IP Address / Subnet Mask", "Default Gateway", "DHCP Enabled", DNS, WINS
			$Details.Description = "$($Adapter.Description)"
			$Details."Physical address" = "$($Adapter.MACaddress)"
			If ($Adapter.IPAddress -ne $Null) {
				$Details."IP Address / Subnet Mask" = "$($Adapter.IPAddress)/$($Adapter.IPSubnet)"
				$Details."Default Gateway" = "$($Adapter.DefaultIPGateway)"
			}
			If ($Adapter.DHCPEnabled -eq "True")	{
				$Details."DHCP Enabled" = "Yes"
			}
			Else {
				$Details."DHCP Enabled" = "No"
			}
			If ($Adapter.DNSServerSearchOrder -ne $Null)	{
				$Details.DNS = "$($Adapter.DNSServerSearchOrder)"
			}
			$Details.WINS = "$($Adapter.WINSPrimaryServer) $($Adapter.WINSSecondaryServer)"
			$IPInfo += $Details
		}
		$MyReport += Get-HTMLTable ($IPInfo)
		$MyReport += Get-CustomHeaderClose
		If ((get-wmiobject -ComputerName $Target -namespace "root/cimv2" -list) | Where-Object { $_.name -match "Win32_Product" }) {
			Write-Output "..Software"
			$MyReport += Get-CustomHeader "2" "Software"
			$MyReport += Get-HTMLTable (get-wmiobject -ComputerName $Target Win32_Product | select Name, Version, Vendor, InstallDate)
			$MyReport += Get-CustomHeaderClose
		}
		Else {
			Write-Output "..Software WMI class not installed"
		}
		Write-Output "..Local Shares"
		$Shares = Get-wmiobject -ComputerName $Target Win32_Share
		$MyReport += Get-CustomHeader "2" "Local Shares"
		$MyReport += Get-HTMLTable ($Shares | Select Name, Path, Caption)
		$MyReport += Get-CustomHeaderClose
		Write-Output "..Printers"
		$InstalledPrinters = Get-WmiObject -ComputerName $Target Win32_Printer
		$MyReport += Get-CustomHeader "2" "Printers"
		$MyReport += Get-HTMLTable ($InstalledPrinters | Select Name, Location)
		$MyReport += Get-CustomHeaderClose
		Write-Output "..Services"
		$ListOfServices = Get-WmiObject -ComputerName $Target Win32_Service
		$MyReport += Get-CustomHeader "2" "Services"
		$Services = @()
		Foreach ($Service in $ListOfServices) {
			$Details = "" | Select Name, Account, "Start Mode", State, "Expected State"
			$Details.Name = $Service.Caption
			$Details.Account = $Service.Startname
			$Details."Start Mode" = $Service.StartMode
			If ($Service.StartMode -eq "Auto") {
				if ($Service.State -eq "Stopped") {
					$Details.State = $Service.State
					$Details."Expected State" = "Unexpected"
				}
			}
			If ($Service.StartMode -eq "Auto") {
				if ($Service.State -eq "Running") {
					$Details.State = $Service.State
					$Details."Expected State" = "OK"
				}
			}
			If ($Service.StartMode -eq "Disabled") {
				If ($Service.State -eq "Running") {
					$Details.State = $Service.State
					$Details."Expected State" = "Unexpected"
				}
			}
			If ($Service.StartMode -eq "Disabled") {
				if ($Service.State -eq "Stopped") {
					$Details.State = $Service.State
					$Details."Expected State" = "OK"
				}
			}
			If ($Service.StartMode -eq "Manual") {
				$Details.State = $Service.State
				$Details."Expected State" = "OK"
			}
			If ($Service.State -eq "Paused") {
				$Details.State = $Service.State
				$Details."Expected State" = "OK"
			}
			$Services += $Details
		}
		$MyReport += Get-HTMLTable ($Services)
		$MyReport += Get-CustomHeaderClose
		$MyReport += Get-CustomHeader "2" "Regional Settings"
		$MyReport += Get-HTMLDetail "Time Zone" ($TimeZone.Description)
		$MyReport += Get-HTMLDetail "Country Code" ($OperatingSystems.Countrycode)
		$MyReport += Get-HTMLDetail "Locale" ($OperatingSystems.Locale)
		$MyReport += Get-HTMLDetail "Operating System Language" ($OperatingSystems.OSLanguage)
		$MyReport += Get-HTMLDetail "Keyboard Layout" ($keyb)
		$MyReport += Get-CustomHeaderClose
		Write-Output "..Event Log Settings"
		$LogFiles = Get-WmiObject -ComputerName $Target Win32_NTEventLogFile
		$MyReport += Get-CustomHeader "2" "Event Logs"
		$MyReport += Get-CustomHeader "2" "Event Log Settings"
		$LogSettings = @()
		Foreach ($Log in $LogFiles) {
			$Details = "" | Select "Log Name", "Overwrite Outdated Records", "Maximum Size (KB)", "Current Size (KB)"
			$Details."Log Name" = $Log.LogFileName
			If ($Log.OverWriteOutdated -lt 0) {
				$Details."Overwrite Outdated Records" = "Never"
			}
			if ($Log.OverWriteOutdated -eq 0) {
				$Details."Overwrite Outdated Records" = "As needed"
			}
			Else {
				$Details."Overwrite Outdated Records" = "After $($Log.OverWriteOutdated) days"
			}
			$MaxFileSize = ($Log.MaxFileSize) / 1024
			$FileSize = ($Log.FileSize) / 1024
				
			$Details."Maximum Size (KB)" = $MaxFileSize
			$Details."Current Size (KB)" = $FileSize
			$LogSettings += $Details
		}
		$MyReport += Get-HTMLTable ($LogSettings)
		$MyReport += Get-CustomHeaderClose
		Write-Output "..Event Log Errors"
		$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::Now.AddDays(-14))
		$LoggedErrors = Get-WmiObject -computer $Target -query ("Select * from Win32_NTLogEvent Where Type='Error' and TimeWritten >='" + $WmidtQueryDT + "'")
		$MyReport += Get-CustomHeader "2" "ERROR Entries"
		$MyReport += Get-HTMLTable ($LoggedErrors | Select EventCode, SourceName, @{N = "Time"; E = { $_.ConvertToDateTime($_.TimeWritten) } }, LogFile, Message)
		$MyReport += Get-CustomHeaderClose
		Write-Output "..Event Log Warnings"
		$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::Now.AddDays(-14))
		$LoggedWarning = Get-WmiObject -computer $Target -query ("Select * from Win32_NTLogEvent Where Type='Warning' and TimeWritten >='" + $WmidtQueryDT + "'")
		$MyReport += Get-CustomHeader "2" "WARNING Entries"
		$MyReport += Get-HTMLTable ($LoggedWarning | Select EventCode, SourceName, @{N = "Time"; E = { $_.ConvertToDateTime($_.TimeWritten) } }, LogFile, Message)
		$MyReport += Get-CustomHeaderClose
		$MyReport += Get-CustomHeaderClose
		$MyReport += Get-CustomHeaderClose
		$MyReport += Get-CustomHeader0Close
		$MyReport += Get-CustomHTMLClose
		$MyReport += Get-CustomHTMLClose

		$Date = Get-Date
		$Filename = ".\" + $Target + "_" + $date.Hour + $date.Minute + "_" + $Date.Day + "-" + $Date.Month + "-" + $Date.Year + ".htm"
		$MyReport | out-file -encoding ASCII -filepath $Filename
		Write "Audit saved as $Filename"
	}
}

function Get-WinISO {
  <#
  .SYNOPSIS
  Downloads Microsoft 365 installation files based on specified parameters.
  
  .DESCRIPTION
  The Get-WinISO function retrieves and downloads Microsoft 365 installation files
  from official sources. It supports various options for Windows versions, release types, editions,
  languages, and architectures.
  
  .PARAMETER Win
  Specifies the Windows version. Valid values are 10, 11, or All.
  
  .PARAMETER Rel
  Specifies the release type. Valid values are Latest, Insider, or Dev.
  
  .PARAMETER Ed
  Specifies the edition. Valid values are Home, Pro, Edu, or All.
  
  .PARAMETER Lang
  Specifies the language for the installation files.
  
  .PARAMETER Arch
  Specifies the architecture. Valid values are x86, x64, arm64, or All.
  
  .EXAMPLE
  Get-WinISO -Win 11 -Rel Latest -Ed Pro -Lang English -Arch x64
  
  Downloads the latest Windows 11 Pro 64-bit installation files in English.
  
  .LINK
  Microsoft 365
  #>
  param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('10', '11', 'All')]
    [string]$Win = '11',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('Latest', 'Insider', 'Dev')]
    [string]$Rel = 'Latest',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('Home', 'Pro', 'Edu', 'All')]
    [string]$Ed = 'Pro',
    
    [Parameter(Mandatory = $false)]
    [string]$Lang = 'English',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('x86', 'x64', 'arm64', 'All')]
    [string]$Arch = 'x64'
  )

  # Construct the command invocation
  $command = "(irm https://raw.githubusercontent.com/roberto-ryan/Fido/refs/heads/master/Fido.ps1) | iex ; Download"
  
  # Add parameters if they're specified
  if ($Win -ne 'All') { $command += " -Win $Win" }
  if ($Rel -ne 'Latest') { $command += " -Rel $Rel" }
  if ($Ed -ne 'All') { $command += " -Ed $Ed" }
  if ($Lang) { $command += " -Lang $Lang" }
  if ($Arch -ne 'All') { $command += " -Arch $Arch" }
  
  # Execute the command
  Write-Host "Executing: $command" -ForegroundColor Cyan
  Invoke-Expression $command
}

function Manage-vtsADUsers {
  <#
  .SYNOPSIS
      Processes a list of Active Directory group memberships by comparing a target list with a source list.
  
  .DESCRIPTION
      This function takes a target group membership list and a source group membership list,
      then identifies items that are in both lists, only in the target list, or only in the source list.
      It returns a custom object containing these categorized memberships.
  
  .PARAMETER targetMembership
      The collection of group memberships considered as the target for comparison.
  
  .PARAMETER sourceMembership
      The collection of group memberships considered as the source for comparison.
  
  .OUTPUTS
      [PSCustomObject] with the following properties:
      - InBoth: Memberships that exist in both target and source lists
      - OnlyInTarget: Memberships that exist only in the target list
      - OnlyInSource: Memberships that exist only in the source list
  
  .EXAMPLE
      $targetMembers = Get-ADGroupMember -Identity "TargetGroup"
      $sourceMembers = Get-ADGroupMember -Identity "SourceGroup"
      $result = Process-GroupMembership -targetMembership $targetMembers -sourceMembership $sourceMembers
      $result.InBoth | ForEach-Object { Write-Host "Member in both groups: $($_)" }
  
  .NOTES
      This function uses Compare-Object to efficiently identify differences between the two lists.
  
  .LINK
      Active Directory
      #>
  # Load required assemblies for Windows Forms
  Add-Type -AssemblyName System.Windows.Forms, System.Drawing
  
  # Setup logging function
  function Write-Log {
      param (
          [string]$Message,
          [string]$Level = "INFO"
      )
      $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
      $logMessage = "[$timestamp] [$Level] $Message"
      Write-Host $logMessage
      Add-Content -Path "$env:TEMP\ADUserManager.log" -Value $logMessage
  }
  
  Write-Host "`n========== AD User Manager ==========`n" -ForegroundColor Cyan
  Write-Host "This tool helps you manage Active Directory users." -ForegroundColor Cyan
  Write-Host "First, you'll select one or more OUs to work with." -ForegroundColor Cyan
  Write-Host "Then you can view and manage users in those OUs." -ForegroundColor Cyan
  Write-Host "======================================`n" -ForegroundColor Cyan
  
  Write-Log "Application started"
  
  # Fetch OUs and display in Out-GridView for selection
  try {
      Write-Host "Fetching organizational units from Active Directory..." -ForegroundColor Yellow
      $OUs = Get-ADOrganizationalUnit -Filter * -Properties Name,DistinguishedName | 
             Select-Object Name,DistinguishedName
      
      Write-Host "Select one or more OUs in the grid view window and click OK." -ForegroundColor Green
      $selectedOUs = $OUs | Out-GridView -Title "Select Organizational Units (Select multiple and click OK)" -OutputMode Multiple
      
      if ($null -eq $selectedOUs -or $selectedOUs.Count -eq 0) {
          Write-Host "No OUs selected. Exiting." -ForegroundColor Red
          Write-Log "No OUs selected, application terminated" -Level "WARN"
          exit
      }
      
      Write-Log "Selected $($selectedOUs.Count) OUs"
  } catch {
      Write-Host "Error fetching OUs: $_" -ForegroundColor Red
      Write-Log "Error fetching OUs: $_" -Level "ERROR"
      exit
  }
  
  # Fetch users from selected OUs
  try {
      Write-Host "Fetching users from selected OUs..." -ForegroundColor Yellow
      $users = @()
      foreach ($OU in $selectedOUs) {
          Write-Host "  Loading users from OU: $($OU.Name)" -ForegroundColor Gray
          # Get users with properties that help determine activity status
          $OUusers = Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Properties Name,SamAccountName,Enabled,LastLogonDate,PasswordLastSet,PasswordExpired,LockedOut,Description | 
                     Select-Object Name,SamAccountName,Enabled,LastLogonDate,PasswordLastSet,PasswordExpired,LockedOut,Description,DistinguishedName
          $users += $OUusers
      }
      
      if ($users.Count -eq 0) {
          Write-Host "No users found in selected OUs. Exiting." -ForegroundColor Red
          Write-Log "No users found in selected OUs" -Level "WARN"
          exit
      }
      
      Write-Log "Fetched $($users.Count) users from selected OUs"
  } catch {
      Write-Host "Error fetching users: $_" -ForegroundColor Red
      Write-Log "Error fetching users: $_" -Level "ERROR"
      exit
  }
  
  # Create the Windows Form
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "AD User Manager"
  $form.Size = [System.Drawing.Size]::new(900, 600)
  $form.StartPosition = "CenterScreen"
  
  # Create the DataGridView to display users
  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.Dock = "Fill"
  $grid.AutoSizeColumnsMode = "Fill"
  $grid.SelectionMode = "FullRowSelect"
  $grid.MultiSelect = $false
  $grid.AllowUserToAddRows = $false
  $grid.AllowUserToDeleteRows = $false
  $grid.ReadOnly = $true
  $form.Controls.Add($grid)
  
  # Create button panel
  $buttonPanel = New-Object System.Windows.Forms.Panel
  $buttonPanel.Dock = "Bottom"
  $buttonPanel.Height = 50
  $form.Controls.Add($buttonPanel)
  
  # Create the Refresh button
  $refreshButton = New-Object System.Windows.Forms.Button
  $refreshButton.Text = "Refresh Users"
  $refreshButton.Location = [System.Drawing.Point]::new(10, 10)
  $refreshButton.Size = [System.Drawing.Size]::new(120, 30)
  $buttonPanel.Controls.Add($refreshButton)
  
  # Create Enable User button
  $enableButton = New-Object System.Windows.Forms.Button
  $enableButton.Text = "Enable User"
  $enableButton.Location = [System.Drawing.Point]::new(140, 10)
  $enableButton.Size = [System.Drawing.Size]::new(120, 30)
  $buttonPanel.Controls.Add($enableButton)
  
  # Create Disable User button
  $disableButton = New-Object System.Windows.Forms.Button
  $disableButton.Text = "Disable User"
  $disableButton.Location = [System.Drawing.Point]::new(270, 10)
  $disableButton.Size = [System.Drawing.Size]::new(120, 30)
  $buttonPanel.Controls.Add($disableButton)
  
  # Create status bar for messages
  $statusBar = New-Object System.Windows.Forms.StatusStrip
  $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
  $statusLabel.Text = "Ready"
  $statusBar.Items.Add($statusLabel)
  $form.Controls.Add($statusBar)
  
  # Function to populate the grid with data
  function Populate-Grid {
      param ($data)
      
      # Create a DataTable for better compatibility with DataGridView
      $dataTable = New-Object System.Data.DataTable
      
      # Add columns to the DataTable
      "Name", "SamAccountName", "Enabled", "LastLogonDate", "PasswordLastSet", 
      "PasswordExpired", "LockedOut", "Description" | ForEach-Object {
          $dataTable.Columns.Add($_) | Out-Null
      }
      
      # Add rows to the DataTable
      foreach ($user in ($data | Sort-Object Enabled -Descending)) {
          $row = $dataTable.NewRow()
          $row["Name"] = $user.Name
          $row["SamAccountName"] = $user.SamAccountName
          $row["Enabled"] = $user.Enabled
          $row["LastLogonDate"] = $user.LastLogonDate
          $row["PasswordLastSet"] = $user.PasswordLastSet
          $row["PasswordExpired"] = $user.PasswordExpired
          $row["LockedOut"] = $user.LockedOut
          $row["Description"] = $user.Description
          $dataTable.Rows.Add($row)
      }
      
      # Set the DataSource to the DataTable
      $grid.DataSource = $dataTable
      
      Write-Log "Grid populated with $($data.Count) users"
  }
  
  # Function to refresh data from Active Directory
  function Refresh-Users {
      try {
          Write-Host "Refreshing user data..." -ForegroundColor Yellow
          $refreshedUsers = @()
          
          foreach ($OU in $selectedOUs) {
              Write-Host "  Refreshing users from OU: $($OU.Name)" -ForegroundColor Gray
              $OUusers = Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Properties Name,SamAccountName,Enabled,LastLogonDate,PasswordLastSet,PasswordExpired,LockedOut,Description | 
                         Select-Object Name,SamAccountName,Enabled,LastLogonDate,PasswordLastSet,PasswordExpired,LockedOut,Description,DistinguishedName
              $refreshedUsers += $OUusers
          }
          
          Populate-Grid -data $refreshedUsers
          $statusLabel.Text = "Users refreshed at $(Get-Date -Format 'HH:mm:ss')"
          Write-Log "Users refreshed" -Level "INFO"
      } catch {
          $statusLabel.Text = "Error refreshing users"
          [System.Windows.Forms.MessageBox]::Show("Error refreshing users: $_", "Error", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
          Write-Log "Error refreshing users: $_" -Level "ERROR"
      }
  }
  
  # Function to enable selected user
  function Enable-SelectedUser {
      if ($grid.SelectedRows.Count -eq 0) {
          $statusLabel.Text = "No user selected"
          [System.Windows.Forms.MessageBox]::Show("Please select a user to enable.", "No Selection", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
          return
      }
      
      $selectedRow = $grid.SelectedRows[0]
      $userName = $selectedRow.Cells["SamAccountName"].Value
      $userEnabled = $selectedRow.Cells["Enabled"].Value
      
      # Convert string representation to boolean if needed
      if ($userEnabled -is [string]) {
          $userEnabled = [System.Boolean]::Parse($userEnabled)
      }
      
      if ($userEnabled -eq $true) {
          $statusLabel.Text = "User already enabled"
          [System.Windows.Forms.MessageBox]::Show("$userName is already enabled.", "Info", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
          return
      }
      
      try {
          $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to enable $userName?", "Confirm", 
              [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
          
          if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
              Enable-ADAccount -Identity $userName
              Write-Log "User $userName enabled" -Level "INFO"
              $statusLabel.Text = "User $userName enabled successfully"
              [System.Windows.Forms.MessageBox]::Show("$userName has been enabled.", "Success", 
                  [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
              Refresh-Users
          }
      } catch {
          $statusLabel.Text = "Error enabling user"
          [System.Windows.Forms.MessageBox]::Show("Error enabling $userName`: $_", "Error", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
          Write-Log "Error enabling user $userName`: $_" -Level "ERROR"
      }
  }
  
  # Function to disable selected user
  function Disable-SelectedUser {
      if ($grid.SelectedRows.Count -eq 0) {
          $statusLabel.Text = "No user selected"
          [System.Windows.Forms.MessageBox]::Show("Please select a user to disable.", "No Selection", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
          return
      }
      
      $selectedRow = $grid.SelectedRows[0]
      $userName = $selectedRow.Cells["SamAccountName"].Value
      $userEnabled = $selectedRow.Cells["Enabled"].Value
      
      # Convert string representation to boolean if needed
      if ($userEnabled -is [string]) {
          $userEnabled = [System.Boolean]::Parse($userEnabled)
      }
      
      if ($userEnabled -eq $false) {
          $statusLabel.Text = "User already disabled"
          [System.Windows.Forms.MessageBox]::Show("$userName is already disabled.", "Info", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
          return
      }
      
      try {
          $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to disable $userName?", "Confirm", 
              [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
          
          if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
              Disable-ADAccount -Identity $userName
              Write-Log "User $userName disabled" -Level "INFO"
              $statusLabel.Text = "User $userName disabled successfully"
              [System.Windows.Forms.MessageBox]::Show("$userName has been disabled.", "Success", 
                  [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
              Refresh-Users
          }
      } catch {
          $statusLabel.Text = "Error disabling user"
          [System.Windows.Forms.MessageBox]::Show("Error disabling $userName`: $_", "Error", 
              [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
          Write-Log "Error disabling user $userName`: $_" -Level "ERROR"
      }
  }
  
  # Attach event handlers
  $refreshButton.Add_Click({ Refresh-Users })
  $enableButton.Add_Click({ Enable-SelectedUser })
  $disableButton.Add_Click({ Disable-SelectedUser })
  
  # Populate the grid with initial data
  Populate-Grid -data $users
  
  Write-Host "`nUser management window opened. You can now:" -ForegroundColor Cyan
  Write-Host "  - View user details in the grid" -ForegroundColor White
  Write-Host "  - Select a user and click 'Enable User' or 'Disable User'" -ForegroundColor White
  Write-Host "  - Click 'Refresh Users' to update the list" -ForegroundColor White
  Write-Host "Logs are being saved to: $env:TEMP\ADUserManager.log" -ForegroundColor White
  Write-Host "`nUse the grid to sort by any column by clicking the column header." -ForegroundColor Cyan
  
  # Display the form
  $form.ShowDialog()
  
  Write-Log "Application closed"
}

function Get-vts365PowerBIUsageReport {
  <#
  .SYNOPSIS
      Performs various operations on a folder structure based on specified configuration settings.

  .DESCRIPTION
      This script processes folders according to the configuration specified in a JSON file.
      It can:
      - Move folders to appropriate destinations based on their names
      - Rename folders according to defined patterns
      - Archive folders that match specified criteria
      - Handle nested folder structures
      
      The script reads its configuration from a JSON file that defines rules for processing
      folders, including pattern matching, destination paths, and archiving settings.

  .PARAMETER ConfigFile
      Path to the JSON configuration file that defines processing rules.

  .PARAMETER RootFolder
      Path to the root folder where processing should begin.

  .PARAMETER ArchiveFolder
      Path to the folder where archived items should be stored.

  .PARAMETER LogFile
      Path to the log file where processing activities will be recorded.

  .PARAMETER Force
      If specified, forces the script to execute without confirmation prompts.

  .EXAMPLE
      .\YourScript.ps1 -ConfigFile "config.json" -RootFolder "C:\Data" -ArchiveFolder "C:\Archive" -LogFile "C:\Logs\process.log"
      
      Processes folders in C:\Data according to rules in config.json, archives to C:\Archive, and logs to the specified file.

  .EXAMPLE
      .\YourScript.ps1 -ConfigFile "config.json" -RootFolder "C:\Data" -Force
      
      Processes folders in C:\Data using the default archive location and logging, without any confirmation prompts.

  .NOTES
      File Name      : YourScript.ps1
      Author         : Your Name
      Prerequisite   : PowerShell 5.1 or later
      Copyright      : Your Copyright Information
      
  .LINK
      M365
  #>

  # Check if the required module is installed, if not, install it
  if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
      Write-Host "MicrosoftPowerBIMgmt module not found. Installing..."
      Install-Module -Name MicrosoftPowerBIMgmt -Force -AllowClobber
  } else {
      Write-Host "MicrosoftPowerBIMgmt module is already installed."
  }

  Connect-PowerBIServiceAccount

  # Define the number of days to retrieve (max 30)
  $daysToExtract = 30

  # Initialize an array to store all activities
  $allActivities = @()

  # Loop through each day
  for ($i = 0; $i -lt $daysToExtract; $i++) {
      # Calculate start and end times for the current day
      $startDate = (Get-Date).AddDays(-$i).ToString("yyyy-MM-ddT00:00:00Z")
      $endDate = (Get-Date).AddDays(-$i).ToString("yyyy-MM-ddT23:59:59Z")
  
      Write-Host "Fetching activities for $startDate to $endDate"

      # Fetch activities for this day
      $activities = Get-PowerBIActivityEvent -StartDateTime $startDate -EndDateTime $endDate -Verbose

      # Parse JSON response (if returned as JSON)
      if ($activities) {
          $activityData = $activities | ConvertFrom-Json
          if ($activityData) {
              $allActivities += $activityData
          }
      }
  }

  # Export to CSV if there’s data
  if ($allActivities.Count -gt 0) {
      $allActivities | Select-Object -Property CreationTime, UserId, Activity, Operation | Export-Csv -Path "C:\Reports\PowerBIActivity_30Days.csv" -NoTypeInformation
      Write-Host "Data exported to C:\Reports\PowerBIActivity_30Days.csv"
  }
  else {
      Write-Host "No activity data found for the specified 30-day period."
  }

}

function Search-vtsTextInFiles {
<#
.SYNOPSIS
Search recursively through a directory for text matches in probable text files.

.DESCRIPTION
Search-TextInFiles scans all files under the specified path, evaluating whether 
each file is likely to be a text file before reading it. It performs a 
line-by-line search for the specified string using either case-sensitive or 
case-insensitive comparison. For each file that contains matches, it outputs 
the file path and the line numbers where matches were found.

.PARAMETER Path
The root directory to begin searching from.

.PARAMETER SearchTerm
The text string to search for within each file.

.PARAMETER CaseSensitive
Switch to enable case-sensitive matching. By default the search is case-insensitive.

.EXAMPLE
Search-TextInFiles -Path "C:\Logs" -SearchTerm "error"
Searches all text files under C:\Logs (case-insensitive) and prints matches.

.EXAMPLE
Search-TextInFiles -Path . -SearchTerm "TokenExpired" -CaseSensitive
Performs case-sensitive search for "TokenExpired" in the current directory.

.NOTES
This function uses .NET APIs for efficient directory traversal and file access.
It can read files that are locked for writing by other processes.
#>

    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [switch]$CaseSensitive
    )

    # Validate path
    if (-not (Test-Path -Path $Path)) {
        Write-Error "Path not found: $Path"
        exit 2
    }

    $comparison = if ($CaseSensitive.IsPresent) {
        [System.StringComparison]::Ordinal
    } else {
        [System.StringComparison]::OrdinalIgnoreCase
    }

    function Is-ProbablyTextFile {
        param(
            [string]$FilePath,
            [int]$SampleBytes = 8192
        )

        try {
            $fs = [System.IO.File]::Open(
                $FilePath,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite
            )
        } catch {
            return $false
        }

        try {
            $buffer = New-Object byte[] ([Math]::Min($SampleBytes, [int]$fs.Length))
            if ($buffer.Length -le 0) { $fs.Close(); return $false }

            $bytesRead = $fs.Read($buffer, 0, $buffer.Length)

            for ($i = 0; $i -lt $bytesRead; $i++) {
                if ($buffer[$i] -eq 0) {
                    $fs.Close()
                    return $false
                }
            }

            $controlCount = 0
            for ($i = 0; $i -lt $bytesRead; $i++) {
                $b = $buffer[$i]
                if ($b -lt 32 -and $b -ne 9 -and $b -ne 10 -and $b -ne 13) { 
                    $controlCount++ 
                }
            }

            $fs.Close()
            if ($controlCount -gt ($bytesRead * 0.3)) { return $false }

            return $true
        } catch {
            try { $fs.Close() } catch {}
            return $false
        }
    }

    # Directory traversal using .NET
    $dirStack = New-Object System.Collections.Generic.Stack[string]
    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $dirStack.Push($resolved)

    while ($dirStack.Count -gt 0) {
        $currentDir = $dirStack.Pop()

        try {
            $files = [System.IO.Directory]::GetFiles($currentDir)
        } catch {
            continue
        }

        foreach ($file in $files) {
            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }

            if (-not (Is-ProbablyTextFile -FilePath $file)) { continue }

            try {
                $fs = [System.IO.File]::Open(
                    $file,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite
                )
                $sr = New-Object System.IO.StreamReader($fs, $true)
            } catch {
                try { if ($fs) { $fs.Close() } } catch {}
                continue
            }

            try {
                $lineNumber = 0
                $matches = New-Object System.Collections.Generic.List[int]

                while (($line = $sr.ReadLine()) -ne $null) {
                    $lineNumber++
                    if ($line.IndexOf($SearchTerm, $comparison) -ge 0) {
                        $matches.Add($lineNumber)
                    }
                }

                if ($matches.Count -gt 0) {
                    $lines = ($matches | ForEach-Object { $_ }) -join ','
                    [Console]::WriteLine("{0}`tLines: {1}", $file, $lines)
                }
            } catch {
            } finally {
                try { $sr.Close() } catch {}
                try { $fs.Close() } catch {}
            }
        }

        try {
            $subdirs = [System.IO.Directory]::GetDirectories($currentDir)
            foreach ($d in $subdirs) {
                $dirStack.Push($d)
            }
        } catch {}
    }
}


