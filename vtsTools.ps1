<#
.Description
Searches the last 500 System and Application logs for a specified search term.
.EXAMPLE
PS> Search-vtsEventLog <search term>
.EXAMPLE
PS> Search-vtsEventLog driver

Output:
TimeGenerated : 9/13/2022 9:14:30 AM
Message       : Media disconnected on NIC /DEVICE/{90E7B0EA-AE78-4836-8CBC-B73F1BCD5894} (Friendly Name: Microsoft
                Network Adapter Multiplexor Driver).
Log           : System
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
Retrieves Mapped Drives from the Windows Registry.
.EXAMPLE
PS> Get-vtsMappedDrive

Output:
Username            : VTS-ROBERTO\rober
DriveLetter         : Y
RemotePath          : https://live.sysinternals.com
ConnectWithUsername : rober
SID                 : S-1-5-21-376445358-2603134888-3166729622-1001
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
Blocks Windows 11 update. Requires Windows 10 version 21H1 or 21H2.
.EXAMPLE
PS> Block-vtsWindows11Upgrade

Output:
The operation completed successfully.
The operation completed successfully.
Success - Current Version (21H2)
#>
function Block-vtsWindows11Upgrade {
    $buildNumber = [System.Environment]::OSVersion.Version.Build

    switch ($buildNumber) {
        19044 {
            cmd /c 'reg add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /v TargetReleaseversion /t REG_DWORD /d 1'
            cmd /c 'reg add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /v TargetReleaseversionInfo /t REG_SZ /d 21H1'
            if ($?) {
                Write-Host 'Success - Current Version (21H2)' -ForegroundColor Green
            }
            else {
                Write-Host "Failed" -ForegroundColor Red
            }
        }
        19043 {
            cmd /c 'reg add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /v TargetReleaseversion /t REG_DWORD /d 1'
            cmd /c 'reg add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /v TargetReleaseversionInfo /t REG_SZ /d 21H1'
            if ($?) {
                Write-Host 'Success - Current Version (21H1)' -ForegroundColor Green
            }
            else {
                Write-Host "Failed" -ForegroundColor Red
            }
        }
        Default { Write-Host "Script only works for Windows 10 versions 21H1 and 21H2" }
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
#>
function Get-vtsDisplayConnectionType {
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
        $connectionType = ($connections | Where-Object { $_.InstanceName -eq $monitor.InstanceName }).VideoOutputTechnology
        if ($manufacturer -ne $null) { $manufacturer = [System.Text.Encoding]::ASCII.GetString($manufacturer -ne 0) }
        if ($name -ne $null) { $name = [System.Text.Encoding]::ASCII.GetString($name -ne 0) }
        $connectionType = $adapterTypes."$connectionType"
        if ($connectionType -eq $null) { $connectionType = 'Unknown' }
        if (($manufacturer -ne $null) -or ($name -ne $null)) { $arrMonitors += "$manufacturer $name ($connectionType)" }
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
Adds a Google Chrome extension to the forced install list. Can be used for forcing installation of any Google Chrome extension. Takes existing extensions into account which might be added by other means, such as GPO and MDM.
.EXAMPLE
PS> Install-vtsChromeExtension -extensionId "ddloeodolhdfbohkokiflfbacbfpjahp"
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
PS> Set-vtsDefaultPrinter -Name brother
.EXAMPLE
PS> Set-vtsDefaultPrinter -Name "HP Laserjet"
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
#>
function Install-vtsPwsh {
    msiexec.exe /i "https://github.com/PowerShell/PowerShell/releases/download/v7.3.3/PowerShell-7.3.3-win-x64.msi" /qn
    Write-Host "Installing PowerShell 7... Please wait" -ForegroundColor Cyan
    While (-not (Test-Path "C:\Program Files\PowerShell\7\pwsh.exe" 2>$null)) { Start-Sleep 5 }
    & "C:\Program Files\PowerShell\7\pwsh.exe" 2>$null
}

<#
.DESCRIPTION
Trace-vtsSession tracks what you do to assist with ticket notes.
.EXAMPLE
PS> Trace-vtsSession
#>
function Trace-vtsSession {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $OpenAIKey
    )
    
    # Check if pwsh is installed. If not installed, install it
    if (-not (Test-Path "C:\Program Files\PowerShell\7\pwsh.exe" -ErrorAction SilentlyContinue)) {
        Write-Host "PowerShell 7 not found, downloading and installing..."
        msiexec.exe /i "https://github.com/PowerShell/PowerShell/releases/download/v7.3.3/PowerShell-7.3.3-win-x64.msi" /qn
    
        while (-not (Test-Path "C:\Program Files\PowerShell\7\pwsh.exe" -ErrorAction SilentlyContinue)) {
            Start-Sleep -Seconds 1
        }
    }

    & "C:\Program Files\PowerShell\7\pwsh.exe" -Command {
        param (
            $PassedOpenAIKey
        )
        try {
            Read-Host "`nPress Enter to begin"
        
            $rec = @"
██████╗ ███████╗ ██████╗ ██████╗ ██████╗ ██████╗ ██╗███╗   ██╗ ██████╗
██╔══██╗██╔════╝██╔════╝██╔═══██╗██╔══██╗██╔══██╗██║████╗  ██║██╔════╝
██████╔╝█████╗  ██║     ██║   ██║██████╔╝██║  ██║██║██╔██╗ ██║██║  ███╗
██╔══██╗██╔══╝  ██║     ██║   ██║██╔══██╗██║  ██║██║██║╚██╗██║██║   ██║
██║  ██║███████╗╚██████╗╚██████╔╝██║  ██║██████╔╝██║██║ ╚████║╚██████╔╝██╗██╗██╗
╚═╝  ╚═╝╚══════╝ ╚═════╝ ╚═════╝ ╚═╝  ╚═╝╚═════╝ ╚═╝╚═╝  ╚═══╝ ╚═════╝ ╚═╝╚═╝╚═╝

"@
            Write-Host $rec -ForegroundColor Red
            Write-Host "Press Ctrl-C when finished." -ForegroundColor Yellow
            While ($true) {
                $dir = "C:\temp\PSDocs"

                # Define Funtions
                function Start-KeyLogger($Path = "C:\temp\PSDocs\keylogger.txt") {
                    # records all key presses until script is aborted

                    # Signatures for API Calls
                    $signatures = @'
[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
public static extern short GetAsyncKeyState(int virtualKeyCode); 
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int GetKeyboardState(byte[] keystate);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int MapVirtualKey(uint uCode, int uMapType);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpkeystate, System.Text.StringBuilder pwszBuff, int cchBuff, uint wFlags);
'@

                    # load signatures and make members available
                    $API = Add-Type -MemberDefinition $signatures -Name 'Win32' -Namespace API -PassThru
    
                    # create output file
                    $null = New-Item -Path $Path -ItemType File -Force

                    # # Add Beginning Timestamp
                    # Add-Content -Path $Path -Value "Start time: $(Get-Date)"

                    # Creates loop that exits when PSR is no longer running.
                    while (get-process psr) {
                        Start-Sleep -Milliseconds 10 #20 #40
      
                        # scan all ASCII codes above 8
                        for ($ascii = 9; $ascii -le 254; $ascii++) {
                            # get current key state
                            $state = $API::GetAsyncKeyState($ascii)

                            # is key pressed?
                            if ($state -eq -32767) {
                                $null = [console]::CapsLock

                                # translate scan code to real code
                                $virtualKey = $API::MapVirtualKey($ascii, 3)

                                # get keyboard state for virtual keys
                                $kbstate = New-Object Byte[] 256
                                $checkkbstate = $API::GetKeyboardState($kbstate)

                                # prepare a StringBuilder to receive input key
                                $mychar = New-Object -TypeName System.Text.StringBuilder

                                # translate virtual key
                                $success = $API::ToUnicode($ascii, $virtualKey, $kbstate, $mychar, $mychar.Capacity, 0)

                                if ($success) {
                                    # add key to logger file
                                    [System.IO.File]::AppendAllText($Path, $mychar, [System.Text.Encoding]::Unicode) 
                                }
                            }
                        }
                    }
                    # # Add Ending Timestamp
                    # Add-Content -Path $Path -Value "End time: $(Get-Date)"
                }

                # Create $path directory if it doesn't exist
                if (-not (Test-Path $dir)) { mkdir $dir | Out-Null }

                $timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss-ff

                # Start PSR
                psr.exe /start /output "$dir\problem_steps_record-$($timestamp).zip" /gui 0 /sc 1 #/maxsc 100

                # Start Keylogger
                Start-KeyLogger
            }
        }
        finally {
            Write-Host "`nRecording complete...`n" -ForegroundColor Green
            Write-Host "Processing...`n`n" -ForegroundColor Green
            if ($null -eq $PassedOpenAIKey) {
                $OpenAIKey = Read-Host -Prompt "Enter OpenAI API Key" -AsSecureString
            }
            else {
                $OpenAIKey = $PassedOpenAIKey
            }

            $dir = "C:\temp\PSDocs"

            # Stop PSR
            psr.exe /stop

            Start-Sleep 3

            # Get PSR Results
            Expand-Archive (Get-ChildItem C:\temp\PSDocs\*.zip | Sort-Object LastWriteTime | Select-Object -last 1) $dir
            Start-Sleep -Milliseconds 250
            $PSRFile = (Get-ChildItem $dir\*.mht | Sort-Object LastWriteTime | Select-Object -last 1)
            $PSRResult = (Get-Content $PSRFile | select-string "^        <p><b>") -replace '^        <p><b>', '' -replace '</b>', '' -replace '</p>', '' -replace '&quot;', "'"

            # Get Keylogger Results
            $KeyloggerResult = Get-Content "$dir\keylogger.txt"

            $StepCount = $PSRResult.Count - 2

            # Compile Results
            $Result = @(
                "RecordedSteps:"
                $PSRResult[0..$StepCount]
                ""
                "Keylogger:"
                $KeyloggerResult
            )

            $prompt = "Act as IT Technician. Based on the following Keyloger and RecordedSteps sections, intrepret what the tech was trying to do while speaking in first person. `
            Don't include that the Problem Steps Recorder was used. `
            Don't include anything related to DesktopWindowXaml. `
            Don't include the word AI. `
            Skip steps that don't make logical sense. `
            Only speak in complete sentences. `
            Do not include key presses like:  [Ctrl] [Alt] [Del] [Enter].
        
            $Result"
            
        
            $body = @{
                'prompt'            = $prompt;
                'temperature'       = 0;
                'max_tokens'        = 500 #250;
                'top_p'             = 1.0;
                'frequency_penalty' = 0.0;
                'presence_penalty'  = 0.0;
                'stop'              = @('"""');
            }
                
            $response = Invoke-RestMethod -Uri "https://api.openai.com/v1/engines/text-davinci-003/completions" -Method Post -Body ($body | ConvertTo-Json) -Headers @{ Authorization = "Bearer $OpenAIKey" } -ContentType "application/json"
            $($response.choices.text) | Out-File "C:\temp\PSDocs\gpt_result.txt" -Force

            Start-sleep -Milliseconds 250
            Write-Host "$(Get-Content "C:\temp\PSDocs\gpt_result.txt")`n`n" -ForegroundColor Yellow
        }
    } -Args $OpenAIKey
    
}