<#
.Description
Searches the last 500 System and Application logs for a specified search term.
.EXAMPLE
PS> Search-vtsEventLog -SearchTerm <search term>
.EXAMPLE
PS> Search-vtsEventLog -SearchTerm driver

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
            cmd /c 'reg add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /v TargetReleaseversionInfo /t REG_SZ /d 21H2'
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

    foreach ($temp in $t.CurrentTemperature)
    {


    $currentTempKelvin = $temp / 10
    $currentTempCelsius = $currentTempKelvin - 273.15

    $currentTempFahrenheit = (9/5) * $currentTempCelsius + 32

    $returntemp += $currentTempCelsius.ToString() + " C : " + $currentTempFahrenheit.ToString() + " F : " + $currentTempKelvin + "K"  
    }
    return $returntemp
}