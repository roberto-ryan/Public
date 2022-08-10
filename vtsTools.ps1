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

    $result | Format-List
}

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
        if ($null -ne $manufacturer) { $manufacturer = [System.Text.Encoding]::ASCII.GetString($manufacturer -ne 0) }
        if ($null -ne $name) { $name = [System.Text.Encoding]::ASCII.GetString($name -ne 0) }
        $connectionType = $adapterTypes."$connectionType"
        if ($null -eq $connectionType) { $connectionType = 'Unknown' }
        if (($null -ne $manufacturer) -or ($null -ne $name)) { $arrMonitors += "$manufacturer $name ($connectionType)" }
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