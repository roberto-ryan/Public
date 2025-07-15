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

