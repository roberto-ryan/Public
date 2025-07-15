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

