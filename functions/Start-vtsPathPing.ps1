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

