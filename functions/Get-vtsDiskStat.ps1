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

