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

