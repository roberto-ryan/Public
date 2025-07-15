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

