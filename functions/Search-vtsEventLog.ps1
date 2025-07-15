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

