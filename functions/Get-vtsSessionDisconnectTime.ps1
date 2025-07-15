function Get-vtsSessionDisconnectTime {
  <#
  .SYNOPSIS
  This function retrieves the disconnect time of a session.
  
  .DESCRIPTION
  The Get-vtsSessionDisconnectTime function uses the Get-WinEvent cmdlet to retrieve the event logs for a specific EventID from the Microsoft-Windows-TerminalServices-LocalSessionManager/Operational log. It then parses the XML of each event to get the username and the disconnect time. It returns an array of custom objects, each containing a username and a disconnect time.
  
  .PARAMETER EventID
  The ID of the event to filter the event logs. The default value is 24.
  
  .PARAMETER LogName
  The name of the log to get the event logs from. The default value is "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational".
  
  .EXAMPLE
  PS C:\> Get-vtsSessionDisconnectTime
  
  This command retrieves the disconnect time of all VTS sessions.
  
  .NOTES
  Additional information about the function.
  
  .LINK
  Log Management
  #>
  param(
    [int]$EventID = 24,
    [string]$LogName = "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational"
  )
    
  $events = Get-WinEvent -FilterHashtable @{LogName = $LogName; ID = $EventID }
  $output = @()
  foreach ($event in $events) {
    $xml = [xml]$event.ToXml()
    $eventUsername = $xml.Event.UserData.EventXML.User
    $eventTime = $event.TimeCreated
    
    $output += [pscustomobject]@{
      Username       = $eventUsername
      DisconnectTime = $eventTime
    }
  }
    
  return $output
}

