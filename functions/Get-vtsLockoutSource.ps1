function Get-vtsLockoutSource {
  <#
  .SYNOPSIS
  This function retrieves and returns the source of a user's lockout event from the Security log. If no lockout events are detected, it offers to enable logging for account lockouts.
  
  .DESCRIPTION
  The Get-vtsLockoutSource function employs the Get-WinEvent cmdlet to search for a lockout event for a specified user in the Security log. 
  If a lockout event is identified, the function generates a custom object that includes the time of the lockout event, the locked user's username, and the source of the lockout. 
  In the absence of a lockout event, the function prompts the user to enable logging for account lockouts. If the user agrees, the function enables logging for account lockouts and informs the user to wait for another account lockout to retry.
  
  .PARAMETER user
  This optional parameter specifies the username for which the lockout source is to be retrieved. If not provided, the function will return all lockout events.
  
  .EXAMPLE
  PS C:\> Get-vtsLockoutSource -user "jdoe"
  
  This command initiates the retrieval of the source of the lockout event for the user "jdoe". If no lockout events are detected for "jdoe", it will offer to enable logging for account lockouts.
  
  .INPUTS
  System.String
  
  .OUTPUTS
  PSCustomObject
  
  .LINK
  Log Management
  #>
  param(
    [Parameter(Mandatory = $false)]
    [string]$user
  )

  $logs = Get-WinEvent -FilterHashtable @{LogName = 'Security'; Id = 4740 } 2>$null

  if ($null -eq $logs) {
    Write-Host "$($env:COMPUTERNAME) - No lockout events have been detected in the logs. It's possible that logging for account lockouts is currently disabled. Would you like to activate this feature now? (y/n)"
    $EnableLogs = Read-Host
    if ($EnableLogs -eq "y") {
      Auditpol /set /category:"Account Logon" /success:enable /failure:enable | Out-Null
      Auditpol /set /category:"Logon/Logoff" /success:enable /failure:enable | Out-Null
      Auditpol /set /category:"Account Management" /success:enable /failure:enable | Out-Null
      if ($?) {
        Write-Host "Logging has been successfully enabled. Please wait for the occurrence of another account lockout to retry." -ForegroundColor Yellow
      }
    }
    return
  }
  # Return all lockout events in a formatted PowerShell table
  $logs | Where-Object Message -like "*$user*" | Select-Object TimeCreated, @{n = 'LockedUser'; e = { $_.Properties[0].Value } }, @{n = 'LockoutSource'; e = { $_.Properties[1].Value } }, @{n = 'LogSource'; e = { $_.Properties[4].Value } } | Format-Table -AutoSize
}

