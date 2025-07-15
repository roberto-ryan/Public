function Suspend-vtsADUser {
  <#
  .SYNOPSIS
  This function suspends an Active Directory user and optionally schedules a task to reactivate the user after a specified number of days.
  
  .DESCRIPTION
  The Suspend-vtsADUser function fetches all Active Directory users and displays them in an Out-GridView for selection. After a user is selected, the function prompts for the number of days after which to reactivate the user. If no input is entered, the user is disabled. If a number of days is entered, the function disables the user and schedules a task to re-enable the user after the specified number of days.
  
  .PARAMETER None
  This function does not take any parameters.
  
  .EXAMPLE
  Suspend-vtsADUser
  
  This example shows how to run the function without any parameters. It will fetch all Active Directory users, allow you to select a user, and then prompt for the number of days after which to reactivate the user.
  
  .NOTES
  This function requires the Active Directory module for Windows PowerShell and the Task Scheduler cmdlets.
  
  .LINK
  Active Directory
  #>

  $users = Get-ADUser -Filter * -Property DisplayName, PhysicalDeliveryOfficeName, Manager | Select-Object DisplayName, PhysicalDeliveryOfficeName, @{Name = 'Manager'; Expression = { (Get-ADUser $_.Manager).Name } }, SamAccountName | Sort-Object DisplayName
    
  $selectedUser = $users | Out-GridView -Title "Select a User to Manage" -PassThru
    
  if ($null -ne $selectedUser) {
    $days = Read-Host "Enter the number of days after which to reactivate the user (Leave empty to disable the user)"
        
    $verificationUser = Read-Host "Please type $($selectedUser.SamAccountName) to confirm suspension."

    if ($verificationUser -ne $selectedUser.SamAccountName) {
      Write-Host "Verification failed. Exiting..." -ForegroundColor Red
      break
    }

    if ([string]::IsNullOrWhiteSpace($days)) {
      Disable-ADAccount -Identity $selectedUser.SamAccountName
      Write-Host "User $($selectedUser.SamAccountName) has been disabled."
    }
    else {
      try {
        $days = [int]$days
        Disable-ADAccount -Identity $selectedUser.SamAccountName
                
        $action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-NoProfile -WindowStyle Hidden -Command `"Enable-ADAccount -Identity $($selectedUser.SamAccountName)`""
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddDays($days).Date
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
            
        $taskName = "EnableADUser_" + $selectedUser.SamAccountName
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Re-enable AD user $($selectedUser.SamAccountName) after $days days. Task created $(Get-Date) by $($env:USERNAME)" | Out-Null
                
        if ($?) {
          Write-Host "`nUser $($selectedUser.SamAccountName) has been disabled. A task has been scheduled to re-enable the account after $days days at:"
          Write-Host "`n$((Get-Date).AddDays($days).Date)`n" -ForegroundColor Yellow
        }
      }
      catch {
        Write-Host "An error occurred: $_"
      }
    }
        
  }
  else {
    Write-Host "No user was selected."
  }

}

