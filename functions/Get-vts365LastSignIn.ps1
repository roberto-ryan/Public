function Get-vts365LastSignIn {
  <#
  .SYNOPSIS
  Retrieves the last sign-in information for all Microsoft 365 users.
  
  .DESCRIPTION
  The Get-vts365LastSignIn function connects to Microsoft Graph, retrieves all users, and gathers the most recent sign-in log for each user. It outputs the user's display name, user principal name, last sign-in time, application used, and the client application used for the last sign-in.
  
  .EXAMPLE
  PS C:\> Get-vts365LastSignIn
  
  This command retrieves the last sign-in information for all Microsoft 365 users and displays it in the console.
  
  .EXAMPLE
  PS C:\> Get-vts365LastSignIn
  Do you want to export a report? (Y/N): Y
  
  This command retrieves the last sign-in information for all Microsoft 365 users, displays it in the console, and prompts the user to export the information to a CSV file.
  
  .NOTES
  Requires the Microsoft.Graph module.
  
  .LINK
  M365
  #>

  Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All"

  # Get all users
  $users = Get-MgUser -All

  # Initialize an array to store results
  $results = @()

  Write-Host "Processing users..."
  foreach ($user in $users) {
      # Get sign-in logs for the user
      $signInLogs = Get-MgAuditLogSignIn -Filter "userId eq '$($user.Id)'" -Top 1 -OrderBy "createdDateTime desc"
  
      if ($signInLogs) {
          $lastSignIn = $signInLogs.CreatedDateTime
          $appDisplayName = $signInLogs.AppDisplayName
          $clientAppUsed = $signInLogs.ClientAppUsed
      }
      else {
          $lastSignIn = "No sign-in record found"
          $appDisplayName = "N/A"
          $clientAppUsed = "N/A"
      }
  
      # Create a custom object with user info and last sign-in details
      $userInfo = [PSCustomObject]@{
          UserPrincipalName = $user.UserPrincipalName
          DisplayName       = $user.DisplayName
          LastSignInTime    = $lastSignIn
          AppDisplayName    = $appDisplayName
          ClientAppUsed     = $clientAppUsed
      }
  
      # Add the user info to the results array
      $results += $userInfo
  
      # Display info in the console
      Write-Host "User: $($userInfo.DisplayName) ($($userInfo.UserPrincipalName))"
      Write-Host "  Last Sign-In: $($userInfo.LastSignInTime)"
      Write-Host "  App: $($userInfo.AppDisplayName)"
      Write-Host "  Client: $($userInfo.ClientAppUsed)"
      Write-Host "------------------------"
  }

  # Ask user if they want to export a report
  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"

  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
      # Check if PSWriteHTML module is installed, if not, install it
      if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
          Install-Module -Name PSWriteHTML -Force -Confirm:$false
      }
    
      # Export the results to an HTML file using the PSWriteHTML module
      $results | Out-HtmlView
  }

  # Disconnect from Microsoft Graph
  Disconnect-MgGraph
}

