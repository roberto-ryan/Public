function Set-vts365CalendarPermissions {
  <#
  .SYNOPSIS
  This function sets the calendar permissions for a specified user in Office 365.
  
  .DESCRIPTION
  The Set-vts365CalendarPermissions function connects to Exchange Online and modifies the calendar permissions for a specified user. It first identifies the calendar path and backs up the existing permissions. Then, it adds or modifies the permissions for the specified access user. Finally, it verifies the updated permissions.
  
  .PARAMETER user
  The email address of the user whose calendar permissions you want to modify. The format should be 'user@domain.com'.
  
  .PARAMETER accessUser
  The email address of the user to whom you want to grant or modify access to the calendar. The format should be 'user@domain.com'.
  
  .PARAMETER accessRights
  The level of access rights you want to grant to the access user. Acceptable values are 'Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author', 'NonEditingAuthor', 'Reviewer', 'Contributor'.
  
  .EXAMPLE
  Set-vts365CalendarPermissions -user 'user1@domain.com' -accessUser 'user2@domain.com' -accessRights 'Editor'
  
  This example modifies the calendar permissions for user1@domain.com, granting 'Editor' access to user2@domain.com.
  
  .NOTES
  The function will backup the existing permissions to C:\temp\CalendarPermissionsBackup.txt before making any changes.
  
  .LINK
  M365
  #>
  param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the email address of the user whose calendar permissions you want to modify. The format should be 'user@domain.com'")]
    $user,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the email address of the user to whom you want to grant or modify access to the calendar. The format should be 'user@domain.com'")]
    $accessUser,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the level of access rights you want to grant to the access user. Acceptable values are 'Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author', 'NonEditingAuthor', 'Reviewer', 'Contributor'")]
    [ValidateSet('Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author', 'NonEditingAuthor', 'Reviewer', 'Contributor')]
    [string]$accessRights
  )

  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }

  # Identify the Calendar Path
  $calendarPath = $user + ":\Calendar"
  Write-Host "Identified Calendar Path: $calendarPath"

  if (-not(Test-Path -Path C:\temp)) {
    New-Item -ItemType Directory -Path C:\temp
  }

  # Backup Calendar Permissions
  Write-Host "Backing up Calendar Permissions to C:\temp\CalendarPermissionsBackup.txt"
  "Original Permissions" | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
  Get-Date | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
  $Permissions = Get-MailboxFolderPermission $calendarPath | Format-Table -AutoSize
  $Permissions | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
    
  # View Existing Calendar Permissions
  Write-Host "Viewing Existing Calendar Permissions..."
  $Permissions
    
  foreach ($member in $accessUser) {
    # Add New Permissions
    Write-Host "Adding New Permissions for $member with access rights $accessRights..."
    try {
      Add-MailboxFolderPermission $calendarPath -User $member -AccessRights $accessRights -ErrorAction Stop
      Write-Host "New permissions for $member with access rights $accessRights added successfully."
    }
    catch {
      Write-Host "Error adding permissions for $member with access rights $accessRights`: $_"
      # Modify Existing Permissions (Optional)
      Write-Host "Attempting to Modify Existing Permissions for $member with access rights $accessRights..."
      try {
        Set-MailboxFolderPermission $calendarPath -User $member -AccessRights $accessRights -ErrorAction Stop
        Write-Host "Permissions for $member with access rights $accessRights modified successfully."
      }
      catch {
        Write-Host "Error setting permissions for $member with access rights $accessRights`: $_"
      }
    }
  }
    
  # Backup Calendar Permissions
  Write-Host "Backing up Calendar Permissions to C:\temp\CalendarPermissionsBackup.txt"
  "Modified Permissions" | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
  Get-Date | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append
  $Permissions = Get-MailboxFolderPermission $calendarPath | Format-Table -AutoSize
  $Permissions | Out-File -FilePath "C:\temp\CalendarPermissionsBackup.txt" -Force -Append

  # Verify the Updated Permissions
  Write-Host "Verifying the Updated Permissions for $calendarPath"
  $Permissions
}

