function Set-vts365MailboxArchive {
  <#
  .SYNOPSIS
  This function sets up mailbox archiving for a specified user in Office 365.
  
  .DESCRIPTION
  The Set-vts365MailboxArchive function connects to Exchange Online and performs several operations related to mailbox archiving. It allows you to view user retention policies, view user archive details, enable archive, and setup auto-archiving.
  
  .PARAMETER UserEmail
  The email address of the user for whom the mailbox archiving will be set up. The email address should be in 'user@domain.com' format.
  
  .EXAMPLE
  Set-vts365MailboxArchive -UserEmail "user@domain.com"
  
  This example sets up mailbox archiving for the user with the email address "user@domain.com".
  
  .LINK
  M365
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Please enter the user in 'user@domain.com' format")]
    $script:UserEmail
  )

  # Attempt to connect to Exchange Online
  Write-Host "Attempting to connect to Exchange Online..."
  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }

  # Function to view user retention policies
  function ViewUserRetentionPolicy {
    Write-Host "Preparing to view user retention policies..."
    $ViewUserRetentionPolicy = Read-Host "View user retention policies? (y/N)"
    switch ($ViewUserRetentionPolicy) {
      "y" { 
        try {
          get-mailbox -ResultSize Unlimited | Select-Object displayname, UserPrincipalName, RetentionPolicy | Out-Host
          Write-Host "User retention policies viewed successfully."
        }
        catch {
          Write-Error "Failed to view user retention policies: $_"
        }
      }
      Default {}
    }
  }

  # Function to view user archive details
  function ViewUserArchiveDetails {
    Write-Host "Preparing to view user archive details..."
    $ViewUserArchiveDetails = Read-Host "View user archive details? (y/N)"
    switch ($ViewUserArchiveDetails) {
      "y" { 
        try {
          Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, ArchiveStatus, ArchiveDatabase, ArchiveDomain, ArchiveGuid, ArchiveName, ArchiveQuota, ArchiveRelease, ArchiveState, ArchiveStatusDescription, ArchiveWarningQuota | Out-Host
          Write-Host "User archive details viewed successfully."
        }
        catch {
          Write-Error "Failed to view user archive details: $_"
        }
      }
      Default {}
    }
  }

  # Function to enable archive
  function EnableArchive {
    Write-Host "Preparing to enable archive..."
    $EnableUserArchive = Read-Host "Enable Archive? (y/N)"
    switch ($EnableUserArchive) {
      "y" { 
        foreach ($user in $UserEmail) {
          try {
            # Set the archive settings for the new user
            Set-Mailbox $user -ArchiveName "Archive"
            Enable-Mailbox $user -Archive
            Write-Host "Archive enabled successfully for $user."
          }
          catch {
            Write-Error "Failed to enable archive for $user`: $_"
          }
        }
      }
      Default {}
    }
  }

  # Function to enable auto archive
  function EnableAutoArchive {
    Write-Host "Preparing to setup auto-archiving..."
    $SetupAutoArchiving = Read-Host "Setup auto-archiving? (y/N)"

    switch ($SetupAutoArchiving) {
      "y" { 
        $RetentionLimit = Read-Host "Move to archive after this many days"

        # Check if the retention tag exists
        Write-Host "Checking if retention tag exists..."
        try {
          $RetentionTag = Get-RetentionPolicyTag -Identity "Archive after $RetentionLimit Days" -ErrorAction SilentlyContinue
          if ($null -eq $RetentionTag) {
            # Create a retention tag
            New-RetentionPolicyTag -Name "Archive after $RetentionLimit Days" -Type All -AgeLimitForRetention $RetentionLimit -RetentionAction MoveToArchive
            Write-Host "Retention tag created successfully."
          }
        }
        catch {
          Write-Error "Failed to create retention tag: $_"
        }

        # Check if the retention policy exists
        Write-Host "Checking if retention policy exists..."
        try {
          $RetentionPolicy = Get-RetentionPolicy -Identity "$RetentionLimit Day Retention Policy" -ErrorAction SilentlyContinue
          if ($null -eq $RetentionPolicy) {
            # Create a retention policy
            New-RetentionPolicy -Name "$RetentionLimit Day Retention Policy" -RetentionPolicyTagLinks "Archive after $RetentionLimit Days"
            Write-Host "Retention policy created successfully."
          }
        }
        catch {
          Write-Error "Failed to create retention policy: $_"
        }

        foreach ($user in $UserEmail) {
          try {
            # Apply the retention policy to a mailbox
            Set-Mailbox -Identity $user -RetentionPolicy "$RetentionLimit Day Retention Policy"
            Write-Host "Retention policy applied successfully to $user."
          }
          catch {
            Write-Error "Failed to apply retention policy to $user`: $_"
          }
        }
      }
      Default {}
    }
  }

  # Call functions
  ViewUserRetentionPolicy
  ViewUserArchiveDetails

  EnableArchive
  EnableAutoArchive

  ViewUserRetentionPolicy
  ViewUserArchiveDetails
}

