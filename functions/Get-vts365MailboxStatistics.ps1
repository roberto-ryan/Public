function Get-vts365MailboxStatistics {
  <#
  .SYNOPSIS
  This function retrieves mailbox statistics for a list of email addresses from Exchange Online.
  
  .DESCRIPTION
  The Get-vts365MailboxStatistics function connects to Exchange Online and retrieves mailbox statistics for each email address provided. The results are stored in an array and outputted at the end. It also provides an option to export the results to an HTML report.
  
  .PARAMETER EmailAddress
  An array of email addresses for which to retrieve mailbox statistics. If not provided, the function retrieves statistics for all mailboxes.
  
  .EXAMPLE
  PS C:\> Get-vts365MailboxStatistics -EmailAddress "user1@example.com", "user2@example.com"
  
  This example retrieves mailbox statistics for user1@example.com and user2@example.com.
  
  .EXAMPLE
  PS C:\> Get-vts365MailboxStatistics
  
  This example retrieves mailbox statistics for all mailboxes.
  
  .LINK
  M365
  #>
  param(
    $EmailAddress
  )

  # Connect to Exchange Online
  Write-Host "Connecting to Exchange Online..."
  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }

  if ($null -eq $EmailAddress) { $EmailAddress = $(get-mailbox | Select-Object -expand UserPrincipalName) }

  # Initialize results array
  $Results = @()

  # Retrieve mailbox statistics for each email address
  foreach ($Address in $EmailAddress) {
    Write-Host "Retrieving mailbox statistics for $Address..."
    $Results += Get-EXOMailboxStatistics -UserPrincipalName $Address
  }

  # Output the results
  Write-Host "Retrieval complete. Here are the results:"
  $Results | Sort-Object TotalItemSize -Descending | Format-Table -AutoSize

  # Ask user if they want to export a report
  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"
  
  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
    # Check if PSWriteHTML module is installed, if not, install it
    if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
      Install-Module -Name PSWriteHTML -Force -Confirm:$false
    }
      
    # Export the results to an HTML file using the PSWriteHTML module
    $Results | Out-HtmlView
  }
}

