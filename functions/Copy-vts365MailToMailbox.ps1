function Copy-vts365MailToMailbox {
  <#
  .SYNOPSIS
  This function copies emails from a specified sender within a date range from all mailboxes to a target mailbox and folder.
  
  .DESCRIPTION
  The Copy-vts365MailToMailbox function connects to Exchange Online PowerShell and creates a new compliance search for emails from a specified sender within a specified date range. It waits for the compliance search to complete, gets the results, parses them, and creates objects from the results. It then gets the mailboxes to search, performs the search, and copies the emails to the target mailbox's specified folder.
  
  .EXAMPLE
  Copy-vts365MailToMailbox -senderAddress "sender@example.com" -targetMailbox "target@example.com" -targetFolder "Folder" -startDate "01/01/2020" -endDate "12/31/2020" -SearchName "Search1"
  This example copies all emails from sender@example.com sent between 01/01/2020 and 12/31/2020 from all mailboxes to the "Folder" in the target@example.com mailbox.
  
  .LINK
  M365
  #>
  param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the sender's email address")]
    [string]$senderAddress,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the target mailbox to copy the results")]
    [string]$targetMailbox,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the target folder to copy the results")]
    [string]$targetFolder,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the start date in the format 'MM/dd/yyyy'")]
    [string]$startDate,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the end date in the format 'MM/dd/yyyy'")]
    [string]$endDate,

    [string]$SearchName = "Copy-vts365MailToMailbox-$(Get-Date -Format MM-dd-yy-mm-ss)"
  )

  # Connect to Exchange Online PowerShell
  Write-Host "Connecting to Exchange Online PowerShell..."
  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }
  Write-Host "Connecting to Security & Compliance PowerShell..."
  Connect-IPPSSession -ShowBanner:$false

  # Create a new compliance search
  Write-Host "Creating new compliance search..."
  New-ComplianceSearch -Name "$SearchName" -ExchangeLocation All -ContentMatchQuery "from:$senderAddress AND received>=$startDate AND received<=$endDate" | Out-Null
  Start-ComplianceSearch -Identity "$SearchName" | Out-Null

  # Wait for the compliance search to complete
  Write-Host "Waiting for compliance search to complete..."
  while ((Get-ComplianceSearch -Identity "$SearchName" | Select-Object -ExpandProperty Status) -ne "Completed") {
    Start-Sleep -Seconds 5
    Get-ComplianceSearch -Identity "$SearchName" | Out-Null
  }
    
  # Get the compliance search results
  Write-Host "Getting compliance search results..."
  $Results = Get-ComplianceSearch -Identity "$SearchName" | Select-Object -expand SuccessResults

  # Parse the results
  Write-Host "Parsing results..."
  $array = ($Results -replace "{|}" -split ",").Trim()

  # Create objects from the results
  Write-Host "Creating objects from results..."
  $objects = for ($i = 0; $i -lt $array.Count; $i += 3) {
    New-Object PSObject -Property @{
      Location  = ($array[$i] -split ": ")[1]
      ItemCount = [int]($array[$i + 1] -split ": ")[1]
      TotalSize = [int]($array[$i + 2] -split ": ")[1]
    }
  }

  # Get the mailboxes to search
  Write-Host "Getting mailboxes to search...`n"
  $MailboxesWithContent = $objects | Where-Object ItemCount -gt 0 | Sort-Object ItemCount -Descending | Select-Object Location, ItemCount, TotalSize

  # Initialize a hashtable for mailboxes
  $mailboxTable = @()
  $key = 1

  # Iterate over each mailbox to search
  foreach ($Box in $MailboxesWithContent) {
    # Add each mailbox to the hashtable with its corresponding details
    $mailboxTable += [pscustomobject]@{
      Key       = $key
      Location  = $Box.Location
      ItemCount = $Box.ItemCount
      TotalSize = $Box.TotalSize
    }
    $key++
  }

  # Output the mailbox details
  $mailboxTable

  if ($null -ne $mailboxTable) {
    # Ask the user which mailboxes to search
    $userInput = Read-Host "`nEnter the numbers of the mailboxes you want to search, separated by commas, or enter * to search all mailboxes"
  
    if ($userInput -eq '*') {
      $SelectedMailboxes = $mailboxTable.Location
    }
    else {
      $SelectedMailboxes = $mailboxTable | Where-Object Key -in ($userInput -split ",") | Select-Object -ExpandProperty Location
    }
  
    foreach ($mailbox in $SelectedMailboxes) {
      # Perform the search and copy the emails to the target mailbox's inbox
      Write-Host "Performing search and copying emails from mailbox: $mailbox..."
      Search-Mailbox -Identity $mailbox -SearchQuery "from:$senderAddress AND received>=$startDate AND received<=$endDate" -TargetMailbox $targetMailbox -TargetFolder $targetFolder -LogLevel Full 3>$null | Out-Null
    }
  }
  else { 
    Write-Host "No matches found." -ForegroundColor Red
  }

  Write-Host "Operation completed."
}

