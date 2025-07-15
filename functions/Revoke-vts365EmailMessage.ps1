function Revoke-vts365EmailMessage {
  <#
  .SYNOPSIS
  This function revokes a specific email message in Office 365.
  
  .DESCRIPTION
  The Revoke-vts365EmailMessage function connects to Exchange Online and IPPS Session, gets the message trace, starts a compliance search, waits for the search to complete, purges and deletes the email instances, and checks the status of the search action.
  
  .PARAMETER TicketNumber
  This is a mandatory parameter. It is the unique identifier for the compliance search.
  
  .PARAMETER From
  This is an optional parameter. It is the sender's email address.
  
  .PARAMETER To
  This is an optional parameter. It is the recipient's email address.
  
  .EXAMPLE
  Revoke-vts365EmailMessage -TicketNumber "12345" -From "sender@example.com" -To "recipient@example.com"
  
  This example shows how to revoke an email message sent from "sender@example.com" to "recipient@example.com" with the ticket number "12345".
  
  .LINK
  Still in Development
  
  #>
  param(
    # [Parameter(Mandatory = $true)]
    [string]$TicketNumber = "TEST4",
    [array]$From,
    [array]$To = "Zach.Koscoe@completehealth.com"
  )
    
  Write-Host "Connecting to Exchange Online and IPPS Session..."
  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }
  # Connect-IPPSSession

  Write-Host "Getting message trace..."
  if (($From -ne $null) -and ($To -ne $null)) {
    $Message = Get-MessageTrace -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date) -RecipientAddress $To -SenderAddress $From | Out-GridView -OutputMode Multiple
  }
    
  if (($To -ne $null) -and ($From -eq $null)) {
    $Message = Get-MessageTrace -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date) -RecipientAddress $To | Out-GridView -OutputMode Multiple
  }
    
  if (($To -eq $null) -and ($From -ne $null)) {
    $Message = Get-MessageTrace -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date) -SenderAddress $From | Out-GridView -OutputMode Multiple
  }
    
  Write-Host "Starting compliance search..."
  New-ComplianceSearch -Name $TicketNumber -ExchangeLocation $($Message.RecipientAddress) -ContentMatchQuery "(c:c)(from=$($Message.SenderAddress))(subjecttitle=""$($Message.subject)"")" | Start-ComplianceSearch
    
  Write-Host "Waiting for the search to complete..."
  do {
    Start-Sleep -Seconds 60
    $searchStatus = (Get-ComplianceSearch -Identity $TicketNumber).Status
  } while ($searchStatus -ne 'Completed')
    
  Write-Host "Purging and deleting the email instances..."
  New-ComplianceSearchAction -SearchName $TicketNumber -Purge -PurgeType HardDelete
    
  Write-Host "Checking the status of the search action..."
  do {
    Start-Sleep -Seconds 60
    $actionStatus = Get-ComplianceSearchAction | Where-Object Name -eq "$($TicketNumber)_Purge" | Select-Object -expand Status
  } while ($actionStatus -ne 'Completed')

  Write-Host "Operation completed."
}

