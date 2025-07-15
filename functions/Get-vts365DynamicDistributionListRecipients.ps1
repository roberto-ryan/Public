function Get-vts365DynamicDistributionListRecipients {
  <#
  .SYNOPSIS
  This function retrieves the recipients of a given set of distribution lists from Exchange Online.
  
  .DESCRIPTION
  The Get-vts365DynamicDistributionListRecipients function connects to Exchange Online and retrieves the recipients of the specified dynamic distribution lists. The results are stored in an array and outputted at the end. It also provides an option to export the results to an HTML report.
  
  .PARAMETER DistributionList
  An array of distribution list names for which to retrieve recipients.
  
  .EXAMPLE
  PS C:\> Get-vts365DynamicDistributionListRecipients -DistributionList "Group1", "Group2"
  
  .LINK
  M365
  #>

  if (-not(Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false
  }

  if((Get-ConnectionInformation).TenantID -like "7*4*b*5*-*4*3-*8*f-*f1*-6*"){
    $IncludeOfficeInfo = $true
  }

  # Get all dynamic distribution groups and suppress error output
  $distributionGroups = Get-DynamicDistributionGroup | Select-Object -ExpandProperty Name 2>$null
  
  # Initialize an empty array for the group table
  $GroupTable = @()
  $key = 1

  # Populate the group table with group names and corresponding keys
  foreach ($group in $distributionGroups) {
    $GroupTable += [pscustomobject]@{
      Key   = $key
      Group = $group
    }
    $key++
  }
  
  # Check if the current user is system authority
  if ($(whoami) -eq "nt authority\system") {
    $GroupTable | Out-Host
    $userInput = Read-Host "Please input the group numbers you wish to query, separated by commas. Alternatively, input '*' to search all groups."
    if ("$userInput" -eq '*') {
      Write-Host "Searching all available groups..."
      $SelectedGroups = $GroupTable.Group
    } else {
      Write-Host "Searching selected groups..."
      $SelectedGroups = $GroupTable | Where-Object Key -in ("$userInput" -split ",") | Select-Object -ExpandProperty Group
    }
  } else {
    $SelectedGroups = $GroupTable | Out-GridView -OutputMode Multiple | Select-Object -ExpandProperty Group
  }

  # Get recipients based on the selected groups and output their details
  $Results = foreach ($group in $SelectedGroups) {
    $DDLGroup = Get-DynamicDistributionGroup -Identity $group
    if($IncludeOfficeInfo){
    Get-Recipient -ResultSize Unlimited -RecipientPreviewFilter ($DDLGroup.RecipientFilter) | ForEach-Object {
      [pscustomobject]@{
        GroupName   = $DDLGroup
        DisplayName = $_.DisplayName
        Email       = $_.PrimarySmtpAddress
        PositionID  = $_.notes
        Title       = $_.title
        Office      = $_.office
      }
    } | Sort-Object DisplayName, Office, Title} else{
      Get-Recipient -ResultSize Unlimited -RecipientPreviewFilter ($DDLGroup.RecipientFilter) | ForEach-Object {
        [pscustomobject]@{
          GroupName   = $DDLGroup
          DisplayName = $_.DisplayName
          Email       = $_.PrimarySmtpAddress
        }
      } | Sort-Object DisplayName
    }
  }

  if($IncludeOfficeInfo){
    # Filter out results where PositionID, Title, or Office properties are null
    $Results = $Results | Where-Object { ![string]::IsNullOrWhiteSpace($_.PositionID) -and ![string]::IsNullOrWhiteSpace($_.Title) -and ![string]::IsNullOrWhiteSpace($_.Office) }
  }

  $Results | Format-Table -AutoSize | Out-Host

  # Ask user if they want to export a report
  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"
  
  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
    # Check if PSWriteHTML module is installed, if not, install it
    if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
      Install-Module -Name PSWriteHTML -Force -Confirm:$false
    }
        
    # Export the results to an HTML file using the PSWriteHTML module
    $Results | Out-HtmlView -Title "Dynamic Distribution List Recipients Report - $(Get-Date -f "dddd MM-dd-yyyy HHmm")" -Filepath "C:\Reports\Dynamic Distribution List Recipients Report - $(Get-Date -f MM_dd_yyyy).html"
  }
}

