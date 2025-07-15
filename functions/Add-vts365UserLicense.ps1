function Add-vts365UserLicense {
  <#
  .SYNOPSIS
  This function adds licenses to a list of users.
  
  .DESCRIPTION
  The function Add-vts365UserLicense connects to the Graph API and adds licenses to a list of users. If no user list is specified, it connects to Exchange Online and allows the user to select mailboxes dynamically. The function also allows the user to select which licenses to add.
  
  .PARAMETER UserList
  An optional parameter. A single string or comma separated email addresses to which the licenses are to be added. If not specified, the user will be prompted to select from a list of all mailboxes.
  
  .EXAMPLE
  Add-vts365UserLicense -UserList "user1@domain.com, user2@domain.com"
  
  .EXAMPLE
  Add-vts365UserLicense
  
  This example prompts the user to select from a list of all mailboxes and then adds licenses to the selected users.
  
  .LINK
  M365
  #>
  param (
    [Parameter(HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  if (-not $UserList) {
    Write-Host "Connecting to Exchange Online to retrieve mailboxes..."
    Connect-ExchangeOnline | Out-Null
    $UserList = get-user | Where UserPrincipalName -notlike "*.onmicrosoft.com*" | sort DisplayName | select DisplayName, UserPrincipalName | Out-GridView -OutputMode Multiple | Select -expand UserPrincipalName
    if (-not $UserList) {
      Write-Host "No users selected or no mailboxes available." -ForegroundColor Red
      return
    }
  } else {
    # Split the user list and trim whitespace
    $UserList = ($UserList -split ",").Trim()
  }

  Write-Host "Connecting to Graph API..."
  Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All -Device

  Write-Host "Getting available SKUs..."
  $SKUs = Get-MgSubscribedSku

  $key = 0
  $LicenseDisplay = @()
  foreach ($item in $SKUs) {
    $key++
    $LicenseDisplay += [PSCustomObject]@{
      Key           = $key
      SkuPartNumber = $item.SkuPartNumber
      SkuID         = $item.SkuID
    }
  }

  Write-Host "Displaying available licenses..."
  $LicenseDisplay | Out-Host

  $userInput = Read-Host "`nEnter the numbers of the licenses you want to add, separated by commas, or enter * to add all available licenses"
  
  if ($userInput -eq '*') {
    Write-Host "Adding all available licenses..."
    $SelectedLicenses = $LicenseDisplay.SkuID
  }
  else {
    Write-Host "Adding selected licenses..."
    $SelectedLicenses = $LicenseDisplay | Where-Object Key -in ($userInput -split ",") | Select-Object -ExpandProperty SkuID
  }
  
  $FormattedSKUs = @()
  foreach ($Sku in $SelectedLicenses) {
    $FormattedSKUs += @{SkuId = $Sku }
  }
  
  Write-Host "Adding licenses to users..."
  foreach ($User in $UserList) {
    Write-Host "Adding licenses to $User..."
    Set-MgUserLicense -UserId $User -AddLicenses $FormattedSKUs -RemoveLicenses @()
  }
  Write-Host "Operation completed."
}

