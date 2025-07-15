function Set-vts365UserLicense {
  <#
  .SYNOPSIS
  This function manages licenses for a list of users by adding and removing licenses in a single operation.
  
  .DESCRIPTION
  The function Set-vts365UserLicense connects to the Graph API and manages licenses for selected users. 
  It allows selecting users from Exchange Online and choosing which licenses to add and remove.
  
  .EXAMPLE
  Set-vts365UserLicense
  
  .LINK
  M365
  #>
  param (
    [Parameter(Mandatory = $false, HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  if (-not $UserList) {
    Write-Host "Connecting to Exchange Online to retrieve mailboxes..."
    Connect-ExchangeOnline | Out-Null
    $UserList = get-user | 
    Where UserPrincipalName -notlike "*.onmicrosoft.com*" | 
    sort DisplayName | 
    select DisplayName, UserPrincipalName | 
    Out-GridView -Title "Select users to manage licenses for" -OutputMode Multiple | 
    Select -expand UserPrincipalName
  }

  if (-not $UserList) {
    Write-Host "No users selected." -ForegroundColor Red
    return
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

  Write-Host "`nSelect licenses to ADD:"
  $AddLicenses = $LicenseDisplay | 
      Out-GridView -Title "Select licenses to ADD (or cancel for none)" -OutputMode Multiple | 
      Select-Object -ExpandProperty SkuID

  Write-Host "`nSelect licenses to REMOVE:"
  $RemoveLicenses = $LicenseDisplay | 
      Out-GridView -Title "Select licenses to REMOVE (or cancel for none)" -OutputMode Multiple | 
      Select-Object -ExpandProperty SkuID

  $FormattedAddSKUs = @()
  foreach ($Sku in $AddLicenses) {
      $FormattedAddSKUs += @{SkuId = $Sku }
  }

  Write-Host "Processing license changes for users..."
  foreach ($User in $UserList) {
      Write-Host "Managing licenses for $User..."
      Set-MgUserLicense -UserId $User -AddLicenses $FormattedAddSKUs -RemoveLicenses $RemoveLicenses
  }
  Write-Host "Operation completed."
}

