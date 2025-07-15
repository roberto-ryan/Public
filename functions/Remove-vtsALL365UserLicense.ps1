function Remove-vtsALL365UserLicense {
  <#
  .SYNOPSIS
  This function removes all licenses from a list of users.
  
  .DESCRIPTION
  The function Remove-vtsALL365UserLicense connects to the Graph API and removes all licenses from a list of users. The user list is passed as a parameter to the function.
  
  .PARAMETER UserList
  A single string or comma separated email addresses from which the licenses are to be removed.
  
  .EXAMPLE
  Remove-vtsALL365UserLicense -UserList "user1@domain.com, user2@domain.com"
  
  .LINK
  M365
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  $UserList = ($UserList -split ",").Trim()

  Write-Host "Connecting to Graph API..."
  Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All -Device

  Write-Host "Removing licenses from users..."
  foreach ($User in $UserList) {
    Write-Host "Removing licenses from $User..."
    $UserLicenses = Get-MgUserLicenseDetail -UserId $User | Select-Object -ExpandProperty SkuId
    Set-MgUserLicense -UserId $User -AddLicenses @() -RemoveLicenses $UserLicenses
  }
  Write-Host "Operation completed."
}

