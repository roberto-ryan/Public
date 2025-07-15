function Get-vts365UserLicense {
  <#
  .SYNOPSIS
  This function retrieves the license details for a list of users after validating their existence.
  
  .DESCRIPTION
  The function Get-vts365UserLicense connects to the Graph API, validates if the users exist, and retrieves the license details for a list of users. The user list is passed as a parameter to the function.
  
  .PARAMETER UserList
  A single string or comma separated email addresses for which the license details are to be retrieved.
  
  .EXAMPLE
  Get-vts365UserLicense -UserList "user1@domain.com, user2@domain.com"
  
  This example validates and retrieves the license details for user1@domain.com and user2@domain.com.
  
  .LINK
  M365
  #>
  param (
    [Parameter(Mandatory = $false, HelpMessage = "Enter a single email or a comma separated list of emails.")]
    $UserList
  )

  if (-not $UserList) {
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

  Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All

  $LicenseDetails = @()

  foreach ($User in $UserList) {
    $User = $User.Trim()
    $UserExists = Get-MgUser -UserId $User -ErrorAction SilentlyContinue
    if ($UserExists) {
      $LicenseDetails += [pscustomobject]@{
        User    = $User
        License = (Get-MgUserLicenseDetail -UserId $User | Select-Object -expand SkuPartNumber) -join ", "
      }
    }
    else {
      Write-Host "User $User does not exist." -ForegroundColor Red
    }
  }

  $LicenseDetails | Out-Host

  # Ask user if they want to export a report
  $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"
  
  if ($exportReport -eq "Y" -or $exportReport -eq "y") {
    # Check if PSWriteHTML module is installed, if not, install it
    if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
      Install-Module -Name PSWriteHTML -Force -Confirm:$false
    }
        
    # Export the results to an HTML file using the PSWriteHTML module
    $LicenseDetails | Out-HtmlView
  }
}

