function Add-vtsSharePointAdminToAllSites {
  <#
  .SYNOPSIS
      This function adds a specified user as an admin to all SharePoint sites.
  
  .DESCRIPTION
      The Add-vtsSharePointAdminToAllSites function connects to SharePoint Online using the provided admin site URL, retrieves all site collections, and adds the specified user as an admin to each site.
  
  .PARAMETER adminSiteUrl
      The URL of the SharePoint admin site. This parameter is mandatory.
  
  .PARAMETER userToMakeOwner
      The username of the user to be made an admin. This parameter is mandatory.
  
  .EXAMPLE
      Add-vtsSharePointAdminToAllSites -adminSiteUrl "https://contoso-admin.sharepoint.com" -userToMakeOwner "user@contoso.com"
      This example adds the user "user@contoso.com" as an admin to all SharePoint sites in the "contoso" tenant.
  
  .NOTES
      This function requires the Microsoft.Online.SharePoint.PowerShell module. If the module is not installed, the function will install it.
      Requres PowerShell 5. PowerShell 7 doesn't work.
  
  .LINK
      M365
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the SharePoint admin site URL in the format 'https://yourdomain-admin.sharepoint.com'")]
    [string]$adminSiteUrl,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the username of the user to be made an admin in the format 'user@yourdomain.com'")]
    [string]$userToMakeOwner
  )
  if (!(Get-InstalledModule Microsoft.Online.SharePoint.PowerShell)) {
    Install-Module Microsoft.Online.SharePoint.PowerShell
  }

  Import-Module Microsoft.Online.SharePoint.PowerShell

  Write-Host "Connecting to SharePoint Online..."
  Connect-SPOService -Url $adminSiteUrl

  Write-Host "Retrieving all site collections..."
  $sites = Get-SPOSite -Limit All

  foreach ($site in $sites) {
    Write-Host "Adding owner to site: " $site.Url
    Set-SPOUser -Site $site.Url -LoginName $userToMakeOwner -IsSiteCollectionAdmin $true

    Write-Host "Successfully added admin to site:" $site.Url
  }
}

