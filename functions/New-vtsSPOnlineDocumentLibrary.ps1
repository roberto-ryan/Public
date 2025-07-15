function New-vtsSPOnlineDocumentLibrary {
  <#
  .SYNOPSIS
  This function creates a new SharePoint Online document library.
  
  .DESCRIPTION
  The New-vtsSPOnlineDocumentLibrary function creates a new document library in a specified SharePoint Online site. 
  It requires the organization name, site URL, and the name of the new document library as parameters. 
  If the PnP.PowerShell module is not installed, the function will install it for the current user.
  
  .PARAMETER orgName
  The name of the organization. This parameter is mandatory. Example: contoso
  
  .PARAMETER siteUrl
  The URL of the SharePoint Online site where the new document library will be created. This parameter is mandatory. Example: https://contoso.sharepoint.com/sites/test
  
  .PARAMETER libraryName
  The name of the new document library to be created. This parameter is mandatory.
  
  .EXAMPLE
  New-vtsSPOnlineDocumentLibrary -orgName "contoso" -siteUrl "https://contoso.sharepoint.com/sites/test" -libraryName "NewLibrary"
  This example creates a new document library named "NewLibrary" in the "test" site of the "contoso" organization.
  
  .LINK
  SharePoint Online
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the organization name. Example: contoso")]
    [string]$orgName,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the site name. Example: https://contoso.sharepoint.com/sites/test")]
    [string]$siteUrl,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the name of the new document library.")]
    [string]$libraryName
  )

  $PnPConnected = Get-PnPConnection

  if (-not($PnPConnected)){
    if (-not(Get-Module PnP.PowerShell -ListAvailable)) {
      Install-Module -Name PnP.PowerShell -Scope CurrentUser
    }
  
    Import-Module PnP.PowerShell
  
    # Connect to SharePoint Online
    # Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Connect-PnPOnline -Url $siteUrl -Interactive
  }

  # Create a new document library
  New-PnPList -Title $libraryName -Template DocumentLibrary -Url $libraryName
}

