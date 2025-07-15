function Get-vtsSPOnlineDocumentLibraryFolders {
  <#
  .SYNOPSIS
  This function retrieves all folders in a specified SharePoint Online document library.
  
  .DESCRIPTION
  The Get-vtsSPOnlineDocumentLibraryFolders function retrieves all folders in a specified SharePoint Online document library. 
  It requires the site URL and the name of the document library as parameters. 
  If the PnP.PowerShell module is not installed, the function will install it for the current user.
  
  .PARAMETER siteUrl
  The URL of the SharePoint Online site where the document library is located. This parameter is mandatory. Example: https://contoso.sharepoint.com/sites/test
  
  .PARAMETER libraryName
  The name of the document library from which to retrieve folders. This parameter is mandatory.
  
  .EXAMPLE
  Get-vtsSPOnlineDocumentLibraryFolders -siteUrl "https://contoso.sharepoint.com/sites/test" -libraryName "LibraryName"
  This example retrieves all folders in the "LibraryName" document library on the "test" site.
  
  .LINK
  SharePoint Online
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the site URL. Example: https://contoso.sharepoint.com/sites/test")]
    [string]$siteUrl,
    [Parameter(Mandatory = $true, HelpMessage = "Enter the document library name.")]
    [string]$libraryName,
    [switch]$Recursive
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

  # Get all folders in the document library
  if ($Recursive){
    $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $libraryName -ItemType Folder -Recursive
  } else {
    $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $libraryName -ItemType Folder
  }
  # Return the folder names
  return $folders.Name
}

