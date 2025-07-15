function Set-vtsDirectoryOwnership {
  <#
  .SYNOPSIS
  This function takes ownership and grants full control permissions to a specified directory and its contents.
  
  .DESCRIPTION
  The Set-vtsDirectoryOwnership function takes ownership of a specified directory and grants full control permissions to the current user. It uses the takeown and icacls commands to achieve this. The function can be used to take ownership and set permissions on any directory.
  
  .PARAMETER DirectoryPath
  The path of the directory to take ownership of and grant permissions to. This parameter is mandatory.
  
  .EXAMPLE
  Set-vtsDirectoryOwnership -DirectoryPath "E:\Users\tmays"
  
  This example takes ownership of the "E:\Users\tmays" directory and grants full control permissions to the current user.
  
  .LINK
  File Management
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the path of the directory to take ownership of and grant permissions to.")]
    [string]$DirectoryPath
  )

  if (!(Test-Path -Path $DirectoryPath)) {
    Write-Error "The specified path does not exist."
    return
  }

  try {
    Write-Host "Taking ownership of $DirectoryPath..."
    takeown /F "$DirectoryPath" /R /D Y

    Write-Host "Granting full control permissions to $($env:USERNAME) on $DirectoryPath..."
    icacls "$DirectoryPath" /grant "$($env:USERNAME):F" /T /C /Q

    Write-Host "Ownership and permissions have been successfully updated for $DirectoryPath."
  }
  catch {
    Write-Error "An error occurred while updating ownership and permissions: $_"
  }
}

