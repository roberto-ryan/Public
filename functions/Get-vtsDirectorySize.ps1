function Get-vtsDirectorySize {
  <#
  .SYNOPSIS
  This function calculates the size of a directory in MB or GB.
  
  .DESCRIPTION
  The Get-vtsDirectorySize function calculates the size of a directory and returns the size in MB or GB. If the size is greater than 1024 MB, it will be converted to GB.
  
  .PARAMETER Path
  The path of the directory you want to calculate the size of. If no path is provided, the function will calculate the size of the current directory.
  
  .EXAMPLE
  Get-vtsDirectorySize -Path "C:\Windows"
  This command will calculate the size of the Windows directory.
  
  .LINK
  File Management
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [string]$Path = (Get-Location).Path
  )

  try {
    # Calculate the size of the directory
    $size = (Get-ChildItem $Path -Recurse -ErrorAction Stop | Measure-Object -Property Length -Sum).Sum
    $sizeInMB = "{0:N2}" -f ($size / 1MB)
    $sizeInGB = "{0:N2}" -f ($size / 1GB)

    # Output the size in MB or GB
    if ($sizeInMB -gt 1024) {
      Write-Output ("The size of $Path is " + $sizeInGB + " GB")
    }
    else {
      Write-Output ("The size of $Path is " + $sizeInMB + " MB")
    }
  }
  catch {
    Write-Error "An error occurred while calculating the size of the directory: $_"
  }
}

