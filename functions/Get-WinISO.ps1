function Get-WinISO {
  <#
  .SYNOPSIS
  Downloads Microsoft 365 installation files based on specified parameters.
  
  .DESCRIPTION
  The Get-WinISO function retrieves and downloads Microsoft 365 installation files
  from official sources. It supports various options for Windows versions, release types, editions,
  languages, and architectures.
  
  .PARAMETER Win
  Specifies the Windows version. Valid values are 10, 11, or All.
  
  .PARAMETER Rel
  Specifies the release type. Valid values are Latest, Insider, or Dev.
  
  .PARAMETER Ed
  Specifies the edition. Valid values are Home, Pro, Edu, or All.
  
  .PARAMETER Lang
  Specifies the language for the installation files.
  
  .PARAMETER Arch
  Specifies the architecture. Valid values are x86, x64, arm64, or All.
  
  .EXAMPLE
  Get-WinISO -Win 11 -Rel Latest -Ed Pro -Lang English -Arch x64
  
  Downloads the latest Windows 11 Pro 64-bit installation files in English.
  
  .LINK
  Microsoft 365
  #>
  param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('10', '11', 'All')]
    [string]$Win = '11',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('Latest', 'Insider', 'Dev')]
    [string]$Rel = 'Latest',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('Home', 'Pro', 'Edu', 'All')]
    [string]$Ed = 'Pro',
    
    [Parameter(Mandatory = $false)]
    [string]$Lang = 'English',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('x86', 'x64', 'arm64', 'All')]
    [string]$Arch = 'x64'
  )

  # Construct the command invocation
  $command = "(irm https://raw.githubusercontent.com/roberto-ryan/Fido/refs/heads/master/Fido.ps1) | iex ; Download"
  
  # Add parameters if they're specified
  if ($Win -ne 'All') { $command += " -Win $Win" }
  if ($Rel -ne 'Latest') { $command += " -Rel $Rel" }
  if ($Ed -ne 'All') { $command += " -Ed $Ed" }
  if ($Lang) { $command += " -Lang $Lang" }
  if ($Arch -ne 'All') { $command += " -Arch $Arch" }
  
  # Execute the command
  Write-Host "Executing: $command" -ForegroundColor Cyan
  Invoke-Expression $command
}

