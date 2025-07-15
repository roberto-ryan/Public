function New-vtsMappedDrive {
  <#
  .DESCRIPTION
  Maps a remote drive.
  .EXAMPLE
  PS> New-vtsMappedDrive -Letter A -Path \\192.168.0.4\sharedfolder
  .EXAMPLE
  PS> New-vtsMappedDrive -Letter A -Path "\\192.168.0.4\folder with spaces"
  
  .LINK
  Drive Management
  #>
  param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Letter,
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
  )
  New-PSDrive -Name "$Letter" -PSProvider FileSystem -Root "$Path" -Persist -Scope Global
}

