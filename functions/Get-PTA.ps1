function Get-PTA {
  <#
  .DESCRIPTION
      Get the default application associated with a specific protocol. If no protocol is provided, it returns a list of all protocols and their associated default applications.
  
  .NOTES
      Version    : 1.2.0
      Author(s)  : Danyfirex & Dany3j
      Credits    : https://bbs.pediy.com/thread-213954.htm
                   LMongrain - Hash Algorithm PureBasic Version
      License    : MIT License
      Copyright  : 2022 Danysys. <danysys.com>
  
  .EXAMPLE
      Get-PTA
  
  .LINK
  File Association Management
      #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [String]
    $Protocol
  )

  if ($Protocol) {
    Write-Verbose "Get Protocol Type Association for $Protocol"

    $assocFile = (Get-ItemProperty "HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\$Protocol\UserChoice" -ErrorAction SilentlyContinue).ProgId
    Write-Output $assocFile
  }
  else {
    Write-Verbose "Get Protocol Type Association List"

    $assocList = Get-ChildItem HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\* |
    ForEach-Object {
      $progId = (Get-ItemProperty "$($_.PSParentPath)\$($_.PSChildName)\UserChoice" -ErrorAction SilentlyContinue).ProgId
      if ($progId) {
        "$($_.PSChildName), $progId"
      }
    }
    Write-Output $assocList
  }
}

