function Get-FTA {
  <#
  .DESCRIPTION
      Get the default application associated with a specific file extension. If no extension is provided, it returns a list of all file extensions and their associated default applications.
  
  .NOTES
      Version    : 1.2.0
      Author(s)  : Danyfirex & Dany3j
      Credits    : https://bbs.pediy.com/thread-213954.htm
                   LMongrain - Hash Algorithm PureBasic Version
      License    : MIT License
      Copyright  : 2022 Danysys. <danysys.com>
  
  .EXAMPLE
      Get-FTA .pdf
      Show Default Application Program Id for an Extension
  
  .LINK
  File Association Management
      #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [String]
    $Extension
  )

  
  if ($Extension) {
    Write-Verbose "Get File Type Association for $Extension"
    
    $assocFile = (Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension\UserChoice" -ErrorAction SilentlyContinue).ProgId
    Write-Output $assocFile
  }
  else {
    Write-Verbose "Get File Type Association List"

    $assocList = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\* |
    ForEach-Object {
      $progId = (Get-ItemProperty "$($_.PSParentPath)\$($_.PSChildName)\UserChoice" -ErrorAction SilentlyContinue).ProgId
      if ($progId) {
        "$($_.PSChildName), $progId"
      }
    }
    Write-Output $assocList
  }
  
}

