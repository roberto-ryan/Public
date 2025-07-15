function Set-PTA {
  <#
  .DESCRIPTION
      Set the default application for a specific protocol. It takes a ProgId, a Protocol, and an optional Icon as parameters. 
      The ProgId is the identifier of the application to be set as default. The Protocol is the protocol for which the application will be set as default. 
      The Icon is an optional parameter that sets the icon for the application.
  
  .NOTES
      Version    : 1.2.0
      Author(s)  : Danyfirex & Dany3j
      Credits    : https://bbs.pediy.com/thread-213954.htm
                   LMongrain - Hash Algorithm PureBasic Version
      License    : MIT License
      Copyright  : 2022 Danysys. <danysys.com>
    
  .EXAMPLE
      Set-PTA ChromeHTML http
      Set Google Chrome as Default for http Protocol
  
  .LINK
  File Association Management
      #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [String]
    $ProgId,

    [Parameter(Mandatory = $true)]
    [String]
    $Protocol,
      
    [String]
    $Icon
  )

  Set-FTA -ProgId $ProgId -Protocol $Protocol -Icon $Icon
}

