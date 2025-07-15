function Get-vtsShortcut {
  <#
  .SYNOPSIS
  Retrieves shortcut (.lnk) file information from specified paths.
  
  .DESCRIPTION
  The Get-vtsShortcut function retrieves detailed information about Windows shortcut files (.lnk), including their target paths, hotkeys, arguments, and icon locations. If no path is specified, it searches both the current user's and all users' Start Menu folders.
  
  .PARAMETER path
  Optional. The path to search for shortcut files. If not specified, searches Start Menu folders.
  
  .EXAMPLE
  PS> Get-vtsShortcut
  Returns all shortcuts from user and system Start Menu folders.
  
  .EXAMPLE
  PS> Get-vtsShortcut -path "C:\Users\Username\Desktop"
  Returns all shortcuts from the specified desktop folder.
  
  .LINK
  File Management
  #>
  param(
    $path = $null
  )
  
  $obj = New-Object -ComObject WScript.Shell

  if ($path -eq $null) {
    $pathUser = [System.Environment]::GetFolderPath('StartMenu')
    $pathCommon = $obj.SpecialFolders.Item('AllUsersStartMenu')
    $path = Get-ChildItem $pathUser, $pathCommon -Filter *.lnk -Recurse 
  }
  if ($path -is [string]) {
    $path = Get-ChildItem $path -Filter *.lnk
  }
  $path | ForEach-Object { 
    if ($_ -is [string]) {
      $_ = Get-ChildItem $_ -Filter *.lnk
    }
    if ($_) {
      $link = $obj.CreateShortcut($_.FullName)

      $info = @{}
      $info.Hotkey = $link.Hotkey
      $info.TargetPath = $link.TargetPath
      $info.LinkPath = $link.FullName
      $info.Arguments = $link.Arguments
      $info.Target = try { Split-Path $info.TargetPath -Leaf } catch { 'n/a' }
      $info.Link = try { Split-Path $info.LinkPath -Leaf } catch { 'n/a' }
      $info.WindowStyle = $link.WindowStyle
      $info.IconLocation = $link.IconLocation

      New-Object PSObject -Property $info
    }
  }
}

