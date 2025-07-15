function Set-vtsShortcut {
  <#
  .SYNOPSIS
  Modifies properties of an existing Windows shortcut file.
  
  .DESCRIPTION
  The Set-vtsShortcut function allows modification of Windows shortcut (.lnk) file properties including the target path, hotkey, arguments, and icon location. It accepts input from the pipeline or direct parameters.
  
  .PARAMETER LinkPath
  The full path to the shortcut file to modify.
  
  .PARAMETER Hotkey
  The keyboard shortcut to assign to the shortcut file.
  
  .PARAMETER IconLocation 
  The path to the icon file and icon index to use.
  
  .PARAMETER Arguments
  The command-line arguments to pass to the target application.
  
  .PARAMETER TargetPath
  The path to the target file that the shortcut will launch.
  
  .EXAMPLE
  PS> Get-vtsShortcut "C:\shortcut.lnk" | Set-vtsShortcut -TargetPath "C:\NewTarget.exe"
  Modifies the target path of an existing shortcut.
  
  .EXAMPLE
  PS> Set-vtsShortcut -LinkPath "C:\shortcut.lnk" -Hotkey "CTRL+ALT+F"
  Sets a keyboard shortcut for an existing shortcut file.
  
  .EXAMPLE
  Get-ChildItem -Path "C:\Path\To\Your\Directory" -Include *.lnk -Recurse -file | 
  Select-Object -expand fullname |
  ForEach-Object { 
    Get-vtsShortcut $_ | 
    Where-Object TargetPath -like "*192.168.1.220*" | 
    ForEach-Object {
      Set-vtsShortcut -LinkPath "$($_.LinkPath)" -IconLocation "$($_.IconLocation)" -TargetPath "$(($_.TargetPath) -replace '192.168.1.220','192.168.5.220')"
    }
  }
  Recursively searches for all `.lnk` (shortcut) files in a specified directory. It then filters these shortcuts to find those whose target path contains the IP address `192.168.1.220`. For each matching shortcut, it updates the target path to replace `192.168.1.220` with `192.168.5.220`, while preserving the original link path and icon location.
  
  .LINK
  File Management
  #>
  param(
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    $LinkPath,
    $Hotkey,
    $IconLocation,
    $Arguments,
    $TargetPath
  )
  begin {
    $shell = New-Object -ComObject WScript.Shell
  }
  
  process {
    $link = $shell.CreateShortcut($LinkPath)

    $PSCmdlet.MyInvocation.BoundParameters.GetEnumerator() |
    Where-Object { $_.key -ne 'LinkPath' } |
    ForEach-Object { $link.$($_.key) = $_.value }
    $link.Save()
  }
}

