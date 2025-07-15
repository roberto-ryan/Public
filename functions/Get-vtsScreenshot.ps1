function Get-vtsScreenshot {
  <#
  .SYNOPSIS
     A script to take a screenshot of the current screen and save it to a specified path.
  
  .DESCRIPTION
     This script uses the System.Windows.Forms and System.Drawing assemblies to capture a screenshot of the current screen. 
     The screenshot is saved as a .png file at the path specified by the $Path parameter. 
     If no path is specified, the screenshot will be saved in the temp folder with a timestamp in the filename.
  
  .PARAMETER Path
     The path where the screenshot will be saved. If not specified, the screenshot will be saved in the temp folder with a timestamp in the filename.
  
  .EXAMPLE
     Get-vtsScreenshot -Path "C:\Users\Username\Pictures\Screenshot.png"
     This command will take a screenshot and save it as Screenshot.png in the Pictures folder of the user Username.
  
  .LINK
      Utilities
  #>
  param (
    [string]$Path = "$env:temp\$(Get-Date -f yyyy-MM-dd-HH-mm)-Screenshot.png"
  )

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing

  $Screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
  $Width = $Screen.Width
  $Height = $Screen.Height
  $Left = $Screen.Left
  $Top = $Screen.Top

  $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
  $graphic = [System.Drawing.Graphics]::FromImage($bitmap)
  $graphic.CopyFromScreen($Left, $Top, 0, 0, $bitmap.Size)

  $bitmap.Save($Path)

  Write-Host "Screenshot saved at $Path"
}

