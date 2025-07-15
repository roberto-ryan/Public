function Install-vtsPwsh {
  <#
  .DESCRIPTION
  Install PowerShell 7
  .EXAMPLE
  PS> Install-vtsPwsh
  
  .LINK
  Package Management
  #>
  param (
    [switch]$InstallLatestVersionWithGUI
  )

  if ($InstallLatestVersionWithGUI){
    iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI"
  } else {
    msiexec.exe /i "https://github.com/PowerShell/PowerShell/releases/download/v7.4.6/PowerShell-7.4.6-win-x64.msi" /qn
    Write-Host "Installing PowerShell 7... Please wait" -ForegroundColor Cyan
    While (-not (Test-Path "C:\Program Files\PowerShell\7\pwsh.exe" 2>$null)) { Start-Sleep 5 }
    & "C:\Program Files\PowerShell\7\pwsh.exe" 2>$null
  }
}

