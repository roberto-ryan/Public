function Install-vtsWindowsUpdate {
  <#
  .DESCRIPTION
  Installs all pending Windows Updates.
  .EXAMPLE
  PS> Install-vtsWindowsUpdate
  
  Output:
  X ComputerName Result     KB          Size Title
  - ------------ ------     --          ---- -----
  1 CH-BIMA-W... Accepted   KB5018202   68MB 2022-10 Cumulative Update Preview for .NET Framework 3.5, 4.8 and 4.8...
  2 CH-BIMA-W... Downloaded KB5018202   68MB 2022-10 Cumulative Update Preview for .NET Framework 3.5, 4.8 and 4.8...
  3 CH-BIMA-W... Installed  KB5018202   68MB 2022-10 Cumulative Update Preview for .NET Framework 3.5, 4.8 and 4.8...
  Reboot is required. Do it now? [Y / N] (default is 'N')
  
  .LINK
  Package Management
  #>
  $NuGet = Get-PackageProvider -Name NuGet
  if ($null -eq $NuGet) {
    Install-PackageProvider -Name NuGet -Force
  }

  $PSWindowsUpdate = Get-Module -Name PSWindowsUpdate
  if ($null -eq $PSWindowsUpdate) {
    Install-Module PSWindowsUpdate -Force -Confirm:$false
  }

  Import-Module PSWindowsUpdate -Force
  Get-WindowsUpdate
  Install-WindowsUpdate -AcceptAll
}

