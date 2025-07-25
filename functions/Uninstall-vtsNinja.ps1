function Uninstall-vtsNinja {
  <#
  .DESCRIPTION
  Uninstall ninja and remove keys, files and services.
  .EXAMPLE
  PS> Uninstall-vtsNinja
  
  .LINK
  Package Management
  #>
  param (
    [Parameter(Mandatory = $false)]
    [switch]$DelTeamViewer = $false,
    [Parameter(Mandatory = $false)]
    [switch]$Cleanup = $true,
    [Parameter(Mandatory = $false)]
    [switch]$Uninstall = $true
  )
    
  #Set-PSDebug -Trace 2
    
  if ([system.environment]::Is64BitOperatingSystem) {
    $ninjaPreSoftKey = 'HKLM:\SOFTWARE\WOW6432Node\NinjaRMM LLC'
    $uninstallKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    $exetomsiKey = 'HKLM:\SOFTWARE\WOW6432Node\EXEMSI.COM\MSI Wrapper\Installed'
  }
  else {
    $ninjaPreSoftKey = 'HKLM:\SOFTWARE\NinjaRMM LLC'
    $uninstallKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    $exetomsiKey = 'HKLM:\SOFTWARE\EXEMSI.COM\MSI Wrapper\Installed'
  }
    
  $ninjaSoftKey = Join-Path $ninjaPreSoftKey -ChildPath 'NinjaRMMAgent'
    
  $ninjaDir = [string]::Empty
  $ninjaDataDir = Join-Path -Path $env:ProgramData -ChildPath "NinjaRMMAgent"
    
  ###################################################################################################
  # locating NinjaRMMAgent
  ###################################################################################################
  $ninjaDirRegLocation = $(Get-ItemPropertyValue $ninjaSoftKey -Name Location)
  if ($ninjaDirRegLocation) {
    if (Join-Path -Path $ninjaDirRegLocation -ChildPath "NinjaRMMAgent.exe" | Test-Path) {
      #location confirmed from registry location
      $ninjaDir = $ninjaDirRegLocation
    }
  }
    
  if (!$ninjaDir) {
    #attempt to get the path from service
    $ss = Get-WmiObject win32_service -Filter 'Name Like "NinjaRMMAgent"'
    if ($ss) {
      $ninjaDirService = ($(Get-WmiObject win32_service -Filter 'Name Like "NinjaRMMAgent"').PathName | Split-Path).Replace("`"", "")
      if (Join-Path -Path $ninjaDirService -ChildPath "NinjaRMMAgentPatcher.exe" | Test-Path) {
        #location confirmed from service location
        $ninjaDir = $ninjaDirService
      }
    }
  }
    
  if ($ninjaDir) {
    $ninjaDir.Replace('/', '\')
    
    if ($Uninstall) {
      #there are few measures agent takes to prevent accidental uninstllation
      #disable those measures now
      #it automatically takes care if those measures are already removed
      #it is not possible to check those measures outside of the agent since agent's development comes parralel to this script
      & "$ninjaDir\NinjaRMMAgent.exe" -disableUninstallPrevention NOUI
    
      # Executes uninstall.exe in Ninja install directory
      $Arguments = @(
        "/uninstall"
        $(Get-WmiObject -Class win32_product -Filter "Name='NinjaRMMAgent'").IdentifyingNumber
        "/quiet"
        "/log"
        "NinjaRMMAgent_uninstall.log"
        "/L*v"
        "WRAPPED_ARGUMENTS=`"--mode unattended`""
      )
    
      Start-Process -FilePath "msiexec.exe"  -Verb RunAs -Wait -NoNewWindow -WhatIf -ArgumentList $Arguments
    }
  }
    
    
  if ($Cleanup) {
    $service = Get-Service "NinjaRMMAgent" 
    if ($service) {
      Stop-Service $service -Force
      & sc.exe DELETE NinjaRMMAgent
      #Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NinjaRMMAgent
    }
    
    # Delete Ninja install directory and all contents
    if (Test-Path $ninjaDir) {
      & cmd.exe /c rd /s /q $ninjaDir
    }
    
    if (Test-Path $ninjaDataDir) {
      & cmd.exe /c rd /s /q $ninjaDataDir
    }
    
    #Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\NinjaRMM LLC\NinjaRMMAgent
    Remove-Item -Path  -Recurse -Force
    
    # Will search registry locations for NinjaRMMAgent value and delete parent key
    # Search $uninstallKey
    $keys = Get-ChildItem $uninstallKey | Get-ItemProperty -name 'DisplayName'
    foreach ($key in $keys) {
      if ($key.'DisplayName' -eq 'NinjaRMMAgent') {
        Remove-Item $key.PSPath -Recurse -Force
      }
    }
    
    #Search $installerKey
    $keys = Get-ChildItem 'HKLM:\SOFTWARE\Classes\Installer\Products' | Get-ItemProperty -name 'ProductName'
    foreach ($key in $keys) {
      if ($key.'ProductName' -eq 'NinjaRMMAgent') {
        Remove-Item $key.PSPath -Recurse -Force
      }
    }
    # Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\A0313090625DD2B4F824C1EAE0958B08\InstallProperties
    $keys = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products'
    foreach ($key in $keys) {
      #ridiculous, MS sucks
      $kn = $key.Name -replace 'HKEY_LOCAL_MACHINE' , 'HKLM:'; 
      $k1 = Join-Path $kn -ChildPath 'InstallProperties';
      if ( $(Get-ItemProperty -Path $k1 -Name DisplayName).DisplayName -eq 'NinjaRMMAgent') {
        $toremove = 
        Get-Item -LiteralPath $kn | Remove-Item -Recurse -Force
      }
    }
    
    #Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\EXEMSI.COM\MSI Wrapper\Installed\NinjaRMMAgent 5.3.3681
    Get-ChildItem $exetomsiKey | Where-Object -Property Name -CLike '*NinjaRMMAgent*'  | Remove-Item -Recurse -Force
    
    #HKLM:\SOFTWARE\WOW6432Node\NinjaRMM LLC
    Get-Item -Path $ninjaPreSoftKey | Remove-Item -Recurse -Force
    
    # agent creates this key by mistake but we delete it here
    Get-Item -Path "HKLM:\SOFTWARE\WOW6432Node\WOW6432Node\NinjaRMM LLC" | Remove-Item -Recurse -Force
    
  }
    
  if (Get-Item -Path $ninjaPreSoftKey) {
    Write-Output "Failed to remove NinjaRMMAgent reg keys ", $ninjaPreSoftKey
  }
    
  if (Get-Service "NinjaRMMAgent") {
    Write-Output "Failed to remove NinjaRMMAgent service"
  }
    
  if ($ninjaDir) {
    if (Test-Path $ninjaDir) {
      Write-Output "Failed to remove NinjaRMMAgent program folder"
      if (Join-Path -Path $ninjaDir -ChildPath "NinjaRMMAgent.exe" | Test-Path) {
        Write-Output "Failed to remove NinjaRMMAgent.exe"
      }
    
      if (Join-Path -Path $ninjaDir -ChildPath "NinjaRMMAgentPatcher.exe" | Test-Path) {
        Write-Output "Failed to remove NinjaRMMAgentPatcher.exe"
      }
    }
  }
    
    
    
  # Uninstall TeamViewer only if -DelTeamViewer parameter specified
  if ($DelTeamViewer -eq $true) {
    $tvProcess = Get-Process -Name 'teamviewer'
    Stop-Process -InputObject $tvProcess -Force # Stops TeamViewer process
    & '%ProgramFiles%\TeamViewer\uninstall.exe' /S | out-null
    & '%ProgramFiles(x86)%\TeamViewer\uninstall.exe' /S | out-null
    Remove-Item -path HKLM:\SOFTWARE\TeamViewer -Recurse
    Remove-Item -path HKLM:\SOFTWARE\WOW6432Node\TeamViewer -Recurse
    Remove-Item -path HKLM:\SOFTWARE\WOW6432Node\TVInstallTemp -Recurse
    Remove-Item -path HKLM:\SOFTWARE\TeamViewer -Recurse
    Remove-Item -path HKLM:\SOFTWARE\Wow6432Node\TeamViewer -Recurse
  }    
}

