function Connect-vtsWiFi {
  <#
  .SYNOPSIS
  This script connects to a WiFi network using the provided SSID and password.
  
  .DESCRIPTION
  The Connect-vtsWiFi function connects to a WiFi network using the provided SSID and password. If no SSID is provided, the function will list all available networks and prompt the user to select one. If no password is provided, the function will prompt the user to enter one. If a WiFi profile for the selected SSID already exists, the function will ask the user if they want to remove the old profile and replace it.
  
  .PARAMETER SSID
  The SSID of the WiFi network to connect to. If not provided, the function will list all available networks and prompt the user to select one.
  
  .PARAMETER Password
  The password for the WiFi network to connect to. If not provided, the function will prompt the user to enter one.
  
  .EXAMPLE
  Connect-vtsWiFi -SSID "MyWiFiNetwork" -Password "MyPassword"
  
  This example connects to the WiFi network "MyWiFiNetwork" using the password "MyPassword".
  
  .LINK
  Network
  #>
  param (
    [Parameter(Mandatory = $false)]
    [string]$SSID,
      
    [Parameter(Mandatory = $false)]
    [SecureString]$Password
  )

  Write-Host "`nChecking for wifiprofilemanagement module..."
  if (-not (Get-Module -ListAvailable -Name wifiprofilemanagement)) {
    Write-Host "`nInstalling wifiprofilemanagement module..."
    Install-Module -Name wifiprofilemanagement -Force -Scope CurrentUser
  }

  if (-not $SSID) {
    Write-Host "`nFetching available networks...`n"

    $networks = (netsh wlan show networks | sls ^SSID) -replace "SSID . : " | sort

    $i = 1
    $networkList = @()
    foreach ($network in $networks) {
      Write-Host "$i`: $network"
      $networkList += $network
      $i++
    }
    $selection = Read-Host -Prompt "`nPlease select a network by number"
    $SSID = ($networkList[$selection - 1]).Trim()
  }

  if (-not $Password) {
    $Password = Read-Host -Prompt "Please enter the password for $SSID" -AsSecureString
  }

  Write-Host "`nChecking for existing WiFi profile..."
  if (Get-WiFiProfile -ProfileName $SSID 2>$null) {
    $userChoice = Read-Host -Prompt "WiFi profile for $SSID already exists. Do you want to remove the old profile and replace it? (y/n)"
    if ($userChoice -eq "y") {
      Write-Host "`nRemoving old WiFi profile..."
      Remove-WiFiProfile -ProfileName $SSID
      Write-Host "`nCreating new WiFi profile..."
      New-WiFiProfile -ProfileName $SSID -ConnectionMode auto -Authentication WPA2PSK -Password $Password -Encryption AES
    }
  }
  else {
    Write-Host "`nCreating new WiFi profile..."
    New-WiFiProfile -ProfileName $SSID -ConnectionMode auto -Authentication WPA2PSK -Password $Password -Encryption AES
  }
  Write-Host "`nConnecting to WiFi profile..."
  Connect-WiFiProfile -ProfileName $SSID

}

