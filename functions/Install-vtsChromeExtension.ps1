function Install-vtsChromeExtension {
  <#
  .DESCRIPTION
  Forces the installation of a specified Google Chrome extension.
  .EXAMPLE
  PS> Install-vtsChromeExtension -extensionId "ddloeodolhdfbohkokiflfbacbfpjahp"
  
  .LINK
  Package Management
  #>
  param(
    [string]$extensionId,
    [switch]$info
  )
  if ($info) {
    $InformationPreference = "Continue"
  }
  if (!($extensionId)) {
    # Empty Extension
    $result = "No Extension ID"
  }
  else {
    Write-Information "ExtensionID = $extensionID"
    $extensionId = "$extensionId;https://clients2.google.com/service/update2/crx"
    $regKey = "HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist"
    if (!(Test-Path $regKey)) {
      New-Item $regKey -Force
      Write-Information "Created Reg Key $regKey"
    }
    # Add Extension to Chrome
    $extensionsList = New-Object System.Collections.ArrayList
    $number = 0
    $noMore = 0
    do {
      $number++
      Write-Information "Pass : $number"
      try {
        $install = Get-ItemProperty $regKey -name $number -ErrorAction Stop
        $extensionObj = [PSCustomObject]@{
          Name  = $number
          Value = $install.$number
        }
        $extensionsList.add($extensionObj) | Out-Null
        Write-Information "Extension List Item : $($extensionObj.name) / $($extensionObj.value)"
      }
      catch {
        $noMore = 1
      }
    }
    until($noMore -eq 1)
    $extensionCheck = $extensionsList | Where-Object { $_.Value -eq $extensionId }
    if ($extensionCheck) {
      $result = "Extension Already Exists"
      Write-Information "Extension Already Exists"
    }
    else {
      $newExtensionId = $extensionsList[-1].name + 1
      New-ItemProperty HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist -PropertyType String -Name $newExtensionId -Value $extensionId
      $result = "Installed"
    }
  }
  $result
}

