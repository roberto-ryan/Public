function Enable-vtsTLS12AndStrongCrypto {
  <#
  .SYNOPSIS
  Enables TLS 1.2 for server and client, sets default secure protocols for WinHTTP, and enables strong cryptography for .NET Framework.
  
  .DESCRIPTION
  This function configures the necessary registry settings to enable TLS 1.2 for both server and client, sets the default secure protocols for WinHTTP, and enables strong cryptography for the .NET Framework.
  
  .EXAMPLE
  Enable-vtsTLS12AndStrongCrypto
  
  .LINK
  Network
  
  #>

  # Enable TLS 1.2 for Server and Client
  $TLSPaths = @(
      "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server",
      "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"
  )

  foreach ($path in $TLSPaths) {
      if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
      Set-ItemProperty -Path $path -Name "Enabled" -Value 1 -Type DWord
      Set-ItemProperty -Path $path -Name "DisabledByDefault" -Value 0 -Type DWord
  }

  # Set Default Secure Protocols for WinHTTP
  $WinHttpPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp"
  if (-not (Test-Path $WinHttpPath)) { New-Item -Path $WinHttpPath -Force | Out-Null }
  Set-ItemProperty -Path $WinHttpPath -Name "DefaultSecureProtocols" -Value 0xA00 -Type DWord

  # Enable Strong Cryptography for .NET Framework
  $DotNetPaths = @(
      "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
      "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319"
  )

  foreach ($path in $DotNetPaths) {
      if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
      Set-ItemProperty -Path $path -Name "SchUseStrongCrypto" -Value 1 -Type DWord
  }

  Write-Output "TLS 1.2 and .NET Strong Cryptography settings have been configured. Please restart the server for the changes to take effect."
}

