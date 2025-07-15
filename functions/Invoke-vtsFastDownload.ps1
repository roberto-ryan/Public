function Invoke-vtsFastDownload {
  <#
  .SYNOPSIS
  This function downloads a file from a given URL using the fast download utility aria2.
  
  .DESCRIPTION
  Invoke-vtsFastDownload is a function that downloads a file from a specified URL to a specified path on the local system. 
  It first checks if the download path exists, if not, it creates it. 
  Then it sets the current location to the download path and sets the execution policy and security protocol. 
  It installs Chocolatey if not already installed and then installs aria2 using Chocolatey. 
  Finally, it downloads the file using aria2 and saves it at the specified location.
  
  .PARAMETER DownloadPath
  The path where the downloaded file will be saved. Default is "C:\temp".
  
  .PARAMETER URL
  The URL of the file to be downloaded. This is a mandatory parameter.
  
  .PARAMETER FileName
  The name of the file to be saved on the local system. This is a mandatory parameter.
  
  .EXAMPLE
  Invoke-vtsFastDownload -URL "http://example.com/file.zip" -FileName "file.zip"
  This will download the file from the specified URL and save it as "file.zip" in the default download path "C:\temp".
  
  .EXAMPLE
  Invoke-vtsFastDownload -DownloadPath "D:\downloads" -URL "http://example.com/file.zip" -FileName "file.zip"
  This will download the file from the specified URL and save it as "file.zip" in the specified download path "D:\downloads".
  
  .LINK
  Utilities
  #>
  param (
    [string]$DownloadPath = "C:\temp",
    [Parameter(Mandatory = $true)]
    [string]$URL,
    [Parameter(Mandatory = $true)]
    [string]$FileName
  )

  Write-Host "Checking if the download path $DownloadPath exists..."
  if (!(Test-Path $DownloadPath)) {
    Write-Host "Download path does not exist. Creating it now..."
    New-Item -ItemType Directory -Force -Path $DownloadPath
  }
  else {
    Write-Host "Download path exists. Proceeding with the download..."
  }

  Write-Host "Setting the current location to $DownloadPath..."
  Set-Location -Path $DownloadPath

  Write-Host "Setting execution policy and security protocol..."
  Set-ExecutionPolicy Bypass -Scope Process -Force
  [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072

  Write-Host "Installing Chocolatey..."
  Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

  Write-Host "Installing aria2 with Chocolatey..."
  & choco install aria2 -y

  Write-Host "`nDownloading file from $URL to $(Join-Path $DownloadPath $FileName)" -ForegroundColor Green
  & aria2c -x16 -s16 -k1M -c -o "$FileName" "$URL" --file-allocation=none

  Write-Host "`nDownload complete. File saved at $(Join-Path $DownloadPath $FileName)" -ForegroundColor Green
}

