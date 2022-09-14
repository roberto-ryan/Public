[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$moduleURL = "https://raw.githubusercontent.com/roberto-ryan/Public/main/vtsTools.ps1"
$moduleName = "VTS"
$filename = "$moduleName.psm1"

Set-ExecutionPolicy Bypass -Scope Process

if ($env:USERNAME -eq "SYSTEM") {
    $modulePath = "$env:SystemDrive\Tools"
}
else {
    $modulePath = $env:PSModulePath -split ";" |
    Select-String "$env:USERNAME" |
    Select-Object -First 1
}

if (-not (Test-Path $modulePath\$moduleName)) {
    New-Item -Path $modulePath\$moduleName -ItemType Directory -Force |
    Out-Null
}

Remove-Item "$modulePath\$moduleName\$filename" -Force 2>$null
Get-Module VTS | Remove-Module

Invoke-WebRequest -uri $moduleURL -UseBasicParsing |
Select-Object -ExpandProperty Content |
Out-File -FilePath "$modulePath\$moduleName\$filename" -Force
Import-Module $modulePath\$moduleName

$commands = @()

Get-Command -Module VTS |
Select-Object Name, Description |
Sort-Object Name |
ForEach-Object {
    $commands += [pscustomobject]@{
        'Installed Command' = $_.Name
        Description         = (Get-Help $_.Name |
            Select-Object -ExpandProperty Description |
            Select-Object -ExpandProperty Text
        )
        Usage            = (Get-Help $_.Name |
        Select-Object -ExpandProperty Examples |
        Select-Object -ExpandProperty Example |
        Select-Object -ExpandProperty Code
        )
    }
}

$commands | Select-Object 'Installed Command', Usage, Description

"`nType 'help' followed by the command name for more information.

Example: PS> help Get-vtsMappedDrive"