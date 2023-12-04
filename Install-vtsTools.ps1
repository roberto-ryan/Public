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
        Category            = (Get-help $_.Name |
        Select-Object -ExpandProperty relatedLinks |
        Select-Object -ExpandProperty navigationLink |
        Select-Object -ExpandProperty linkText )
        'Installed Command' = $_.Name
        Description         = (Get-Help $_.Name |
            Select-Object -ExpandProperty Description |
            Select-Object -ExpandProperty Text
        )
    }
}

# $commands | Sort-Object Category, 'Installed Command' | Select-Object 'Installed Command', Category, Description | Format-Table -View Category

$groupedCommands = $commands | Group-Object -Property Category | Sort-Object Name

function LineAcrossScreen {
    $script:windowWidth = (Get-Host).UI.RawUI.WindowSize.Width
    Write-Host ('-' * $windowWidth)
}

LineAcrossScreen

Write-Host "$('█' * ($windowWidth/2 - 6)) VTS Toolkit $('█' * ($windowWidth/2 - 7))" -ForegroundColor Cyan

foreach ($group in $groupedCommands) {
    LineAcrossScreen
    Write-Host "`nCategory: $($group.Name)" -ForegroundColor Yellow
    $group.Group | Format-Table 'Installed Command', Description
}

LineAcrossScreen

"Type 'get-help -full' followed by the command name for more information.

Example: PS> get-help -full Get-vtsMappedDrive"

LineAcrossScreen
""