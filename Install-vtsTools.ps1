[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$moduleURL = "https://raw.githubusercontent.com/roberto-ryan/Public/main/vtsTools.ps1"
$moduleName = "VTS"
$filename = "$moduleName.psm1"

# Set execution policy (Windows only - silently skip on Linux/macOS)
try { Set-ExecutionPolicy Bypass -Scope Process -ErrorAction Stop } catch { }

# Determine current username (cross-platform)
$currentUser = if ($env:USERNAME) { $env:USERNAME } elseif ($env:USER) { $env:USER } else { 'unknown' }

# Determine module path (cross-platform)
if ($currentUser -eq 'SYSTEM') {
    $modulePath = "$env:SystemDrive\Tools"
}
else {
    $pathSeparator = if ($IsWindows -or $env:OS -match 'Windows') { ';' } else { ':' }
    $modulePath = $env:PSModulePath -split $pathSeparator |
        Where-Object { $_ -and $currentUser -and $_ -match [regex]::Escape($currentUser) } |
        Select-Object -First 1
    
    # Fallback to user's home-based module path if not found
    if (-not $modulePath) {
        if ($IsWindows -or $env:OS -match 'Windows') {
            $modulePath = Join-Path $env:USERPROFILE 'Documents\PowerShell\Modules'
        } else {
            $modulePath = Join-Path $HOME '.local/share/powershell/Modules'
        }
    }
}

$moduleDir = Join-Path $modulePath $moduleName
if (-not (Test-Path $moduleDir)) {
    New-Item -Path $moduleDir -ItemType Directory -Force | Out-Null
}

$moduleFile = Join-Path $moduleDir $filename
Remove-Item $moduleFile -Force -ErrorAction SilentlyContinue
Get-Module VTS | Remove-Module -ErrorAction SilentlyContinue

Invoke-WebRequest -Uri $moduleURL -UseBasicParsing |
    Select-Object -ExpandProperty Content |
    Out-File -FilePath $moduleFile -Force
Import-Module $moduleDir

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

# LineAcrossScreen

"`nType 'get-help -full' followed by the command name for more information.

Example: PS> get-help -full Get-vtsMappedDrive`n"

$script:windowWidth = (Get-Host).UI.RawUI.WindowSize.Width
Write-Host ('█' * $windowWidth) -ForegroundColor Cyan
""