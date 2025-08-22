param(
  [string]$Branch = 'main'
)


Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

function Write-Info($msg){ Write-Host $msg -ForegroundColor Cyan }
function Write-Warn($msg){ Write-Host $msg -ForegroundColor Yellow }
function Write-Err ($msg){ Write-Host $msg -ForegroundColor Red }

$repoOwner   = 'roberto-ryan'
$repoName    = 'Public'
$zipUrl      = "https://github.com/$repoOwner/$repoName/archive/refs/heads/$Branch.zip"

# Choose install base
if ($env:USERNAME -eq 'SYSTEM') {
  $base = Join-Path $env:SystemDrive 'Tools'
} else {
  $base = Join-Path $env:LOCALAPPDATA 'VTS'
}
if (-not (Test-Path $base)) { New-Item -ItemType Directory -Path $base -Force | Out-Null }

$zipPath  = Join-Path $env:TEMP "$repoName-$Branch.zip"
$destRoot = $base  # Expand-Archive will create $repoName-$Branch under this root
$repoDir  = Join-Path $destRoot "$repoName-$Branch"

try {
  Write-Info "Downloading $zipUrl ..."
  Invoke-WebRequest -Uri $zipUrl -UseBasicParsing -OutFile $zipPath

  if (Test-Path $repoDir) {
    Write-Warn "Removing existing folder: $repoDir"
    Remove-Item -Recurse -Force -Path $repoDir -ErrorAction SilentlyContinue
  }

  Write-Info "Extracting to $destRoot ..."
  Expand-Archive -Path $zipPath -DestinationPath $destRoot -Force

  $startScript = Join-Path $repoDir 'Start-vtsTools.ps1'
  $funcDir     = Join-Path $repoDir 'functions'
  if (-not (Test-Path $startScript)) { throw "Start-vtsTools.ps1 not found in $repoDir" }
  if (-not (Test-Path $funcDir)) { throw "functions folder not found in $repoDir" }

  Write-Info "Launching Start-vtsTools.ps1 ..."
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $startScript -FunctionsPath $funcDir
}
catch {
  Write-Err ("Install/Run failed: " + $_.Exception.Message)
  if ($_.InvocationInfo.PositionMessage) { Write-Warn $_.InvocationInfo.PositionMessage }
  exit 1
}
finally {
  if (Test-Path $zipPath) { Remove-Item -Force $zipPath -ErrorAction SilentlyContinue }
}
