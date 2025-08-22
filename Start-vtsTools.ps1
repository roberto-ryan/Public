param(
  [string]$FunctionsPath = 'C:\Users\Administrator\source\Public\functions'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Console {
  if (-not ($Host.Name -match 'ConsoleHost|Visual Studio Code')) {
    Write-Host 'This script is designed for an interactive console host.' -ForegroundColor Yellow
  }
  try { [void][Console]::CursorVisible } catch {
    throw "This host doesn't expose [Console]. Try running in Windows Terminal/PowerShell console."
  }
}

function Safe-Substring([string]$s,[int]$len) {
  if ([string]::IsNullOrEmpty($s)) { return '' }
  if ($s.Length -le $len) { return $s }
  return ($s.Substring(0,[Math]::Max(0,$len-3)) + '...')
}

function TruncatePad([string]$s,[int]$width){
  $s = ($s -replace '\s+',' ').Trim()
  $t = Safe-Substring $s $width
  return $t + (' ' * [Math]::Max(0, $width - ($t).Length))
}

function Read-Key { [Console]::ReadKey($true) }

function Draw-Line([int]$y,[int]$x,[int]$width,[ConsoleColor]$fg,[string]$ch=' ') {
  try {
    if ($width -le 0) { return }
    [Console]::SetCursorPosition([Math]::Max(0,$x),[Math]::Max(0,$y))
    $s = $ch * $width
    Write-Host $s -NoNewline -ForegroundColor $fg
  } catch { }
}

function Write-At([int]$y,[int]$x,[string]$text,[ConsoleColor]$fg,[switch]$NoNewline) {
  try {
    if ($y -ge 0 -and $x -ge 0) { [Console]::SetCursorPosition($x,$y) }
    Write-Host $text -ForegroundColor $fg -NoNewline:$NoNewline
  } catch { }
}

function Box([int]$top,[int]$left,[int]$height,[int]$width,[string]$title,[ConsoleColor]$fg=[ConsoleColor]::Gray) {
  $right = $left + $width - 1
  $bottom = $top + $height - 1
  # Top
  Write-At $top $left "+" $fg
  Draw-Line $top ($left+1) ($width-2) $fg "-"
  Write-At $top $right "+" $fg
  # Sides
  for ($y = $top+1; $y -lt $bottom; $y++) {
    Write-At $y $left "|" $fg
    Write-At $y $right "|" $fg
  }
  # Bottom
  Write-At $bottom $left "+" $fg
  Draw-Line $bottom ($left+1) ($width-2) $fg "-"
  Write-At $bottom $right "+" $fg
  if ($title) {
    $titleClean = Safe-Substring $title ($width-4)
    Write-At $top ($left+2) " $titleClean " $fg
  }
}

function Wrap([string]$t,[int]$w) {
  if (-not $t) { return @() }
  $t = ($t -replace '\s+',' ').Trim()
  $out = @()
  while ($t.Length -gt $w) {
    $break = $t.LastIndexOf(' ', [Math]::Min($w, $t.Length-1))
    if ($break -lt 1) { $break = $w }
    $out += $t.Substring(0,$break)
    $t = $t.Substring($break).Trim()
  }
  if ($t) { $out += $t }
  return $out
}

function Wrap-Block([string]$text,[int]$w) {
  if (-not $text) { return @() }
  $nl = ($text -replace "\r\n","`n") -replace "\r","`n"
  $paras = $nl -split "`n`n+"  # split on blank lines
  $out = @()
  foreach ($p in $paras) {
    $p2 = $p.Trim()
    if (-not $p2) { $out += ''; continue }
    $out += (Wrap $p2 $w)
    $out += ''
  }
  if ($out.Count -gt 0 -and $out[-1] -eq '') { $out = $out[0..($out.Count-2)] }
  return $out
}

function Wrap-Bullet([string]$text,[int]$w,[string]$prefix=' - ',[string]$cont='   '){
  if (-not $text) { return @() }
  $inner = [Math]::Max(8, $w - $prefix.Length)
  $wrapped = @(Wrap $text $inner)  # force array to avoid .Count on scalars
  $out = @()
  $first = $true
  foreach ($seg in $wrapped) {
    $out += ($(if($first){$prefix}else{$cont}) + $seg)
    $first = $false
  }
  return $out
}

function Normalize-Category([object]$c){
  # Handle arrays: pick first non-empty string
  if ($c -is [System.Collections.IEnumerable] -and -not ($c -is [string])) {
    foreach ($e in $c) {
      $n = Normalize-Category $e
      if ($n -and $n -ne 'Uncategorized') { return $n }
    }
    return 'Uncategorized'
  }
  try { $s = if ($null -eq $c) { '' } else { [string]$c } } catch { $s = '' }
  # Strip zero-width and non-breaking spaces, BOM
  $s = $s -replace "[\u200B\uFEFF\u00A0]", ''
  # collapse whitespace and trim
  $s = ($s -replace '\s+', ' ').Trim()
  if ([string]::IsNullOrWhiteSpace($s)) { return 'Uncategorized' }
  if ($s -match '^(?i)uncategorized$') { return 'Uncategorized' }
  return $s
}

function Prompt-String([string]$label,[string]$default=''){
  $prompt = if ($default) { "$label [$default]" } else { $label }
  Read-Host $prompt
}

function Try-Cast([string]$value,[string]$typeName,[ref]$casted){
  if ($null -eq $value) { $casted.Value = $null; return $true }
  if ($value -eq '' ) { $casted.Value = $null; return $true }
  $tn = "$typeName"
  try {
    # Arrays: split by comma and cast elements to inner type
    if ($tn -match '^((System\.)?\w+)\[\]$') {
      $elemType = $Matches[1]
      $parts = $value -split '\s*,\s*'
      $arr = @()
      foreach ($part in $parts) {
        $tmp = $null
        if (-not (Try-Cast $part $elemType ([ref]$tmp))) { $tmp = $part }
        $arr += ,$tmp
      }
      $casted.Value = $arr
      return $true
    }
    if ($tn -match '^(System\.)?Int16$') { $casted.Value = [int16]$value; return $true }
    elseif ($tn -match '^(System\.)?Int32$') { $casted.Value = [int]$value; return $true }
    elseif ($tn -match '^(System\.)?Int64$') { $casted.Value = [long]$value; return $true }
    elseif ($tn -match '^(System\.)?UInt16$') { $casted.Value = [uint16]$value; return $true }
    elseif ($tn -match '^(System\.)?UInt32$') { $casted.Value = [uint32]$value; return $true }
    elseif ($tn -match '^(System\.)?UInt64$') { $casted.Value = [uint64]$value; return $true }
    elseif ($tn -match '^(System\.)?Double$') { $casted.Value = [double]$value; return $true }
    elseif ($tn -match '^(System\.)?Single$') { $casted.Value = [single]$value; return $true }
    elseif ($tn -match '^(System\.)?Decimal$') { $casted.Value = [decimal]$value; return $true }
    elseif ($tn -match '^(System\.)?Boolean$') { $casted.Value = [bool]$value; return $true }
    elseif ($tn -match '^(System\.)?Date(Time)?$') { $casted.Value = [datetime]$value; return $true }
    elseif ($tn -match '^(System\.)?String$') { $casted.Value = [string]$value; return $true }
    elseif ($tn -match '^(System\.)?Management\.Automation\.PSCredential$') {
      $cred = Get-Credential -Message "Enter PSCredential for parameter"
      $casted.Value = $cred
      return $true
    }
    else { $casted.Value = $value; return $true }
  } catch {
    return $false
  }
}

function Load-ModuleFromFolder([string]$path){
  if (-not (Test-Path $path)) { throw "FunctionsPath not found: $path" }
  $files = Get-ChildItem -Path $path -Filter *.ps1 -File | Sort-Object Name
  if (-not $files) { throw "No .ps1 files found in $path" }
  $sb = {
    param($files)
    foreach ($f in $files) {
      . $f.FullName
    }
    Export-ModuleMember -Function *
  }
  $module = New-Module -Name VTS.Dynamic -ScriptBlock $sb -ArgumentList (,$files)
  Import-Module $module -Force -Global
  return $module
}

function Get-CommandMeta([System.Management.Automation.CommandInfo]$cmd){
  $help = $null
  try { $help = Get-Help $cmd.Name -ErrorAction Stop } catch { }
  $category = 'Uncategorized'
  if ($help -and $help.RelatedLinks) {
    try {
      # Use .LINK text as category (Install-vtsTools.ps1 does this)
      $lnk = $help.RelatedLinks | Select-Object -ExpandProperty NavigationLink -ErrorAction SilentlyContinue
      if ($lnk) {
        $txt = $lnk | Select-Object -ExpandProperty LinkText -ErrorAction SilentlyContinue
        if ($txt) { $category = $txt | Select-Object -First 1 }
      }
    } catch { }
  }
  $desc = ''
  if ($help -and $help.Description) {
    try { $desc = ($help.Description | Select-Object -ExpandProperty Text) -join ' ' } catch { }
  }
  $synopsis = ''
  if ($help) { try { $synopsis = $help.Synopsis } catch { } }
  $examples = @()
  if ($help -and $help.Examples) {
    foreach ($ex in @($help.Examples.Example)) {
      if ($null -ne $ex) {
        $code = $ex.Code
        if (-not $code) { $code = $ex.Title }
        if ($code) { $examples += ($code -replace '^\s+|\s+$','') }
      }
    }
  }
  $params = @()
  if ($help -and $help.Parameters) {
    foreach ($p in @($help.Parameters.Parameter)) {
      $pName = $p.Name
      $pRequired = [bool]$p.Required
      $pTypeName = 'String'
      if ($p.Type -and $p.Type.Name) { $pTypeName = $p.Type.Name }
      $pPosition = [int]::MinValue
      if ($null -ne $p.Position) {
        try { $pPosition = [int]$p.Position } catch { $pPosition = [int]::MinValue }
      }

      $params += [pscustomobject]@{
        Name          = $pName
        Required      = $pRequired
        Type          = $pTypeName
        Position      = $pPosition
      }
    }
  }

  $synopsisVal = $synopsis
  if (-not $synopsisVal) { $synopsisVal = $desc }
  $descVal = $desc
  if (-not $descVal) { $descVal = $synopsis }

  [pscustomobject]@{
    Name        = $cmd.Name
    Category    = $category
    Synopsis    = $synopsisVal
    Description = $descVal
    Examples    = $examples
    Parameters  = $params
  }
}

function Parse-HelpFromFile([string]$filePath){
  $syn = ''
  $desc = ''
  $examples = @()
  $category = 'Uncategorized'
  $lines = Get-Content -Path $filePath -ErrorAction SilentlyContinue -TotalCount 400
  if (-not $lines) { return [pscustomobject]@{ Synopsis=''; Description=''; Examples=@(); Category='Uncategorized' } }
  $inHelp = $false
  $curTag = ''
  $buf = @()

  foreach ($raw in $lines) {
    $line = "$raw"
    if (-not $inHelp) {
      if ($line -match '^\s*<#') { $inHelp = $true }
      continue
    }
    if ($line -match '#>') {
      # flush
      $text = ($buf -join " `n").Trim()
      if ($curTag -eq 'SYNOPSIS' -and $text) { $syn = $text }
      elseif ($curTag -eq 'DESCRIPTION' -and $text) { $desc = $text }
      elseif ($curTag -eq 'EXAMPLE' -and $text) { $examples += $text }
      elseif ($curTag -eq 'LINK' -and $text) {
        $firstLine = ($text -split "`n")[0].Trim()
        $cand = $firstLine
        if ($cand -match '^(https?://|www\.|mailto:|file:)') { $cand = '' }
        if (-not $cand) {
          foreach ($ln in ($text -split "`n")) {
            $tln = $ln.Trim()
            if ($tln -and ($tln -notmatch '^(https?://|www\.|mailto:|file:)')) { $cand = $tln; break }
          }
        }
        if ($cand) { $category = $cand }
      }
      break
    }
    if ($line -match '^\s*\.(\w+)\b') {
      # flush previous tag
      $text = ($buf -join " `n").Trim()
      if ($curTag -eq 'SYNOPSIS' -and $text) { $syn = $text }
      elseif ($curTag -eq 'DESCRIPTION' -and $text) { $desc = $text }
      elseif ($curTag -eq 'EXAMPLE' -and $text) { $examples += $text }
      elseif ($curTag -eq 'LINK' -and $text) {
        $firstLine = ($text -split "`n")[0].Trim()
        $cand = $firstLine
        if ($cand -match '^(https?://|www\.|mailto:|file:)') { $cand = '' }
        if (-not $cand) {
          foreach ($ln in ($text -split "`n")) {
            $tln = $ln.Trim()
            if ($tln -and ($tln -notmatch '^(https?://|www\.|mailto:|file:)')) { $cand = $tln; break }
          }
        }
        if ($cand) { $category = $cand }
      }
      # start new tag
      $curTag = $Matches[1].ToUpperInvariant()
      $buf = @()
      $content = ($line -replace '^\s*\.\w+\s*','').Trim()
      if ($content) { $buf += $content }
      continue
    }
    if ($curTag) { $buf += $line }
  }
  if (-not $syn) { $syn = $desc }
  if (-not $desc) { $desc = $syn }
  [pscustomobject]@{ Synopsis=$syn; Description=$desc; Examples=$examples; Category=$category }
}

function Get-ParameterMeta([System.Management.Automation.CommandInfo]$cmd){
  $arr = @()
  foreach ($kv in $cmd.Parameters.GetEnumerator()) {
    $p = $kv.Value
    $name = $p.Name
  $typeName = if ($p.ParameterType) { $p.ParameterType.Name } else { 'String' }
    $pos = [int]::MinValue
    $required = $false
    $attrs = @()
    try { $attrs = @($p.Attributes) } catch { $attrs = @() }
    $paramAttr = $null
    foreach ($a in $attrs) {
      if ($a -is [System.Management.Automation.ParameterAttribute]) { $paramAttr = $a; break }
    }
    if ($paramAttr) {
  try { if ($null -ne $paramAttr.Position) { $pos = [int]$paramAttr.Position } } catch { $pos = [int]::MinValue }
      try { $required = [bool]$paramAttr.Mandatory } catch { $required = $false }
    }
  $arr += [pscustomobject]@{ Name=$name; Required=$required; Type=$typeName; Position=$pos }
  }
  return ($arr | Sort-Object Position, Name)
}

function Filter-UserParameters([object[]]$params){
  $common = @(
    'Verbose','Debug','ErrorAction','ErrorVariable','WarningAction','WarningVariable',
    'InformationAction','InformationVariable','OutVariable','OutBuffer','PipelineVariable',
    'WhatIf','Confirm'
  )
  $params |
    Where-Object { $_ -and $_.PSObject -and $_.PSObject.Properties['Name'] } |
    Where-Object { $name = $_.Name; $name -and -not ($common -contains $name) }
}

function Build-Model([string]$functionsPath){
  Load-ModuleFromFolder -path $functionsPath
  $files = Get-ChildItem -Path $functionsPath -Filter *.ps1 -File | Sort-Object Name
  $items = @()
  foreach ($f in $files) {
    $fn = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    $cmd = Get-Command -Name $fn -ErrorAction SilentlyContinue
    if (-not $cmd) { continue }
    $help = Parse-HelpFromFile -filePath $f.FullName
  $params = Get-ParameterMeta -cmd $cmd
  $params = @(Filter-UserParameters -params $params)
    $items += [pscustomobject]@{
      Name        = $fn
      Category    = (Normalize-Category $help.Category)
      Synopsis    = $help.Synopsis
      Description = $help.Description
      Examples    = $help.Examples
      Parameters  = $params
    }
  }
  # Normalize to ensure Category is always a single string
  $norm = foreach ($it in $items) {
    [pscustomobject]@{
      Name        = $it.Name
      Category    = (Normalize-Category $it.Category)
      Synopsis    = $it.Synopsis
      Description = $it.Description
      Examples    = $it.Examples
      Parameters  = $it.Parameters
    }
  }
  $groups = $norm | Group-Object Category | Sort-Object Name
  $model = foreach ($g in $groups) { [pscustomobject]@{ Category=$g.Name; Commands=($g.Group | Sort-Object Name) } }
  return $model
}

function Prompt-And-Run([pscustomobject]$cmdMeta){
  [Console]::Clear()
  Write-Host ("Running: " + $cmdMeta.Name) -ForegroundColor Cyan
  Write-Host ""

  # Order: required first (by position if any), then optional
  $paramListAll = @($cmdMeta.Parameters)
  $paramListAll = $paramListAll |
    Where-Object { $_ -and $_.PSObject -and $_.PSObject.Properties['Name'] } |
    ForEach-Object { $_ }
  # Ensure arrays even when only one element survives sorting
  $req = @(( $paramListAll | Where-Object { $_.Required } | Sort-Object Position ))
  $opt = @(( $paramListAll | Where-Object { -not $_.Required } | Sort-Object Name ))
  $ordered = @($req + $opt)

  $splat = @{}
  foreach ($p in $ordered) {
    $pTypeName = 'String'
    try {
      if ($null -ne $p) {
        if ($p.PSObject -and $p.PSObject.Properties['Type']) {
          $pTypeName = "$($p.Type)"
        } elseif ($p.PSObject -and $p.PSObject.Properties['ParameterType']) {
          $pTypeName = if ($p.ParameterType -and $p.ParameterType.Name) { $p.ParameterType.Name } else { "$($p.ParameterType)" }
        }
      }
    } catch { $pTypeName = 'String' }
  # Ensure we have a valid parameter name; skip otherwise
  $pNameVal = if ($p -and $p.PSObject.Properties['Name']) { "$($p.Name)" } else { $null }
  if (-not $pNameVal) { continue }
  $label = "{0} ({1}{2})" -f $pNameVal, $pTypeName, ($(if($p.Required) {'; required'} else {''}))
    if ($pTypeName -match '^(SwitchParameter|Switch)$') {
      $ans = Prompt-String "$label - toggle [y/N]" ''
  if ($ans -match '^(y|yes|true|1)$') { $splat[$pNameVal] = $true }
    }
    elseif ($pTypeName -match 'PSCredential') {
      $ans = Prompt-String "$label - press Enter to prompt for credential or leave blank to skip" ''
      if ($ans -ne '') {
        $cast = $null
  if (Try-Cast $ans $pTypeName ([ref]$cast)) { $splat[$pNameVal] = $cast }
      } else {
        $cred = Get-Credential -Message "Enter PSCredential for -$($p.Name) (Esc to cancel)"
  if ($cred) { $splat[$pNameVal] = $cred }
      }
    }
    else {
      $ans = Prompt-String $label ''
      if ($ans -ne '') {
        $cast = $null
        if (Try-Cast $ans $pTypeName ([ref]$cast)) { $splat[$pNameVal] = $cast }
        else { $splat[$pNameVal] = $ans }
      } elseif ($p.Required) {
        Write-Host "Parameter -$($p.Name) is required. Please enter a value." -ForegroundColor Yellow
        $ans = Prompt-String $label ''
        if ($ans -ne '') {
          $cast = $null
          if (Try-Cast $ans $pTypeName ([ref]$cast)) { $splat[$pNameVal] = $cast } else { $splat[$pNameVal] = $ans }
        }
      }
    }
  }

  Write-Host ""
  Write-Host "Command preview:" -ForegroundColor DarkGray
  $preview = $cmdMeta.Name
  foreach ($k in $splat.Keys) {
    $v = $splat[$k]
    if ($v -is [switch] -or ($v -is [bool] -and $v -eq $true)) {
      $preview += " -$k"
    } elseif ($v -is [securestring]) {
      $preview += " -$k ********"
    } elseif ($v -is [string[]]) {
      $preview += " -$k `"$($v -join ',')`""
    } else {
      $preview += " -$k `"$v`""
    }
  }
  Write-Host "`n$preview`n" -ForegroundColor Green

  $confirm = Read-Host "Press Enter to run, or type 'n' to cancel"
  if ($confirm -match '^(n|no)$') { return }

  # Execute the command in a 'raw' way: relax strict mode and error preferences
  # so the called functions behave like when run standalone.
  Write-Host ""
  $prevEA = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try {
    try { Set-StrictMode -Off } catch { }
  & (Get-Command $cmdMeta.Name) @splat | Out-Host
  } catch {
    # Surface the original error without extra TUI formatting
    Write-Error $_
  } finally {
    $ErrorActionPreference = $prevEA
    try { Set-StrictMode -Version Latest } catch { }
  }
  Write-Host ""
  Write-Host "Press any key to return..." -ForegroundColor DarkGray
  [void](Read-Key)
}

function Run-TUI([object[]]$model){
  # Cache model groups once
  $cats = @($model)
  $catIdx = 0
  $cmdIdx = 0
  $focus = 'cat' # cat | cmd | detail
  $scrollCat = 0
  $scrollCmd = 0
  $scrollDetail = 0
  $detailLines = @()

  # Track size and dirty state
  $lastW = -1; $lastH = -1
  $needFull = $true
  $dirtyCats = $true
  $dirtyCmds = $true
  $dirtyDetails = $true
  $prevCatIdx = -1
  $prevCmdIdx = -1

  while ($true) {
    $w = [Console]::WindowWidth
    $h = [Console]::WindowHeight
    if ($w -ne $lastW -or $h -ne $lastH) {
      $lastW = $w; $lastH = $h
      $needFull = $true
      $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true
    }

  $header = " VTS Tools Browser "
  $footer = " Up/Down/Left/Right: Navigate/Scroll | Enter: Run "

    # Layout
    $paneTop = 1
    $paneHeight = $h - 2
    $leftW = [Math]::Max(18, [Math]::Min(30, [int]($w*0.22)))
    $midW  = [Math]::Max(24, [Math]::Min(42, [int]($w*0.28)))
    $rightW = $w - $leftW - $midW - 4

    if ($needFull) {
      [Console]::Clear()
      # Header and footer
  Draw-Line 0 0 $w ([ConsoleColor]::DarkCyan) ' '
  Write-At 0 0 (TruncatePad $header $w) ([ConsoleColor]::Black) -NoNewline
  Draw-Line ($h-1) 0 $w ([ConsoleColor]::DarkCyan) ' '
  Write-At ($h-1) 0 (TruncatePad $footer $w) ([ConsoleColor]::Black) -NoNewline
      # Boxes
      Box $paneTop 0 $paneHeight $leftW "Categories" ([ConsoleColor]::DarkGray)
      Box $paneTop ($leftW+1) $paneHeight $midW "Commands" ([ConsoleColor]::DarkGray)
      Box $paneTop ($leftW+$midW+2) $paneHeight $rightW "Details" ([ConsoleColor]::DarkGray)
    }

    if (-not $cats) {
      Write-At ([int]($paneTop+2)) 2 "No commands found." ([ConsoleColor]::Yellow)
      Write-At ($h-2) 2 "Press any key to refresh." ([ConsoleColor]::DarkGray)
      [void](Read-Key)
      $needFull = $true; $dirtyCats=$true; $dirtyCmds=$true; $dirtyDetails=$true
      continue
    }

    if ($catIdx -ge $cats.Count) { $catIdx = 0 }
    $selectedCat = $cats[$catIdx]
    $cmds = @()
    if ($selectedCat -and ($selectedCat.PSObject.Properties['Commands'])) { $cmds = @($selectedCat.Commands) }
    if ($cmdIdx -ge $cmds.Count) { $cmdIdx = 0 }

    # Determine dirty by selection change
  if ($prevCatIdx -ne $catIdx) { $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true; $scrollDetail = 0; $prevCatIdx = $catIdx }
  if ($prevCmdIdx -ne $cmdIdx) { $dirtyCmds = $true; $dirtyDetails = $true; $scrollDetail = 0; $prevCmdIdx = $cmdIdx }

    # Render categories when dirty
    if ($dirtyCats -or $needFull) {
      $listTop = $paneTop + 1
      $listHeight = $paneHeight - 2
      if ($catIdx -lt $scrollCat) { $scrollCat = $catIdx }
      if ($catIdx -ge ($scrollCat + $listHeight)) { $scrollCat = $catIdx - $listHeight + 1 }

      for ($i=0; $i -lt $listHeight; $i++) {
        $idx = $scrollCat + $i
        $y = $listTop + $i
        $line = ' ' * ($leftW - 2)
        if ($idx -lt $cats.Count) {
          $catItem = $cats[$idx]
          $name = if ($catItem -and ($catItem.PSObject.Properties['Category'])) { Normalize-Category $catItem.Category } else { 'Uncategorized' }
          $count = if ($catItem -and ($catItem.PSObject.Properties['Commands'])) { @($catItem.Commands).Count } else { 0 }
          $txt = ("{0} ({1})" -f $name, $count)
          $line = TruncatePad $txt ($leftW - 2)
        }
        $isSel = ($idx -eq $catIdx)
        if ($isSel -and $focus -eq 'cat') { $fg = [ConsoleColor]::Black; $bg = [ConsoleColor]::Cyan }
        else { $fg = [ConsoleColor]::Gray;  $bg = [ConsoleColor]::Black }
        [Console]::SetCursorPosition(1, $y); [Console]::ForegroundColor = $fg; [Console]::BackgroundColor = $bg; [Console]::Write($line); [Console]::ResetColor()
      }
      # Draw scroll indicators (far-right column) when overflow exists
      $xIndCat = ($leftW - 2)
      if ($scrollCat -gt 0 -and $cats.Count -gt 0) {
        $idxTop = $scrollCat
        $isSelTop = ($idxTop -eq $catIdx -and $focus -eq 'cat')
        $bgTop = if ($isSelTop) { [ConsoleColor]::Cyan } else { [ConsoleColor]::Black }
        [Console]::SetCursorPosition($xIndCat, $listTop)
        [Console]::BackgroundColor = $bgTop; [Console]::ForegroundColor = [ConsoleColor]::Green; [Console]::Write("^"); [Console]::ResetColor()
      }
      if ( ($scrollCat + $listHeight) -lt $cats.Count ) {
        $idxBot = [Math]::Min($cats.Count-1, $scrollCat + $listHeight - 1)
        $isSelBot = ($idxBot -eq $catIdx -and $focus -eq 'cat')
        $bgBot = if ($isSelBot) { [ConsoleColor]::Cyan } else { [ConsoleColor]::Black }
        [Console]::SetCursorPosition($xIndCat, ($listTop + $listHeight - 1))
        [Console]::BackgroundColor = $bgBot; [Console]::ForegroundColor = [ConsoleColor]::Green; [Console]::Write("v"); [Console]::ResetColor()
      }
      $dirtyCats = $false
    }

    # Render commands when dirty
    if ($dirtyCmds -or $needFull) {
      $cmdTop = $paneTop + 1
      $cmdHeight = $paneHeight - 2
      if ($cmdIdx -lt $scrollCmd) { $scrollCmd = $cmdIdx }
      if ($cmdIdx -ge ($scrollCmd + $cmdHeight)) { $scrollCmd = $cmdIdx - $cmdHeight + 1 }

      for ($i=0; $i -lt $cmdHeight; $i++) {
        $idx = $scrollCmd + $i
        $y = $cmdTop + $i
        $line = ' ' * ($midW - 2)
        if ($idx -lt $cmds.Count) {
          $txt = $cmds[$idx].Name
          $line = TruncatePad $txt ($midW - 2)
        }
        $isSel = ($idx -eq $cmdIdx)
        if ($isSel -and $focus -eq 'cmd') { $fg = [ConsoleColor]::Black; $bg = [ConsoleColor]::Green }
        else { $fg = [ConsoleColor]::Gray;  $bg = [ConsoleColor]::Black }
        [Console]::SetCursorPosition($leftW+2, $y); [Console]::ForegroundColor = $fg; [Console]::BackgroundColor = $bg; [Console]::Write($line); [Console]::ResetColor()
      }
      # Draw scroll indicators (far-right column) when overflow exists
      $xIndCmd = ($leftW + $midW - 1)
      if ($scrollCmd -gt 0 -and $cmds.Count -gt 0) {
        $idxTop = $scrollCmd
        $isSelTop = ($idxTop -eq $cmdIdx -and $focus -eq 'cmd')
        $bgTop = if ($isSelTop) { [ConsoleColor]::Green } else { [ConsoleColor]::Black }
        [Console]::SetCursorPosition($xIndCmd, $cmdTop)
        [Console]::BackgroundColor = $bgTop; [Console]::ForegroundColor = [ConsoleColor]::Green; [Console]::Write("^"); [Console]::ResetColor()
      }
      if ( ($scrollCmd + $cmdHeight) -lt $cmds.Count ) {
        $idxBot = [Math]::Min($cmds.Count-1, $scrollCmd + $cmdHeight - 1)
        $isSelBot = ($idxBot -eq $cmdIdx -and $focus -eq 'cmd')
        $bgBot = if ($isSelBot) { [ConsoleColor]::Green } else { [ConsoleColor]::Black }
        [Console]::SetCursorPosition($xIndCmd, ($cmdTop + $cmdHeight - 1))
        [Console]::BackgroundColor = $bgBot; [Console]::ForegroundColor = [ConsoleColor]::Green; [Console]::Write("v"); [Console]::ResetColor()
      }
      $dirtyCmds = $false
    }

    # Render details when dirty
    if ($dirtyDetails -or $needFull) {
      $detail = if ($cmds.Count) { $cmds[$cmdIdx] } else { $null }
      $dTop = $paneTop + 1
      $dHeight = $paneHeight - 2
      $wrapWidth = $rightW - 3
      # Clear details area
      for ($yy = 0; $yy -lt $dHeight; $yy++) {
        [Console]::SetCursorPosition($leftW+$midW+3, $dTop+$yy)
        [Console]::Write((' ' * $wrapWidth))
      }
  $lineY = $dTop
      if ($detail) {
        $lines = @()
        # Header
        $lines += "Name: $($detail.Name)"
        $lines += "Category: $($detail.Category)"
        # Synopsis
        if ($detail.Synopsis) {
          $lines += ''
          $lines += 'Synopsis:'
          $lines += @(Wrap-Block $detail.Synopsis $wrapWidth)
        }
        # Description
        if ($detail.Description) {
          $lines += ''
          $lines += 'Description:'
          $lines += @(Wrap-Block $detail.Description $wrapWidth)
        }
        # Parameters
        $paramList = @($detail.Parameters)
        if ($paramList.Count) {
          $lines += ''
          $lines += 'Parameters:'
          foreach ($p in $paramList | Sort-Object { -not $_.Required }, Position, Name) {
            $req = if ($p.Required) { 'required' } else { 'optional' }
            $pline = "$($p.Name) <$($p.Type)> ($req)"
            $lines += @(Wrap-Bullet $pline $wrapWidth ' - ' '   ')
          }
        }
        # Examples
        $exList = @($detail.Examples)
        if ($exList.Count) {
          $lines += ''
          $lines += 'Examples:'
          foreach ($ex in $exList | Select-Object -First 3) {
            $lines += @(Wrap-Bullet $ex $wrapWidth ' - ' '   ')
          }
        }
        # Update cached lines and clamp scroll
        $detailLines = @($lines)
        $maxScroll = [Math]::Max(0, $detailLines.Count - $dHeight)
        if ($scrollDetail -gt $maxScroll) { $scrollDetail = $maxScroll }
        if ($scrollDetail -lt 0) { $scrollDetail = 0 }

        # Render a visible slice with section highlighting
        $rendered = 0
        for ($i = $scrollDetail; $i -lt $detailLines.Count -and $rendered -lt $dHeight; $i++) {
          $l = $detailLines[$i]
          $color = [ConsoleColor]::Gray
          if ($l -like 'Name:*' -or $l -like 'Category:*' -or $l -in @('Synopsis:','Description:','Parameters:','Examples:')) { $color = [ConsoleColor]::White }
          Write-At $lineY ($leftW+$midW+3) (TruncatePad $l $wrapWidth) $color
          $lineY++
          $rendered++
        }
        # Draw scroll indicators at far-right column if overflow
        $xInd = ($leftW+$midW+3) + ($wrapWidth - 1)
        if ($scrollDetail -gt 0) {
          # Up indicator at first visible line
          Write-At $dTop $xInd "^" ([ConsoleColor]::Green)
        }
        if (($scrollDetail + $dHeight) -lt $detailLines.Count) {
          # Down indicator at last visible line
          Write-At ($dTop + $dHeight - 1) $xInd "v" ([ConsoleColor]::Green)
        }
      } else {
        Write-At $dTop ($leftW+$midW+3) "No commands in this category." ([ConsoleColor]::Yellow)
      }
      $dirtyDetails = $false
    }

    $needFull = $false

    # Key handling
    $k = Read-Key
  switch ($k.Key) {
  'Enter' {
        $detail = $null
        $selectedCat = if ($cats.Count) { $cats[$catIdx] } else { $null }
        if ($selectedCat -and ($selectedCat.PSObject.Properties['Commands'])) { $cmds = @($selectedCat.Commands) } else { $cmds = @() }
        if ($cmds.Count) { $detail = $cmds[$cmdIdx] }
        if ($detail) { Prompt-And-Run $detail }
        $needFull = $true; $dirtyCats=$true; $dirtyCmds=$true; $dirtyDetails=$true
      }
      'UpArrow' {
        switch ($focus) {
          'cat' { if ($catIdx -gt 0) { $catIdx--; $cmdIdx = 0 } }
          'cmd' { if ($cmdIdx -gt 0) { $cmdIdx-- } }
          'detail' { if ($scrollDetail -gt 0) { $scrollDetail--; $dirtyDetails = $true } }
        }
      }
      'DownArrow' {
        switch ($focus) {
          'cat' { if ($catIdx -lt ($cats.Count-1)) { $catIdx++; $cmdIdx = 0 } }
          'cmd' { if ($cmdIdx -lt ($cmds.Count-1)) { $cmdIdx++ } }
          'detail' {
            $dHeight = $paneHeight - 2
            $maxScroll = [Math]::Max(0, @($detailLines).Count - $dHeight)
            if ($scrollDetail -lt $maxScroll) { $scrollDetail++; $dirtyDetails = $true }
          }
        }
      }
      'LeftArrow' {
        if ($focus -eq 'cmd') { $focus = 'cat' }
        elseif ($focus -eq 'detail') { $focus = 'cmd' }
        $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true
      }
      'RightArrow' {
        if ($focus -eq 'cat') { $focus = 'cmd' }
        elseif ($focus -eq 'cmd') { $focus = 'detail' }
        $dirtyCats = $true; $dirtyCmds = $true; $dirtyDetails = $true
      }
      default { }
    }
  }
}

try {
  Write-Host "[DEBUG] Starting Ensure-Console" -ForegroundColor DarkGray
  Ensure-Console
  Write-Host "[DEBUG] Ensure-Console OK" -ForegroundColor DarkGray
  Write-Host "[DEBUG] Building model from $FunctionsPath" -ForegroundColor DarkGray
  $model = Build-Model -functionsPath $FunctionsPath
  $catCount = if ($null -eq $model) { 0 } else { @($model).Count }
  Write-Host "[DEBUG] Model built: $catCount categories" -ForegroundColor DarkGray
  Write-Host "[DEBUG] Launching TUI" -ForegroundColor DarkGray
  Run-TUI -model $model
} catch {
  Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
  if ($_.InvocationInfo.PositionMessage) {
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor DarkGray
  }
  exit 1
}