#Requires -Version 5.1
<#
.SYNOPSIS
    VTS Minimal - Ultra-Clean Blade System
.DESCRIPTION
    A super minimal blade interface with clean teal text on black background
    for the vtsTools PowerShell Toolkit.
.EXAMPLE
    PS> .\vtsToolsGUI_Blades.ps1
.NOTES
    Author: VTS Minimal System
    Version: 7.0.0 MINIMAL EDITION
#>

param(
    [string]$ToolsPath = "$PSScriptRoot\vtsTools.ps1"
)

# Performance optimizations
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::CursorVisible = $false
$Host.UI.RawUI.WindowTitle = "VTS Minimal | Blade System"

# Pre-compile regex patterns for speed
$script:DescriptionRegex = [regex]::new('\.Description\s*\n\s*(.+?)(?=\n\s*\.|\n\s*#>|$)', [System.Text.RegularExpressions.RegexOptions]::Compiled)
$script:LinkRegex = [regex]::new('\.LINK\s*\n\s*(.+?)(?=\n|$)', [System.Text.RegularExpressions.RegexOptions]::Compiled)

# Load vtsTools efficiently
if (-not (Test-Path $ToolsPath)) {
    Write-Host "ERROR: Missing vtsTools.ps1 at $ToolsPath" -ForegroundColor Red
    exit 1
}
. $ToolsPath

# Screen dimensions
$script:Width = $Host.UI.RawUI.WindowSize.Width
$script:Height = $Host.UI.RawUI.WindowSize.Height

# Minimal teal-on-black color scheme
$script:Colors = @{
    Background = [ConsoleColor]::Black
    Text = [ConsoleColor]::Cyan
    Highlight = [ConsoleColor]::White
    Selected = [ConsoleColor]::Black
    Border = [ConsoleColor]::DarkCyan
}

# Blade system state
$script:State = @{
    OpenBlades = @()
    ActiveBlade = 0
    Position = 0
    SubPosition = 0
    Running = $true
    LastUpdate = @{}
    # Store previous positions for each blade type
    BladePositions = @{}
}

# Blade definitions
$script:BladeTypes = @{
    Categories = @{
        Title = "Categories"
        Width = 30
        Icon = "📁"
    }
    Functions = @{
        Title = "Functions"
        Width = 40
        Icon = "⚙️"
    }
    Details = @{
        Title = "Function Details"
        Width = 50
        Icon = "📋"
    }
}

# Pre-load all functions
$script:AllFunctions = @{}
$script:Categories = @()

# Ultra-fast function parser
function Initialize-Functions {
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ToolsPath, [ref]$null, [ref]$null)
    $functionAsts = $ast.FindAll({$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true)
    
    foreach ($func in $functionAsts) {
        $name = $func.Name
        if ($name -notlike "*-vts*") { continue }
        
        $text = $func.Extent.Text
        $category = "Utilities"
        $description = ""
        
        # Fast regex extraction
        $descMatch = $script:DescriptionRegex.Match($text)
        if ($descMatch.Success) {
            $description = $descMatch.Groups[1].Value.Trim()
        }
        
        $linkMatch = $script:LinkRegex.Match($text)
        if ($linkMatch.Success) {
            $category = $linkMatch.Groups[1].Value.Trim()
        }
        
        if (-not $script:AllFunctions.ContainsKey($category)) {
            $script:AllFunctions[$category] = @()
        }
        
        $script:AllFunctions[$category] += @{
            Name = $name
            Description = $description
            Category = $category
        }
    }
    
    $script:Categories = $script:AllFunctions.Keys | Sort-Object
}

# Main rendering function
function Render-Screen {
    # Check if we need to update
    $currentState = "$($script:State.OpenBlades.Count)_$($script:State.ActiveBlade)_$($script:State.Position)_$($script:State.SubPosition)"
    if ($script:State.LastUpdate.State -eq $currentState) {
        return
    }
    
    # Clear screen
    Clear-Host
    
    # Draw background
    Draw-Background
    
    # Draw all open blades with margins
    $x = 2  # Start with left margin
    for ($i = 0; $i -lt $script:State.OpenBlades.Count; $i++) {
        $blade = $script:State.OpenBlades[$i]
        $isActive = ($i -eq $script:State.ActiveBlade)
        
        Draw-Blade -X $x -Blade $blade -IsActive $isActive
        $x += $blade.Width + 3  # Add 3 spaces between blades
    }
    
    # Draw status bar
    Draw-StatusBar
    
    # Update state tracking
    $script:State.LastUpdate.State = $currentState
}

# Draw Azure-style background
function Draw-Background {
    for ($y = 0; $y -lt $script:Height - 1; $y++) {
        [Console]::SetCursorPosition(0, $y)
        Write-Host (" " * $script:Width) -ForegroundColor $script:Colors.Background -BackgroundColor $script:Colors.Background -NoNewline
    }
}

# Draw minimal blade
function Draw-Blade {
    param(
        [int]$X,
        [hashtable]$Blade,
        [bool]$IsActive
    )
    
    $width = $Blade.Width
    $height = $script:Height - 1
    
    # Clean title line
    [Console]::SetCursorPosition($X, 0)
    $titleText = "$($Blade.Icon) $($Blade.Title)"
    if ($IsActive) {
        Write-Host $titleText -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
    } else {
        Write-Host $titleText -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
    }
    
    # Fill rest of title line
    $remaining = $width - $titleText.Length
    if ($remaining -gt 0) {
        Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
    }
    
    # Add blank line between title and content
    [Console]::SetCursorPosition($X, 1)
    Write-Host (" " * $width) -BackgroundColor $script:Colors.Background -NoNewline
    
    # Blade content with margin
    Draw-BladeContent -X $X -Y 2 -Width $width -Height ($height - 2) -Blade $Blade -IsActive $IsActive
}

# Draw blade content based on type
function Draw-BladeContent {
    param(
        [int]$X, [int]$Y, [int]$Width, [int]$Height,
        [hashtable]$Blade, [bool]$IsActive
    )
    
    switch ($Blade.Type) {
        "Categories" {
            Draw-CategoriesBlade -X $X -Y $Y -Width $Width -Height $Height -IsActive $IsActive
        }
        "Functions" {
            Draw-FunctionsBlade -X $X -Y $Y -Width $Width -Height $Height -Blade $Blade -IsActive $IsActive
        }
        "Details" {
            Draw-DetailsBlade -X $X -Y $Y -Width $Width -Height $Height -Blade $Blade -IsActive $IsActive
        }
    }
}

# Draw categories blade
function Draw-CategoriesBlade {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height, [bool]$IsActive)
    
    # Check if we have a Functions blade open to show selected function details
    $functionsBlade = $script:State.OpenBlades | Where-Object { $_.Type -eq "Functions" } | Select-Object -First 1
    $showFunctionDetails = $functionsBlade -and $script:State.OpenBlades.Count -ge 2
    
    if ($showFunctionDetails) {
        # Show selected function details from Functions blade
        $functions = $script:AllFunctions[$functionsBlade.Category]
        $functionIndex = if ($script:State.ActiveBlade -eq 1) { $script:State.Position } else { 0 }
        if ($functionIndex -lt $functions.Count) {
            $selectedFunction = $functions[$functionIndex]
            
            # Function name
            [Console]::SetCursorPosition($X + 1, $Y)
            Write-Host "Function: $($selectedFunction.Name)" -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
            
            # Description
            if ($selectedFunction.Description) {
                $currentY = $Y + 2
                [Console]::SetCursorPosition($X + 1, $currentY)
                Write-Host "Description:" -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
                $currentY++
                
                # Word wrap description
                $words = $selectedFunction.Description -split ' '
                $currentLine = ""
                foreach ($word in $words) {
                    if (($currentLine + " " + $word).Length -gt $Width - 4) {
                        [Console]::SetCursorPosition($X + 1, $currentY)
                        Write-Host $currentLine -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
                        $currentY++
                        $currentLine = $word
                    } else {
                        $currentLine = if ($currentLine) { "$currentLine $word" } else { $word }
                    }
                }
                if ($currentLine) {
                    [Console]::SetCursorPosition($X + 1, $currentY)
                    Write-Host $currentLine -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
                }
            }
        }
    } else {
        # Show category list when no Functions blade is open
        $startY = $Y
        for ($i = 0; $i -lt $script:Categories.Count; $i++) {
            if ($startY + $i -ge $Y + $Height - 1) { break }
            
            $category = $script:Categories[$i]
            $count = $script:AllFunctions[$category].Count
            $icon = Get-CategoryIcon $category
            
            [Console]::SetCursorPosition($X + 1, $startY + $i)
            
            if ($IsActive -and $i -eq $script:State.Position) {
                # Selected item - minimal highlight
                Write-Host "► $icon $category ($count)" -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
            } else {
                # Regular item
                Write-Host "  $icon $category ($count)" -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
            }
        }
    }
}

# Draw functions blade
function Draw-FunctionsBlade {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height, [hashtable]$Blade, [bool]$IsActive)
    
    $category = $Blade.Category
    $functions = $script:AllFunctions[$category]
    
    # Function list
    $startY = $Y
    for ($i = 0; $i -lt $functions.Count; $i++) {
        if ($startY + $i -ge $Y + $Height - 1) { break }
        
        $func = $functions[$i]
        
        [Console]::SetCursorPosition($X + 1, $startY + $i)
        
        if ($IsActive -and $i -eq $script:State.Position) {
            # Selected item - minimal highlight
            Write-Host "► $($func.Name)" -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
        } else {
            # Regular item
            Write-Host "  $($func.Name)" -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
        }
    }
}

# Draw details blade
function Draw-DetailsBlade {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height, [hashtable]$Blade, [bool]$IsActive)
    
    $func = $Blade.Function
    $currentY = $Y
    
    # Description
    if ($func.Description) {
        # Word wrap description
        $words = $func.Description -split ' '
        $currentLine = ""
        foreach ($word in $words) {
            if (($currentLine + " " + $word).Length -gt $Width - 4) {
                [Console]::SetCursorPosition($X + 1, $currentY)
                Write-Host $currentLine -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
                $currentY++
                $currentLine = $word
            } else {
                $currentLine = if ($currentLine) { "$currentLine $word" } else { $word }
            }
        }
        if ($currentLine) {
            [Console]::SetCursorPosition($X + 1, $currentY)
            Write-Host $currentLine -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
            $currentY++
        }
        $currentY += 2
    }
    
    # Actions
    $actions = @("Execute", "Help", "Back")
    for ($i = 0; $i -lt $actions.Count; $i++) {
        if ($currentY + $i -ge $Y + $Height - 1) { break }
        
        [Console]::SetCursorPosition($X + 1, $currentY + $i)
        
        if ($IsActive -and $i -eq $script:State.SubPosition) {
            # Selected action - minimal highlight
            Write-Host "► $($actions[$i])" -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
        } else {
            # Regular action
            Write-Host "  $($actions[$i])" -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
        }
    }
}

# Draw minimal status bar
function Draw-StatusBar {
    $y = $script:Height - 1
    [Console]::SetCursorPosition(0, $y)
    
    # Status bar with input instructions
    $helpText = "↑↓ Navigate │ ← → Blades │ Enter Select/Next │ Q Quit"
    $centerX = ($script:Width - $helpText.Length) / 2
    
    # Draw separator line
    Write-Host ("─" * [Math]::Floor($centerX)) -ForegroundColor Gray -BackgroundColor $script:Colors.Background -NoNewline
    
    # Draw help text
    Write-Host $helpText -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
    
    # Fill remaining space
    $remaining = $script:Width - [Math]::Floor($centerX) - $helpText.Length
    if ($remaining -gt 0) {
        Write-Host ("─" * $remaining) -ForegroundColor Gray -BackgroundColor $script:Colors.Background -NoNewline
    }
}

# Get category icon
function Get-CategoryIcon {
    param([string]$Category)
    
    switch ($Category) {
        "Log Management" { return "📊" }
        "Drive Management" { return "💾" }
        "Network" { return "🌐" }
        "Utilities" { return "🔧" }
        "System Information" { return "💻" }
        "Package Management" { return "📦" }
        "Device Management" { return "🖥️" }
        "Print Management" { return "🖨️" }
        "File Association Management" { return "📁" }
        default { return "⚡" }
    }
}

# Input handling
function Read-Input {
    if (-not [Console]::KeyAvailable) { return }
    
    $key = [Console]::ReadKey($true)
    
    switch ($key.Key) {
        "UpArrow" { Handle-NavigateUp }
        "DownArrow" { Handle-NavigateDown }
        "RightArrow" { Handle-BladeRight }
        "LeftArrow" { Handle-BladeLeft }
        "Enter" { Handle-EnterKey }
        "Q" { $script:State.Running = $false }
    }
}

# Navigation handlers
function Handle-NavigateUp {
    if ($script:State.OpenBlades.Count -eq 0) { return }
    
    $activeBlade = $script:State.OpenBlades[$script:State.ActiveBlade]
    
    switch ($activeBlade.Type) {
        "Categories" {
            if ($script:State.Position -gt 0) {
                $script:State.Position--
            } else {
                $script:State.Position = $script:Categories.Count - 1
            }
        }
        "Functions" {
            $functions = $script:AllFunctions[$activeBlade.Category]
            if ($script:State.Position -gt 0) {
                $script:State.Position--
            } else {
                $script:State.Position = $functions.Count - 1
            }
        }
        "Details" {
            if ($script:State.SubPosition -gt 0) {
                $script:State.SubPosition--
            } else {
                $script:State.SubPosition = 2
            }
        }
    }
}

function Handle-NavigateDown {
    if ($script:State.OpenBlades.Count -eq 0) { return }
    
    $activeBlade = $script:State.OpenBlades[$script:State.ActiveBlade]
    
    switch ($activeBlade.Type) {
        "Categories" {
            if ($script:State.Position -lt $script:Categories.Count - 1) {
                $script:State.Position++
            } else {
                $script:State.Position = 0
            }
        }
        "Functions" {
            $functions = $script:AllFunctions[$activeBlade.Category]
            if ($script:State.Position -lt $functions.Count - 1) {
                $script:State.Position++
            } else {
                $script:State.Position = 0
            }
        }
        "Details" {
            if ($script:State.SubPosition -lt 2) {
                $script:State.SubPosition++
            } else {
                $script:State.SubPosition = 0
            }
        }
    }
}

# Handle right arrow - open new blade or navigate to next blade
function Handle-BladeRight {
    if ($script:State.OpenBlades.Count -eq 0) { return }
    
    # If we can move to an existing blade to the right
    if ($script:State.ActiveBlade -lt $script:State.OpenBlades.Count - 1) {
        # Store current position before moving
        $currentBlade = $script:State.OpenBlades[$script:State.ActiveBlade]
        $script:State.BladePositions[$currentBlade.Type] = @{
            Position = $script:State.Position
            SubPosition = $script:State.SubPosition
        }
        
        $script:State.ActiveBlade++
        
        # Restore position for the blade we're moving to
        $activeBlade = $script:State.OpenBlades[$script:State.ActiveBlade]
        if ($script:State.BladePositions.ContainsKey($activeBlade.Type)) {
            $script:State.Position = $script:State.BladePositions[$activeBlade.Type].Position
            $script:State.SubPosition = $script:State.BladePositions[$activeBlade.Type].SubPosition
        } else {
            $script:State.Position = 0
            $script:State.SubPosition = 0
        }
        return
    }
    
    # Otherwise, try to open a new blade
    $activeBlade = $script:State.OpenBlades[$script:State.ActiveBlade]
    
    switch ($activeBlade.Type) {
        "Categories" {
            # Store current Categories position before opening Functions blade
            $script:State.BladePositions["Categories"] = @{
                Position = $script:State.Position
                SubPosition = $script:State.SubPosition
            }
            
            # Close any existing blades to the right before opening new one
            if ($script:State.OpenBlades.Count -gt 1) {
                $script:State.OpenBlades = $script:State.OpenBlades[0..0]
            }
            
            # Open functions blade
            $category = $script:Categories[$script:State.Position]
            $newBlade = @{
                Type = "Functions"
                Title = $category
                Width = $script:BladeTypes.Functions.Width
                Icon = $script:BladeTypes.Functions.Icon
                Category = $category
            }
            $script:State.OpenBlades += $newBlade
            $script:State.ActiveBlade = $script:State.OpenBlades.Count - 1
            
            # Restore position if we've been to this Functions blade before
            if ($script:State.BladePositions.ContainsKey("Functions")) {
                $script:State.Position = $script:State.BladePositions["Functions"].Position
                $script:State.SubPosition = $script:State.BladePositions["Functions"].SubPosition
            } else {
                $script:State.Position = 0
                $script:State.SubPosition = 0
            }
        }
        "Functions" {
            # Store current Functions position before opening Details blade
            $script:State.BladePositions["Functions"] = @{
                Position = $script:State.Position
                SubPosition = $script:State.SubPosition
            }
            
            # Close any existing details blades before opening new one
            if ($script:State.OpenBlades.Count -gt 2) {
                $script:State.OpenBlades = $script:State.OpenBlades[0..1]
            }
            
            # Open details blade
            $functions = $script:AllFunctions[$activeBlade.Category]
            $selectedFunction = $functions[$script:State.Position]
            $newBlade = @{
                Type = "Details"
                Title = $selectedFunction.Name
                Width = $script:BladeTypes.Details.Width
                Icon = $script:BladeTypes.Details.Icon
                Function = $selectedFunction
            }
            $script:State.OpenBlades += $newBlade
            $script:State.ActiveBlade = $script:State.OpenBlades.Count - 1
            
            # Restore position if we've been to this Details blade before
            if ($script:State.BladePositions.ContainsKey("Details")) {
                $script:State.Position = $script:State.BladePositions["Details"].Position
                $script:State.SubPosition = $script:State.BladePositions["Details"].SubPosition
            } else {
                $script:State.Position = 0
                $script:State.SubPosition = 0
            }
        }
    }
}

# Handle left arrow - always close rightmost blade or move focus left
function Handle-BladeLeft {
    if ($script:State.OpenBlades.Count -eq 0) { return }
    
    # If there are multiple blades, always close the rightmost one
    if ($script:State.OpenBlades.Count -gt 1) {
        # Store current position before closing blade
        $closingBlade = $script:State.OpenBlades[-1]
        $script:State.BladePositions[$closingBlade.Type] = @{
            Position = $script:State.Position
            SubPosition = $script:State.SubPosition
        }
        
        $script:State.OpenBlades = $script:State.OpenBlades[0..($script:State.OpenBlades.Count - 2)]
        
        # If we were on the blade that got closed, move to the new rightmost blade
        if ($script:State.ActiveBlade -ge $script:State.OpenBlades.Count) {
            $script:State.ActiveBlade = $script:State.OpenBlades.Count - 1
        }
        
        # Restore position for the blade we're returning to
        $activeBlade = $script:State.OpenBlades[$script:State.ActiveBlade]
        if ($script:State.BladePositions.ContainsKey($activeBlade.Type)) {
            $script:State.Position = $script:State.BladePositions[$activeBlade.Type].Position
            $script:State.SubPosition = $script:State.BladePositions[$activeBlade.Type].SubPosition
        } else {
            $script:State.Position = 0
            $script:State.SubPosition = 0
        }
    }
}

function Handle-EnterKey {
    if ($script:State.OpenBlades.Count -eq 0) { return }
    
    $activeBlade = $script:State.OpenBlades[$script:State.ActiveBlade]
    
    # For Details blade, handle actions
    if ($activeBlade.Type -eq "Details") {
        switch ($script:State.SubPosition) {
            0 { Execute-Function }
            1 { Show-Help }
            2 { Handle-BladeLeft }  # Back action
        }
    }
    # For other blades, progress to next blade like right arrow
    else {
        Handle-BladeRight
    }
}

function Handle-Select {
    Handle-EnterKey
}

# Function execution
function Execute-Function {
    if ($script:State.OpenBlades.Count -eq 0) { return }
    
    $detailsBlade = $script:State.OpenBlades | Where-Object { $_.Type -eq "Details" } | Select-Object -First 1
    if (-not $detailsBlade) { return }
    
    $func = $detailsBlade.Function
    
    Clear-Host
    Write-Host "`n  Executing: $($func.Name)" -ForegroundColor Cyan
    Write-Host ("  " + "═" * 60) -ForegroundColor DarkCyan
    
    try {
        & $func.Name
        Write-Host "`n  " -NoNewline
        Write-Host ("═" * 60) -ForegroundColor DarkCyan
        Write-Host "  Execution completed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "`n  " -NoNewline
        Write-Host ("═" * 60) -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
    }
    
    Write-Host "`n  Press any key to continue..." -ForegroundColor Gray
    $null = [Console]::ReadKey($true)
    
    # Return to blade interface
    $script:State.LastUpdate.State = $null  # Force redraw
    Clear-Host
    Render-Screen
}

# Help display
function Show-Help {
    if ($script:State.OpenBlades.Count -eq 0) { return }
    
    $detailsBlade = $script:State.OpenBlades | Where-Object { $_.Type -eq "Details" } | Select-Object -First 1
    if (-not $detailsBlade) { return }
    
    $func = $detailsBlade.Function
    
    Clear-Host
    Get-Help $func.Name -Full
    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
    $null = [Console]::ReadKey($true)
    
    # Return to blade interface
    $script:State.LastUpdate.State = $null  # Force redraw
    Clear-Host
    Render-Screen
}

# Main loop
function Start-BladeSystem {
    Clear-Host
    
    # Initialize
    Write-Host "Loading VTS Minimal..." -ForegroundColor Cyan
    Initialize-Functions
    
    # Open initial categories blade
    $initialBlade = @{
        Type = "Categories"
        Title = $script:BladeTypes.Categories.Title
        Width = $script:BladeTypes.Categories.Width
        Icon = $script:BladeTypes.Categories.Icon
    }
    $script:State.OpenBlades = @($initialBlade)
    $script:State.ActiveBlade = 0
    $script:State.Position = 0
    
    # Main loop
    while ($script:State.Running) {
        Render-Screen
        Read-Input
        Start-Sleep -Milliseconds 16  # 60 FPS
    }
    
    # Cleanup
    Clear-Host
    [Console]::CursorVisible = $true
    Write-Host "Thank you for using VTS Minimal!" -ForegroundColor Cyan
}

# Error handling
try {
    Start-BladeSystem
}
catch {
    [Console]::CursorVisible = $true
    Write-Host "Error: $_" -ForegroundColor Red
}
finally {
    [Console]::CursorVisible = $true
}