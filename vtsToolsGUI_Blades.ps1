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
$script:ExampleRegex = [regex]::new('\.EXAMPLE\s*\n\s*(.+?)(?=\n\s*\.|\n\s*#>|$)', [System.Text.RegularExpressions.RegexOptions]::Compiled)

# Load vtsTools efficiently
if (-not (Test-Path $ToolsPath)) {
    Write-Host "ERROR: Missing vtsTools.ps1 at $ToolsPath" -ForegroundColor Red
    exit 1
}
. $ToolsPath

# Screen dimensions with safety checks
try {
    $script:Width = [Math]::Max(80, $Host.UI.RawUI.WindowSize.Width)
    $script:Height = [Math]::Max(24, $Host.UI.RawUI.WindowSize.Height)
} catch {
    $script:Width = 120
    $script:Height = 30
}

# Enhanced modern color scheme
$script:Colors = @{
    Background = [ConsoleColor]::Black
    Text = [ConsoleColor]::White
    Highlight = [ConsoleColor]::Cyan
    Selected = [ConsoleColor]::Black
    Border = [ConsoleColor]::DarkCyan
    Accent = [ConsoleColor]::Green
    Secondary = [ConsoleColor]::Gray
    Warning = [ConsoleColor]::Yellow
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
    # Search functionality
    SearchMode = $true
    SearchQuery = ""
    FilteredFunctions = @()
    ParameterInput = @{}
    InputMode = $false
    InputField = ""
    InputValue = ""
}

# Blade definitions with modern icons
$script:BladeTypes = @{
    Search = @{
        Title = "Search Results"
        Width = 45
        Icon = "🔍"
    }
    Categories = @{
        Title = "Categories"
        Width = 32
        Icon = "▶"
    }
    Functions = @{
        Title = "Functions"
        Width = 42
        Icon = "◆"
    }
    Details = @{
        Title = "Function Details"
        Width = 55
        Icon = "●"
    }
}

# Pre-load all functions
$script:AllFunctions = @{}
$script:Categories = @()
$script:AllFunctionsFlat = @()

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
        $example = ""
        
        # Fast regex extraction
        $descMatch = $script:DescriptionRegex.Match($text)
        if ($descMatch.Success) {
            $description = $descMatch.Groups[1].Value.Trim()
        }
        
        $linkMatch = $script:LinkRegex.Match($text)
        if ($linkMatch.Success) {
            $category = $linkMatch.Groups[1].Value.Trim()
        }
        
        $exampleMatch = $script:ExampleRegex.Match($text)
        if ($exampleMatch.Success) {
            $example = $exampleMatch.Groups[1].Value.Trim()
        }
        
        if (-not $script:AllFunctions.ContainsKey($category)) {
            $script:AllFunctions[$category] = @()
        }
        
        $funcObj = @{
            Name = $name
            Description = $description
            Category = $category
            Example = $example
        }
        $script:AllFunctions[$category] += $funcObj
        $script:AllFunctionsFlat += $funcObj
    }
    
    $script:Categories = $script:AllFunctions.Keys | Sort-Object
    
    # Initialize search with all functions
    Update-FilteredFunctions
}

# Search and filter functions with performance optimization
function Update-FilteredFunctions {
    if ([string]::IsNullOrWhiteSpace($script:State.SearchQuery)) {
        $script:State.FilteredFunctions = $script:AllFunctionsFlat
    } else {
        $query = $script:State.SearchQuery.ToString().ToLower()
        $script:State.FilteredFunctions = @()
        # Use simple loop instead of Where-Object for better performance
        foreach ($func in $script:AllFunctionsFlat) {
            $name = if ($func.Name) { $func.Name.ToString().ToLower() } else { "" }
            $description = if ($func.Description) { $func.Description.ToString().ToLower() } else { "" }
            $category = if ($func.Category) { $func.Category.ToString().ToLower() } else { "" }
            
            if ($name.Contains($query) -or
                $description.Contains($query) -or
                $category.Contains($query)) {
                $script:State.FilteredFunctions += $func
            }
        }
    }
}

# Main rendering function
function Render-Screen {
    # Check if we need to update
    $currentState = "$($script:State.OpenBlades.Count)_$($script:State.ActiveBlade)_$($script:State.Position)_$($script:State.SubPosition)_$($script:State.SearchQuery)_$($script:State.InputMode)"
    if ($script:State.LastUpdate.State -eq $currentState) {
        return
    }
    
    # Clear screen
    Clear-Host
    
    # Draw background
    Draw-Background
    
    # Draw search bar at top
    Draw-SearchBar
    
    # Draw all open blades with margins (offset by search bar height)
    $x = 2  # Start with left margin
    for ($i = 0; $i -lt $script:State.OpenBlades.Count; $i++) {
        $blade = $script:State.OpenBlades[$i]
        $isActive = ($i -eq $script:State.ActiveBlade)
        
        Draw-Blade -X $x -Blade $blade -IsActive $isActive -YOffset 3
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

# Draw search bar at top
function Draw-SearchBar {
    # Search bar border
    [Console]::SetCursorPosition(0, 0)
    Write-Host ("─" * $script:Width) -ForegroundColor $script:Colors.Border -BackgroundColor $script:Colors.Background -NoNewline
    
    # Search input line
    [Console]::SetCursorPosition(0, 1)
    $searchLabel = "🔍 Search: "
    Write-Host $searchLabel -ForegroundColor $script:Colors.Accent -BackgroundColor $script:Colors.Background -NoNewline
    
    # Input field
    $inputWidth = $script:Width - $searchLabel.Length - 2
    $displayQuery = $script:State.SearchQuery
    if ($script:State.InputMode) {
        $displayQuery = $script:State.InputValue
    }
    
    # Truncate if too long
    if ($displayQuery.Length -gt $inputWidth) {
        $displayQuery = $displayQuery.Substring(0, $inputWidth)
    }
    
    # Show cursor if in input mode
    if ($script:State.InputMode) {
        Write-Host "$displayQuery█" -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
        $remaining = $inputWidth - $displayQuery.Length - 1
    } else {
        Write-Host $displayQuery -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
        $remaining = $inputWidth - $displayQuery.Length
    }
    
    # Fill remaining space
    if ($remaining -gt 0) {
        Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
    }
    
    # Results count
    [Console]::SetCursorPosition(0, 2)
    $resultsText = "Found: $($script:State.FilteredFunctions.Count) functions"
    Write-Host $resultsText -ForegroundColor $script:Colors.Secondary -BackgroundColor $script:Colors.Background -NoNewline
    $remaining = $script:Width - $resultsText.Length
    if ($remaining -gt 0) {
        Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
    }
}

# Draw blade with proper boundaries
function Draw-Blade {
    param(
        [int]$X,
        [hashtable]$Blade,
        [bool]$IsActive,
        [int]$YOffset = 0
    )
    
    $width = $Blade.Width
    $height = $script:Height - 1 - $YOffset
    
    # Clean title line with proper truncation
    [Console]::SetCursorPosition($X, $YOffset)
    $titleText = "$($Blade.Icon) $($Blade.Title)"
    if ($titleText.Length -gt $width -and $width -gt 3) {
        $titleText = $titleText.Substring(0, $width - 3) + "..."
    }
    
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
    [Console]::SetCursorPosition($X, $YOffset + 1)
    Write-Host (" " * $width) -BackgroundColor $script:Colors.Background -NoNewline
    
    # Blade content with proper boundaries
    Draw-BladeContent -X $X -Y ($YOffset + 2) -Width $width -Height ($height - 2) -Blade $Blade -IsActive $IsActive
    
    # Fill any remaining vertical space to prevent artifacts
    for ($y = $YOffset + 2; $y -lt $YOffset + $height; $y++) {
        $lineContent = $false
        # Check if this line has content (rough estimation)
        if ($Blade.Type -eq "Search" -and $y - $YOffset - 2 -lt $script:State.FilteredFunctions.Count) {
            $lineContent = $true
        } elseif ($Blade.Type -eq "Categories" -and $y - $YOffset - 2 -lt $script:Categories.Count) {
            $lineContent = $true
        } elseif ($Blade.Type -eq "Functions" -and $Blade.Category -and $script:AllFunctions.ContainsKey($Blade.Category) -and $y - $YOffset - 2 -lt $script:AllFunctions[$Blade.Category].Count) {
            $lineContent = $true
        } elseif ($Blade.Type -eq "Details") {
            $lineContent = $true  # Details blade manages its own content
        }
        
        if (-not $lineContent) {
            [Console]::SetCursorPosition($X, $y)
            Write-Host (" " * $width) -BackgroundColor $script:Colors.Background -NoNewline
        }
    }
}

# Draw blade content based on type
function Draw-BladeContent {
    param(
        [int]$X, [int]$Y, [int]$Width, [int]$Height,
        [hashtable]$Blade, [bool]$IsActive
    )
    
    switch ($Blade.Type) {
        "Search" {
            Draw-SearchBlade -X $X -Y $Y -Width $Width -Height $Height -IsActive $IsActive
        }
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

# Draw search results blade
function Draw-SearchBlade {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height, [bool]$IsActive)
    
    # Show filtered functions
    $startY = $Y
    for ($i = 0; $i -lt $script:State.FilteredFunctions.Count; $i++) {
        if ($startY + $i -ge $Y + $Height - 1) { break }
        
        $func = $script:State.FilteredFunctions[$i]
        
        [Console]::SetCursorPosition($X + 1, $startY + $i)
        
        if ($IsActive -and $i -eq $script:State.Position) {
            # Selected item with modern styling
            $text = "▶ $($func.Name)"
            # Truncate if too long
            $maxLength = $Width - 2
            if ($text.Length -gt $maxLength -and $maxLength -gt 3) {
                $text = $text.Substring(0, $maxLength - 3) + "..."
            }
            Write-Host $text -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space
            $remaining = $Width - $text.Length - 1
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
        } else {
            # Regular item
            $text = "◆ $($func.Name)"
            # Truncate if too long
            $maxLength = $Width - 2
            if ($text.Length -gt $maxLength -and $maxLength -gt 3) {
                $text = $text.Substring(0, $maxLength - 3) + "..."
            }
            Write-Host $text -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space
            $remaining = $Width - $text.Length - 1
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
        }
    }
}

# Draw categories blade - only show categories list with proper spacing
function Draw-CategoriesBlade {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height, [bool]$IsActive)
    
    # Always show category list
    $startY = $Y
    for ($i = 0; $i -lt $script:Categories.Count; $i++) {
        if ($startY + $i -ge $Y + $Height - 1) { break }
        
        $category = $script:Categories[$i]
        $count = $script:AllFunctions[$category].Count
        $icon = Get-CategoryIcon $category
        
        [Console]::SetCursorPosition($X + 1, $startY + $i)
        
        if ($IsActive -and $i -eq $script:State.Position) {
            # Selected item with modern styling
            $text = "▶ $icon $category ($count)"
            # Truncate if too long
            $maxLength = $Width - 2
            if ($text.Length -gt $maxLength -and $maxLength -gt 3) {
                $text = $text.Substring(0, $maxLength - 3) + "..."
            }
            Write-Host $text -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space
            $remaining = $Width - $text.Length - 1
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
        } else {
            # Regular item
            $text = "  $icon $category ($count)"
            # Truncate if too long
            $maxLength = $Width - 2
            if ($text.Length -gt $maxLength -and $maxLength -gt 3) {
                $text = $text.Substring(0, $maxLength - 3) + "..."
            }
            Write-Host $text -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space
            $remaining = $Width - $text.Length - 1
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
        }
    }
}

# Draw functions blade with proper text boundaries
function Draw-FunctionsBlade {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height, [hashtable]$Blade, [bool]$IsActive)
    
    $category = $Blade.Category
    $functions = $script:AllFunctions[$category]
    
    # Function list with better styling and proper spacing
    $startY = $Y
    for ($i = 0; $i -lt $functions.Count; $i++) {
        if ($startY + $i -ge $Y + $Height - 1) { break }
        
        $func = $functions[$i]
        
        [Console]::SetCursorPosition($X + 1, $startY + $i)
        
        if ($IsActive -and $i -eq $script:State.Position) {
            # Selected item with modern styling
            $text = "▶ $($func.Name)"
            # Truncate if too long
            $maxLength = $Width - 2
            if ($text.Length -gt $maxLength -and $maxLength -gt 3) {
                $text = $text.Substring(0, $maxLength - 3) + "..."
            }
            Write-Host $text -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space
            $remaining = $Width - $text.Length - 1
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
        } else {
            # Regular item with subtle styling
            $text = "◆ $($func.Name)"
            # Truncate if too long
            $maxLength = $Width - 2
            if ($text.Length -gt $maxLength -and $maxLength -gt 3) {
                $text = $text.Substring(0, $maxLength - 3) + "..."
            }
            Write-Host $text -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space
            $remaining = $Width - $text.Length - 1
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
        }
    }
}

# Draw details blade with proper word wrapping
function Draw-DetailsBlade {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height, [hashtable]$Blade, [bool]$IsActive)
    
    $func = $Blade.Function
    $currentY = $Y
    $maxY = $Y + $Height - 1
    $contentWidth = $Width - 4  # Leave padding for margins
    
    # Description section
    if ($func.Description -and $currentY -lt $maxY) {
        [Console]::SetCursorPosition($X + 1, $currentY)
        Write-Host "Description:" -ForegroundColor $script:Colors.Accent -BackgroundColor $script:Colors.Background -NoNewline
        # Fill remaining space on title line
        $remaining = $Width - 13  # "Description:" is 12 chars + 1 for position
        if ($remaining -gt 0) {
            Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
        }
        $currentY++
        
        # Word wrap description with proper boundaries
        $words = $func.Description -split ' '
        $currentLine = ""
        foreach ($word in $words) {
            if ($currentY -ge $maxY) { break }
            
            $testLine = if ($currentLine) { "$currentLine $word" } else { $word }
            if ($testLine.Length -gt $contentWidth) {
                # Output current line if it has content
                if ($currentLine) {
                    [Console]::SetCursorPosition($X + 2, $currentY)
                    Write-Host $currentLine -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
                    # Fill remaining space on line
                    $remaining = $Width - $currentLine.Length - 2
                    if ($remaining -gt 0) {
                        Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
                    }
                    $currentY++
                }
                $currentLine = $word
            } else {
                $currentLine = $testLine
            }
        }
        # Output final line if it has content
        if ($currentLine -and $currentY -lt $maxY) {
            [Console]::SetCursorPosition($X + 2, $currentY)
            Write-Host $currentLine -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space on line
            $remaining = $Width - $currentLine.Length - 2
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
            $currentY++
        }
        $currentY++  # Add space after description
    }
    
    # Example section
    if ($func.Example -and $currentY -lt $maxY) {
        [Console]::SetCursorPosition($X + 1, $currentY)
        Write-Host "Example:" -ForegroundColor $script:Colors.Accent -BackgroundColor $script:Colors.Background -NoNewline
        # Fill remaining space on title line
        $remaining = $Width - 9  # "Example:" is 8 chars + 1 for position
        if ($remaining -gt 0) {
            Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
        }
        $currentY++
        
        # Word wrap example with proper boundaries
        $words = $func.Example -split ' '
        $currentLine = ""
        foreach ($word in $words) {
            if ($currentY -ge $maxY) { break }
            
            $testLine = if ($currentLine) { "$currentLine $word" } else { $word }
            if ($testLine.Length -gt $contentWidth) {
                # Output current line if it has content
                if ($currentLine) {
                    [Console]::SetCursorPosition($X + 2, $currentY)
                    Write-Host $currentLine -ForegroundColor $script:Colors.Secondary -BackgroundColor $script:Colors.Background -NoNewline
                    # Fill remaining space on line
                    $remaining = $Width - $currentLine.Length - 2
                    if ($remaining -gt 0) {
                        Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
                    }
                    $currentY++
                }
                $currentLine = $word
            } else {
                $currentLine = $testLine
            }
        }
        # Output final line if it has content
        if ($currentLine -and $currentY -lt $maxY) {
            [Console]::SetCursorPosition($X + 2, $currentY)
            Write-Host $currentLine -ForegroundColor $script:Colors.Secondary -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space on line
            $remaining = $Width - $currentLine.Length - 2
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
            $currentY++
        }
        $currentY++  # Add space after example
    }
    
    # Actions with modern styling and proper spacing
    $actions = @("Execute", "Parameters", "Help", "Back")
    $actionIcons = @("▶", "⚙", "●", "◄")
    for ($i = 0; $i -lt $actions.Count; $i++) {
        if ($currentY + $i -ge $maxY) { break }
        
        [Console]::SetCursorPosition($X + 1, $currentY + $i)
        
        if ($IsActive -and $i -eq $script:State.SubPosition) {
            # Selected action with modern styling
            Write-Host "$($actionIcons[$i]) " -ForegroundColor $script:Colors.Accent -BackgroundColor $script:Colors.Background -NoNewline
            Write-Host $actions[$i] -ForegroundColor $script:Colors.Highlight -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space on line
            $remaining = $Width - $actions[$i].Length - 3
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
        } else {
            # Regular action
            Write-Host "$($actionIcons[$i]) " -ForegroundColor $script:Colors.Secondary -BackgroundColor $script:Colors.Background -NoNewline
            Write-Host $actions[$i] -ForegroundColor $script:Colors.Text -BackgroundColor $script:Colors.Background -NoNewline
            # Fill remaining space on line
            $remaining = $Width - $actions[$i].Length - 3
            if ($remaining -gt 0) {
                Write-Host (" " * $remaining) -BackgroundColor $script:Colors.Background -NoNewline
            }
        }
    }
}

# Draw modern status bar
function Draw-StatusBar {
    $y = $script:Height - 1
    [Console]::SetCursorPosition(0, $y)
    
    # Status bar with modern styling
    if ($script:State.InputMode) {
        $helpText = "Type to search ● Esc Cancel ● Enter Confirm"
    } else {
        $helpText = "▲▼ Navigate ● ◄► Blades ● Enter Select ● / Search ● Q Quit"
    }
    $centerX = [Math]::Max(0, ($script:Width - $helpText.Length) / 2)
    
    # Draw separator line with modern style
    Write-Host ("─" * [Math]::Floor($centerX)) -ForegroundColor $script:Colors.Secondary -BackgroundColor $script:Colors.Background -NoNewline
    
    # Draw help text with accent colors
    Write-Host $helpText -ForegroundColor $script:Colors.Accent -BackgroundColor $script:Colors.Background -NoNewline
    
    # Fill remaining space
    $remaining = $script:Width - [Math]::Floor($centerX) - $helpText.Length
    if ($remaining -gt 0) {
        Write-Host ("─" * $remaining) -ForegroundColor $script:Colors.Secondary -BackgroundColor $script:Colors.Background -NoNewline
    }
}

# Get category icon - modern geometric style
function Get-CategoryIcon {
    param([string]$Category)
    
    switch ($Category) {
        "Log Management" { return "■" }
        "Drive Management" { return "◆" }
        "Network" { return "○" }
        "Utilities" { return "▲" }
        "System Information" { return "□" }
        "Package Management" { return "◇" }
        "Device Management" { return "◐" }
        "Print Management" { return "◑" }
        "File Association Management" { return "△" }
        default { return "●" }
    }
}

# Update Details blade when function selection changes
function Update-DetailsBlade {
    # Find the Functions blade
    $functionsBlade = $null
    $functionsBladeIndex = -1
    for ($i = 0; $i -lt $script:State.OpenBlades.Count; $i++) {
        if ($script:State.OpenBlades[$i].Type -eq "Functions") {
            $functionsBlade = $script:State.OpenBlades[$i]
            $functionsBladeIndex = $i
            break
        }
    }
    
    if ($functionsBlade -and $script:State.OpenBlades.Count -gt 2) {
        # Get the selected function
        $functions = $script:AllFunctions[$functionsBlade.Category]
        if ($functions -and $script:State.Position -ge 0 -and $script:State.Position -lt $functions.Count) {
            $selectedFunction = $functions[$script:State.Position]
            
            # Update the Details blade
            $script:State.OpenBlades[2] = @{
                Type = "Details"
                Title = $selectedFunction.Name
                Width = $script:BladeTypes.Details.Width
                Icon = $script:BladeTypes.Details.Icon
                Function = $selectedFunction
            }
            
            # Force redraw
            $script:State.LastUpdate.State = $null
        }
    }
}

# Input handling
function Read-Input {
    if (-not [Console]::KeyAvailable) { return }
    
    $key = [Console]::ReadKey($true)
    
    # Handle input mode (search typing)
    if ($script:State.InputMode) {
        switch ($key.Key) {
            "Escape" {
                $script:State.InputMode = $false
                $script:State.InputValue = ""
                $script:State.SearchQuery = ""
                Update-FilteredFunctions
                $script:State.Position = 0
            }
            "Enter" {
                $script:State.SearchQuery = $script:State.InputValue
                $script:State.InputMode = $false
                $script:State.InputValue = ""
                Update-FilteredFunctions
                $script:State.Position = 0
            }
            "Backspace" {
                if ($script:State.InputValue.Length -gt 0) {
                    $script:State.InputValue = $script:State.InputValue.Substring(0, $script:State.InputValue.Length - 1)
                    # Update search in real-time but with debouncing
                    $script:State.SearchQuery = $script:State.InputValue.ToString()
                    Update-FilteredFunctions
                    $script:State.Position = 0
                }
            }
            default {
                # Add character to input if printable
                if ($key.KeyChar -match '[\x20-\x7E]') {
                    $script:State.InputValue += $key.KeyChar
                    # Update search in real-time
                    $script:State.SearchQuery = $script:State.InputValue.ToString()
                    Update-FilteredFunctions
                    $script:State.Position = 0
                }
            }
        }
        return
    }
    
    # Normal navigation mode
    switch ($key.Key) {
        "UpArrow" { Handle-NavigateUp }
        "DownArrow" { Handle-NavigateDown }
        "RightArrow" { Handle-BladeRight }
        "LeftArrow" { Handle-BladeLeft }
        "Enter" { Handle-EnterKey }
        "Q" { $script:State.Running = $false }
        "Oem2" { # Forward slash key
            if ($key.Modifiers -eq [ConsoleModifiers]::None) {
                # Switch to search blade and enter input mode
                $searchBladeIndex = -1
                for ($i = 0; $i -lt $script:State.OpenBlades.Count; $i++) {
                    if ($script:State.OpenBlades[$i].Type -eq "Search") {
                        $searchBladeIndex = $i
                        break
                    }
                }
                if ($searchBladeIndex -ge 0) {
                    $script:State.ActiveBlade = $searchBladeIndex
                    $script:State.InputMode = $true
                    $script:State.InputValue = ""
                }
            }
        }
        default {
            # Check for direct typing to start search
            if ($key.KeyChar -match '[\x20-\x7E]') {
                # Switch to search blade and enter input mode
                $searchBladeIndex = -1
                for ($i = 0; $i -lt $script:State.OpenBlades.Count; $i++) {
                    if ($script:State.OpenBlades[$i].Type -eq "Search") {
                        $searchBladeIndex = $i
                        break
                    }
                }
                if ($searchBladeIndex -ge 0) {
                    $script:State.ActiveBlade = $searchBladeIndex
                    $script:State.InputMode = $true
                    $script:State.InputValue = $key.KeyChar.ToString()
                    $script:State.SearchQuery = $key.KeyChar.ToString()
                    Update-FilteredFunctions
                    $script:State.Position = 0
                }
            }
        }
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
            if ($functions -and $functions.Count -gt 0) {
                if ($script:State.Position -gt 0) {
                    $script:State.Position--
                } else {
                    $script:State.Position = $functions.Count - 1
                }
            }
            # Auto-update Details blade
            Update-DetailsBlade
        }
        "Details" {
            if ($script:State.SubPosition -gt 0) {
                $script:State.SubPosition--
            } else {
                $script:State.SubPosition = 3
            }
        }
        "Search" {
            if ($script:State.Position -gt 0) {
                $script:State.Position--
            } else {
                $script:State.Position = [Math]::Max(0, $script:State.FilteredFunctions.Count - 1)
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
            if ($functions -and $functions.Count -gt 0) {
                if ($script:State.Position -lt $functions.Count - 1) {
                    $script:State.Position++
                } else {
                    $script:State.Position = 0
                }
            }
            # Auto-update Details blade
            Update-DetailsBlade
        }
        "Details" {
            if ($script:State.SubPosition -lt 3) {
                $script:State.SubPosition++
            } else {
                $script:State.SubPosition = 0
            }
        }
        "Search" {
            if ($script:State.Position -lt $script:State.FilteredFunctions.Count - 1) {
                $script:State.Position++
            } else {
                $script:State.Position = 0
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
        $bladeKey = "$($currentBlade.Type)_$($currentBlade.Category)"
        $script:State.BladePositions[$bladeKey] = @{
            Position = $script:State.Position
            SubPosition = $script:State.SubPosition
        }
        
        $script:State.ActiveBlade++
        
        # Restore position for the blade we're moving to
        $activeBlade = $script:State.OpenBlades[$script:State.ActiveBlade]
        $bladeKey = "$($activeBlade.Type)_$($activeBlade.Category)"
        if ($script:State.BladePositions.ContainsKey($bladeKey)) {
            $script:State.Position = $script:State.BladePositions[$bladeKey].Position
            $script:State.SubPosition = $script:State.BladePositions[$bladeKey].SubPosition
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
            $script:State.BladePositions["Categories_"] = @{
                Position = $script:State.Position
                SubPosition = $script:State.SubPosition
            }
            
            # Close any existing blades to the right before opening new one
            if ($script:State.OpenBlades.Count -gt 1) {
                $script:State.OpenBlades = @($script:State.OpenBlades[0])
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
            
            # Always start at position 0 for new category
            $script:State.Position = 0
            $script:State.SubPosition = 0
            
            # Auto-open Details blade with the first function
            $functions = $script:AllFunctions[$category]
            if ($functions.Count -gt 0) {
                $selectedFunction = $functions[$script:State.Position]
                $detailsBlade = @{
                    Type = "Details"
                    Title = $selectedFunction.Name
                    Width = $script:BladeTypes.Details.Width
                    Icon = $script:BladeTypes.Details.Icon
                    Function = $selectedFunction
                }
                $script:State.OpenBlades += $detailsBlade
            }
            
            # Set active blade to Functions
            $script:State.ActiveBlade = 1
        }
        "Functions" {
            # Move to Details blade if it exists
            if ($script:State.OpenBlades.Count -gt 2) {
                # Store current Functions position
                $script:State.BladePositions["Functions_$($activeBlade.Category)"] = @{
                    Position = $script:State.Position
                    SubPosition = $script:State.SubPosition
                }
                $script:State.ActiveBlade = 2
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
        $bladeKey = "$($closingBlade.Type)_$($closingBlade.Category)"
        $script:State.BladePositions[$bladeKey] = @{
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
        $bladeKey = "$($activeBlade.Type)_$($activeBlade.Category)"
        if ($script:State.BladePositions.ContainsKey($bladeKey)) {
            $script:State.Position = $script:State.BladePositions[$bladeKey].Position
            $script:State.SubPosition = $script:State.BladePositions[$bladeKey].SubPosition
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
            1 { Show-ParameterInput }
            2 { Show-Help }
            3 { Handle-BladeLeft }  # Back action
        }
    }
    # For Search blade, open details for selected function
    elseif ($activeBlade.Type -eq "Search") {
        if ($script:State.FilteredFunctions.Count -gt 0 -and $script:State.Position -ge 0 -and $script:State.Position -lt $script:State.FilteredFunctions.Count) {
            $selectedFunction = $script:State.FilteredFunctions[$script:State.Position]
            Open-FunctionDetails $selectedFunction
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

# Open function details from search
function Open-FunctionDetails {
    param([hashtable]$Function)
    
    # Close existing blades and open details
    $detailsBlade = @{
        Type = "Details"
        Title = $Function.Name
        Width = $script:BladeTypes.Details.Width
        Icon = $script:BladeTypes.Details.Icon
        Function = $Function
    }
    
    # Replace current blades with search and details
    $script:State.OpenBlades = @(
        $script:State.OpenBlades[0],  # Keep search blade
        $detailsBlade
    )
    $script:State.ActiveBlade = 1
    $script:State.Position = 0
    $script:State.SubPosition = 0
}

# Parameter input system
function Show-ParameterInput {
    $detailsBlade = $script:State.OpenBlades | Where-Object { $_.Type -eq "Details" } | Select-Object -First 1
    if (-not $detailsBlade) { return }
    
    $func = $detailsBlade.Function
    
    Clear-Host
    Write-Host "`n  ⚙ Parameters for: " -ForegroundColor $script:Colors.Accent -NoNewline
    Write-Host $func.Name -ForegroundColor $script:Colors.Highlight
    Write-Host ("  " + "─" * 60) -ForegroundColor $script:Colors.Secondary
    
    try {
        # Get function parameters
        $command = Get-Command $func.Name -ErrorAction SilentlyContinue
        if ($command) {
            $parameters = $command.Parameters
            if ($parameters.Count -gt 0) {
                Write-Host "`n  Available Parameters:" -ForegroundColor $script:Colors.Accent
                
                foreach ($param in $parameters.Keys) {
                    $paramInfo = $parameters[$param]
                    $type = $paramInfo.ParameterType.Name
                    $mandatory = $paramInfo.Attributes | Where-Object { $_.Mandatory } | Select-Object -First 1
                    $mandatoryText = if ($mandatory) { "[Required]" } else { "[Optional]" }
                    
                    Write-Host "    • " -ForegroundColor $script:Colors.Secondary -NoNewline
                    Write-Host "$param" -ForegroundColor $script:Colors.Highlight -NoNewline
                    Write-Host " ($type) $mandatoryText" -ForegroundColor $script:Colors.Text
                }
                
                Write-Host "`n  Example usage:" -ForegroundColor $script:Colors.Accent
                Write-Host "    $($func.Name) -Parameter Value" -ForegroundColor $script:Colors.Secondary
            } else {
                Write-Host "`n  ● This function has no parameters" -ForegroundColor $script:Colors.Accent
            }
        } else {
            Write-Host "`n  ● Function not found or not accessible" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "`n  ● Error retrieving parameters: $_" -ForegroundColor Red
    }
    
    Write-Host "`n  Press any key to continue..." -ForegroundColor $script:Colors.Secondary
    $null = [Console]::ReadKey($true)
    
    # Return to blade interface
    $script:State.LastUpdate.State = $null  # Force redraw
    Clear-Host
    Render-Screen
}

# Function execution
function Execute-Function {
    if ($script:State.OpenBlades.Count -eq 0) { return }
    
    $detailsBlade = $script:State.OpenBlades | Where-Object { $_.Type -eq "Details" } | Select-Object -First 1
    if (-not $detailsBlade) { return }
    
    $func = $detailsBlade.Function
    
    Clear-Host
    Write-Host "`n  ▶ Executing: " -ForegroundColor $script:Colors.Accent -NoNewline
    Write-Host $func.Name -ForegroundColor $script:Colors.Highlight
    Write-Host ("  " + "─" * 60) -ForegroundColor $script:Colors.Secondary
    
    try {
        if (Get-Command $func.Name -ErrorAction SilentlyContinue) {
            & $func.Name
            Write-Host "`n  " -NoNewline
            Write-Host ("─" * 60) -ForegroundColor $script:Colors.Accent
            Write-Host "  ● Execution completed successfully!" -ForegroundColor $script:Colors.Accent
        } else {
            Write-Host "`n  " -NoNewline
            Write-Host ("─" * 60) -ForegroundColor Red
            Write-Host "  ● Error: Function '$($func.Name)' not found" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "`n  " -NoNewline
        Write-Host ("─" * 60) -ForegroundColor Red
        Write-Host "  ● Error: $_" -ForegroundColor Red
    }
    
    Write-Host "`n  Press any key to continue..." -ForegroundColor $script:Colors.Secondary
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
    try {
        if (Get-Command $func.Name -ErrorAction SilentlyContinue) {
            Get-Help $func.Name -Full
        } else {
            Write-Host "Help not available for '$($func.Name)'" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error retrieving help: $_" -ForegroundColor Red
    }
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
    
    # Open initial blades - Categories and Search
    $categoriesBlade = @{
        Type = "Categories"
        Title = $script:BladeTypes.Categories.Title
        Width = $script:BladeTypes.Categories.Width
        Icon = $script:BladeTypes.Categories.Icon
    }
    $searchBlade = @{
        Type = "Search"
        Title = $script:BladeTypes.Search.Title
        Width = $script:BladeTypes.Search.Width
        Icon = $script:BladeTypes.Search.Icon
    }
    $script:State.OpenBlades = @($categoriesBlade, $searchBlade)
    $script:State.ActiveBlade = 1  # Start with search blade active
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