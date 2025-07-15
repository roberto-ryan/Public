#requires -version 5.1
<#
.SYNOPSIS
    Interactive console menu for VTS Tools PowerShell functions using psCandy
.DESCRIPTION
    Creates a beautiful, navigable menu system for all VTS Tools PowerShell functions
    organized by category with arrow key navigation support via psCandy module.
.EXAMPLE
    .\Start-VTSMenu.ps1
.LINK
    Menu
#>

# Set UTF-8 encoding for proper display
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Check if psCandy module is installed
if (-not (Get-Module -ListAvailable -Name psCandy)) {
    Write-Host "psCandy module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name psCandy -Force -Scope CurrentUser
        Write-Host "psCandy module installed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install psCandy module: $_"
        exit 1
    }
}

# Import psCandy module
Import-Module psCandy -Force

# Import the module classes after loading
using module psCandy

# Define VTS Theme
$VTSTheme = @{
    "list" = @{
        "SearchColor" = "Cyan"
        "SelectedColor" = "Yellow"
        "SelectedStyle" = [Styles]::Bold
        "FilterColor" = "Green"
        "NoFilterColor" = "Red"
        "FilterStyle" = [Styles]::Underline
        "Checked" = "◉"
        "Unchecked" = "○"
    }
    "spinner" = @{
        "spincolor" = "Cyan"
        "spinType" = "Dots"
    }
    "choice" = @{
        "SelectedForeground" = "Black"
        "SelectedBackground" = "Yellow"
        "OptionColor" = "Cyan"
        "MessageColor" = "White"
    }
}

# Function to get script information
function Get-ScriptInfo {
    param([string]$FilePath)
    
    try {
        $content = Get-Content $FilePath -Raw
        $synopsis = ""
        $category = "Utilities"
        
        # Extract synopsis
        if ($content -match '\.SYNOPSIS\s*\n\s*(.+?)(?=\n\s*\.|\n\s*#>|\n\s*$)') {
            $synopsis = $matches[1].Trim()
        }
        
        # Extract category from .LINK
        if ($content -match '\.LINK\s*\n\s*(.+?)(?=\n\s*\.|\n\s*#>|\n\s*$)') {
            $category = $matches[1].Trim()
        }
        
        # Fallback to filename if no synopsis
        if ([string]::IsNullOrWhiteSpace($synopsis)) {
            $synopsis = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        }
        
        return @{
            Name = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
            Synopsis = $synopsis
            Category = $category
            Path = $FilePath
        }
    }
    catch {
        return @{
            Name = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
            Synopsis = "PowerShell script"
            Category = "Utilities"
            Path = $FilePath
        }
    }
}

# Function to get category icon
function Get-CategoryIcon {
    param([string]$Category)
    
    switch ($Category.ToLower()) {
        "m365" { return "[M365]" }
        "network" { return "[NET]" }
        "system information" { return "[SYS]" }
        "device management" { return "[DEV]" }
        "utilities" { return "[UTIL]" }
        "active directory" { return "[AD]" }
        "security" { return "[SEC]" }
        "menu" { return "[MENU]" }
        default { return "[SCRIPT]" }
    }
}

# Function to get script icon
function Get-ScriptIcon {
    param([string]$ScriptName, [string]$Category)
    
    switch ($Category.ToLower()) {
        "m365" { return "[M365]" }
        "network" { return "[NET]" }
        "system information" { return "[SYS]" }
        "device management" { return "[DEV]" }
        "utilities" { return "[UTIL]" }
        "active directory" { return "[AD]" }
        "security" { return "[SEC]" }
        default { return "[PS]" }
    }
}

# Function to execute selected script
function Invoke-SelectedScript {
    param([string]$ScriptPath)
    
    try {
        Clear-Host
        Write-Host "=> Executing: $([System.IO.Path]::GetFileName($ScriptPath))" -ForegroundColor Green
        Write-Host "=" * 60 -ForegroundColor Cyan
        
        # Execute the script
        & $ScriptPath
        
        Write-Host "`n" + "=" * 60 -ForegroundColor Cyan
        Write-Host "=> Script execution completed." -ForegroundColor Green
        Write-Host "Press any key to return to menu..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    catch {
        Write-Error "=> Failed to execute script: $_"
        Write-Host "Press any key to return to menu..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# Function to show category menu
function Show-CategoryMenu {
    param([hashtable]$ScriptsByCategory)
    
    $categoryItems = [System.Collections.Generic.List[ListItem]]::new()
    
    # Add categories
    foreach ($category in ($ScriptsByCategory.Keys | Sort-Object)) {
        $icon = Get-CategoryIcon $category
        $count = $ScriptsByCategory[$category].Count
        $displayText = "$category ($count scripts)"
        $categoryItems.Add([ListItem]::new($displayText, $category, $icon, [System.Drawing.Color]::Cyan))
    }
    
    # Add exit option
    $categoryItems.Add([ListItem]::new("Exit", "EXIT", "[EXIT]", [System.Drawing.Color]::Red))
    
    $categoryList = [List]::new($categoryItems)
    $categoryList.LoadTheme($VTSTheme)
    $categoryList.SetHeight(15)
    $categoryList.SetTitle(">> VTS Tools - Select Category")
    $categoryList.SetLimit($true)
    
    return $categoryList.Display()
}

# Function to show scripts in category
function Show-ScriptsMenu {
    param([string]$Category, [array]$Scripts)
    
    $scriptItems = [System.Collections.Generic.List[ListItem]]::new()
    
    # Add back option
    $scriptItems.Add([ListItem]::new("<- Back to Categories", "BACK", "[BACK]", [System.Drawing.Color]::Orange))
    
    # Add scripts
    foreach ($script in ($Scripts | Sort-Object Name)) {
        $icon = Get-ScriptIcon $script.Name $script.Category
        $displayText = "$($script.Name) - $($script.Synopsis)"
        $scriptItems.Add([ListItem]::new($displayText, $script.Path, $icon, [System.Drawing.Color]::Green))
    }
    
    $scriptList = [List]::new($scriptItems)
    $scriptList.LoadTheme($VTSTheme)
    $scriptList.SetHeight(20)
    $scriptList.SetTitle(">> $Category Scripts")
    $scriptList.SetLimit($true)
    
    return $scriptList.Display()
}

# Main function
function Start-VTSMenu {
    # Clear screen and show header
    Clear-Host
    
    # Show VTS ASCII art header
    Write-Host @"
╔═══════════════════════════════════════════════════════════════════════════════╗
║                                                                               ║
║  ██╗   ██╗████████╗███████╗    ████████╗ ██████╗  ██████╗ ██╗     ███████╗   ║
║  ██║   ██║╚══██╔══╝██╔════╝    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██╔════╝   ║
║  ██║   ██║   ██║   ███████╗       ██║   ██║   ██║██║   ██║██║     ███████╗   ║
║  ╚██╗ ██╔╝   ██║   ╚════██║       ██║   ██║   ██║██║   ██║██║     ╚════██║   ║
║   ╚████╔╝    ██║   ███████║       ██║   ╚██████╔╝╚██████╔╝███████╗███████║   ║
║    ╚═══╝     ╚═╝   ╚══════╝       ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚══════╝   ║
║                                                                               ║
║                        Interactive PowerShell Menu                           ║
║                                                                               ║
╚═══════════════════════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan
    
    # Get functions directory
    $functionsPath = Join-Path $PSScriptRoot "functions"
    
    if (-not (Test-Path $functionsPath)) {
        Write-Error "Functions directory not found at: $functionsPath"
        return
    }
    
    # Get all PowerShell scripts
    $scripts = Get-ChildItem $functionsPath -Filter "*.ps1" | ForEach-Object {
        Get-ScriptInfo $_.FullName
    }
    
    # Group scripts by category
    $scriptsByCategory = @{}
    foreach ($script in $scripts) {
        if (-not $scriptsByCategory.ContainsKey($script.Category)) {
            $scriptsByCategory[$script.Category] = @()
        }
        $scriptsByCategory[$script.Category] += $script
    }
    
    # Main menu loop
    while ($true) {
        $categoryChoice = Show-CategoryMenu $scriptsByCategory
        
        if (-not $categoryChoice -or $categoryChoice.Value -eq "EXIT") {
            Clear-Host
            Write-Host "Thank you for using VTS Tools!" -ForegroundColor Cyan
            break
        }
        
        # Show scripts in selected category
        while ($true) {
            $scriptChoice = Show-ScriptsMenu $categoryChoice.Value $scriptsByCategory[$categoryChoice.Value]
            
            if (-not $scriptChoice -or $scriptChoice.Value -eq "BACK") {
                break
            }
            
            # Execute selected script
            Invoke-SelectedScript $scriptChoice.Value
        }
    }
}

# Start the menu
Start-VTSMenu