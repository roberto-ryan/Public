#Requires -Version 5.1

# Simple VTS Tools Menu
# No fancy features, just basic PowerShell that works

function Show-Title {
    param([string]$Text)
    Clear-Host
    Write-Host ""
    Write-Host "================================" -ForegroundColor Cyan
    Write-Host " $Text" -ForegroundColor Cyan
    Write-Host "================================" -ForegroundColor Cyan
    Write-Host ""
}

function Get-UserChoice {
    param([array]$Options)
    
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "  $($i + 1). $($Options[$i])" -ForegroundColor White
    }
    Write-Host ""
    
    do {
        $choice = Read-Host "Enter choice (1-$($Options.Count))"
        $num = $choice -as [int]
    } while ($num -lt 1 -or $num -gt $Options.Count)
    
    return $Options[$num - 1]
}

function Get-Functions {
    $functionsPath = Join-Path $PSScriptRoot "functions"
    
    if (-not (Test-Path $functionsPath)) {
        Write-Host "Functions folder not found!" -ForegroundColor Red
        return @()
    }
    
    $functions = @()
    $files = Get-ChildItem -Path $functionsPath -Filter "*.ps1"
    
    foreach ($file in $files) {
        $content = Get-Content $file.FullName -Raw
        $name = $file.BaseName
        $category = "General"
        $description = ""
        
        # Get category from .LINK
        if ($content -match '\.LINK\s*\n\s*(.+)') {
            $category = $matches[1].Trim()
        }
        
        # Get description from .SYNOPSIS
        if ($content -match '\.SYNOPSIS\s*\n\s*(.+)') {
            $description = $matches[1].Trim()
        }
        
        $functions += [PSCustomObject]@{
            Name = $name
            Category = $category
            Description = $description
            Path = $file.FullName
        }
    }
    
    return $functions
}

function Show-MainMenu {
    param([array]$Functions)
    
    Show-Title "VTS Tools Menu"
    
    # Group by category
    $categories = $Functions | Group-Object Category | Sort-Object Name
    
    Write-Host "Categories:" -ForegroundColor Yellow
    Write-Host ""
    
    $menuItems = @()
    foreach ($cat in $categories) {
        $menuItems += "$($cat.Name) ($($cat.Count) tools)"
    }
    $menuItems += "Exit"
    
    $selected = Get-UserChoice -Options $menuItems
    
    if ($selected -eq "Exit") {
        return "EXIT"
    }
    
    # Extract category name
    $categoryName = $selected -replace ' \(\d+ tools\)$', ''
    return $categoryName
}

function Show-CategoryMenu {
    param([string]$CategoryName, [array]$Functions)
    
    Show-Title "Category: $CategoryName"
    
    $categoryFunctions = $Functions | Where-Object { $_.Category -eq $CategoryName } | Sort-Object Name
    
    Write-Host "Available tools:" -ForegroundColor Yellow
    Write-Host ""
    
    $menuItems = @()
    foreach ($func in $categoryFunctions) {
        $item = $func.Name
        if ($func.Description) {
            $item += " - $($func.Description)"
        }
        $menuItems += $item
    }
    $menuItems += "Back to Main Menu"
    
    $selected = Get-UserChoice -Options $menuItems
    
    if ($selected -eq "Back to Main Menu") {
        return $null
    }
    
    # Find the function
    $functionName = ($selected -split ' - ')[0]
    $function = $categoryFunctions | Where-Object { $_.Name -eq $functionName }
    return $function
}

function Show-FunctionMenu {
    param([object]$Function)
    
    Show-Title "Tool: $($Function.Name)"
    
    if ($Function.Description) {
        Write-Host "Description: $($Function.Description)" -ForegroundColor Green
        Write-Host ""
    }
    
    Write-Host "What do you want to do?" -ForegroundColor Yellow
    Write-Host ""
    
    $actions = @("Run this tool", "View source code", "Back to category")
    $selected = Get-UserChoice -Options $actions
    
    return $selected
}

function Run-Function {
    param([object]$Function)
    
    Show-Title "Running: $($Function.Name)"
    
    try {
        # Load the function
        . $Function.Path
        
        Write-Host "Starting $($Function.Name)..." -ForegroundColor Green
        Write-Host ""
        
        # Run it
        & $Function.Name
        
        Write-Host ""
        Write-Host "Completed!" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Show-SourceCode {
    param([object]$Function)
    
    Show-Title "Source: $($Function.Name)"
    
    Write-Host "File: $($Function.Path)" -ForegroundColor Gray
    Write-Host ""
    
    try {
        $lines = Get-Content $Function.Path
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $lineNum = ($i + 1).ToString().PadLeft(3)
            Write-Host "$lineNum : $($lines[$i])"
        }
    }
    catch {
        Write-Host "Cannot read file: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Main program
function Start-Menu {
    Write-Host "Loading tools..." -ForegroundColor Yellow
    $functions = Get-Functions
    
    if ($functions.Count -eq 0) {
        Write-Host "No tools found!" -ForegroundColor Red
        return
    }
    
    Write-Host "Found $($functions.Count) tools" -ForegroundColor Green
    Start-Sleep -Seconds 1
    
    # Main loop
    while ($true) {
        $selectedCategory = Show-MainMenu -Functions $functions
        
        if ($selectedCategory -eq "EXIT") {
            Show-Title "Goodbye!"
            break
        }
        
        # Category loop
        while ($true) {
            $selectedFunction = Show-CategoryMenu -CategoryName $selectedCategory -Functions $functions
            
            if ($selectedFunction -eq $null) {
                break
            }
            
            # Function loop
            while ($true) {
                $action = Show-FunctionMenu -Function $selectedFunction
                
                if ($action -eq "Run this tool") {
                    Run-Function -Function $selectedFunction
                }
                elseif ($action -eq "View source code") {
                    Show-SourceCode -Function $selectedFunction
                }
                elseif ($action -eq "Back to category") {
                    break
                }
            }
        }
    }
}

# Start the menu
Start-Menu