#Requires -Version 5.1
<#
.SYNOPSIS
    Simple Terminal User Interface for VTS Tools
.DESCRIPTION
    A clean, streamlined terminal interface for browsing and executing VTS PowerShell tools.
    Designed for Windows PowerShell 5.1 with minimal dependencies.
#>

# Global variables
$Global:CurrentCategory = ""
$Global:AllFunctions = @()
$Global:Categories = @()

# Color scheme
$Colors = @{
    Header = "Cyan"
    SubHeader = "DarkCyan"
    Success = "Green"
    Warning = "Yellow"
    Error = "Red"
    Info = "White"
    Prompt = "Magenta"
    Highlight = "DarkYellow"
}

# Function to display a clean header
function Show-Header {
    param([string]$Title, [string]$Color = $Colors.Header)
    
    Clear-Host
    $line = "=" * 60
    Write-Host $line -ForegroundColor $Color
    Write-Host $Title.PadLeft(($line.Length + $Title.Length) / 2) -ForegroundColor $Color
    Write-Host $line -ForegroundColor $Color
    Write-Host ""
}

# Function to show a simple menu
function Show-Menu {
    param(
        [string]$Title,
        [array]$Options,
        [switch]$ShowNumbers = $true
    )
    
    Write-Host $Title -ForegroundColor $Colors.SubHeader
    Write-Host ("-" * $Title.Length) -ForegroundColor $Colors.SubHeader
    Write-Host ""
    
    if ($ShowNumbers) {
        for ($i = 0; $i -lt $Options.Count; $i++) {
            Write-Host "  [$($i + 1)] $($Options[$i])" -ForegroundColor $Colors.Info
        }
    } else {
        foreach ($option in $Options) {
            Write-Host "  • $option" -ForegroundColor $Colors.Info
        }
    }
    Write-Host ""
    
    if ($ShowNumbers) {
        do {
            $choice = Read-Host "Enter your choice (1-$($Options.Count))"
            $index = $choice -as [int]
        } while ($index -lt 1 -or $index -gt $Options.Count)
        
        return $Options[$index - 1]
    }
}

# Function to parse function information
function Get-FunctionInfo {
    param([string]$FilePath)
    
    $content = Get-Content $FilePath -Raw -ErrorAction SilentlyContinue
    if (-not $content) { return $null }
    
    $functionInfo = @{
        Name = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        Path = $FilePath
        Category = "General"
        Synopsis = ""
        Description = ""
        Examples = @()
    }
    
    # Extract category from .LINK
    if ($content -match '\.LINK\s*\r?\n\s*([^\r\n]+)') {
        $functionInfo.Category = $matches[1].Trim()
    }
    
    # Extract synopsis
    if ($content -match '\.SYNOPSIS\s*\r?\n\s*([^\r\n]+)') {
        $functionInfo.Synopsis = $matches[1].Trim()
    }
    
    # Extract description (first paragraph only)
    if ($content -match '\.DESCRIPTION\s*\r?\n\s*([^\r\n]+)') {
        $functionInfo.Description = $matches[1].Trim()
    }
    
    # Extract examples (simplified)
    $exampleMatches = [regex]::Matches($content, '\.EXAMPLE\s*\r?\n\s*([^\r\n]+)')
    foreach ($match in $exampleMatches) {
        $example = $match.Groups[1].Value.Trim()
        if ($example -and $example -ne "PS>") {
            $functionInfo.Examples += $example
        }
    }
    
    return $functionInfo
}

# Function to load all functions
function Initialize-Functions {
    $functionPath = Join-Path $PSScriptRoot "functions"
    
    if (-not (Test-Path $functionPath)) {
        Write-Host "Functions directory not found at: $functionPath" -ForegroundColor $Colors.Error
        return $false
    }
    
    Write-Host "Loading functions..." -ForegroundColor $Colors.Info
    
    $Global:AllFunctions = Get-ChildItem -Path $functionPath -Filter "*.ps1" | ForEach-Object {
        $funcInfo = Get-FunctionInfo -FilePath $_.FullName
        if ($funcInfo) { $funcInfo }
    } | Where-Object { $_ -ne $null }
    
    if ($Global:AllFunctions.Count -eq 0) {
        Write-Host "No functions found!" -ForegroundColor $Colors.Error
        return $false
    }
    
    # Group by category
    $Global:Categories = $Global:AllFunctions | Group-Object -Property Category | Sort-Object Name
    
    Write-Host "Loaded $($Global:AllFunctions.Count) functions in $($Global:Categories.Count) categories" -ForegroundColor $Colors.Success
    Start-Sleep -Milliseconds 800
    
    return $true
}

# Main menu
function Show-MainMenu {
    Show-Header "VTS Tools - Main Menu"
    
    Write-Host "Available Categories:" -ForegroundColor $Colors.SubHeader
    Write-Host "--------------------" -ForegroundColor $Colors.SubHeader
    Write-Host ""
    
    $menuOptions = @()
    $menuOptions += $Global:Categories | ForEach-Object { "$($_.Name) ($($_.Count) functions)" }
    $menuOptions += "Exit"
    
    $selected = Show-Menu -Title "Select a category:" -Options $menuOptions
    
    if ($selected -eq "Exit") {
        return "Exit"
    }
    
    # Extract category name (remove count)
    $categoryName = $selected -replace ' \(\d+ functions\)$', ''
    return $categoryName
}

# Category menu
function Show-CategoryMenu {
    param([string]$CategoryName)
    
    $categoryFunctions = $Global:AllFunctions | Where-Object { $_.Category -eq $CategoryName } | Sort-Object Name
    
    Show-Header "Category: $CategoryName"
    
    Write-Host "Functions in this category:" -ForegroundColor $Colors.SubHeader
    Write-Host "--------------------------" -ForegroundColor $Colors.SubHeader
    Write-Host ""
    
    $menuOptions = @()
    foreach ($func in $categoryFunctions) {
        $displayName = $func.Name
        if ($func.Synopsis) {
            $synopsis = $func.Synopsis
            if ($synopsis.Length -gt 40) {
                $synopsis = $synopsis.Substring(0, 37) + "..."
            }
            $displayName += " - $synopsis"
        }
        $menuOptions += $displayName
    }
    $menuOptions += "← Back to Main Menu"
    
    $selected = Show-Menu -Title "Select a function:" -Options $menuOptions
    
    if ($selected -eq "← Back to Main Menu") {
        return $null
    }
    
    # Extract function name
    $functionName = ($selected -split ' - ')[0]
    return $categoryFunctions | Where-Object { $_.Name -eq $functionName } | Select-Object -First 1
}

# Function details
function Show-FunctionDetails {
    param($Function)
    
    Show-Header "Function: $($Function.Name)"
    
    if ($Function.Synopsis) {
        Write-Host "Synopsis:" -ForegroundColor $Colors.Highlight
        Write-Host "  $($Function.Synopsis)" -ForegroundColor $Colors.Info
        Write-Host ""
    }
    
    if ($Function.Description) {
        Write-Host "Description:" -ForegroundColor $Colors.Highlight
        Write-Host "  $($Function.Description)" -ForegroundColor $Colors.Info
        Write-Host ""
    }
    
    if ($Function.Examples.Count -gt 0) {
        Write-Host "Examples:" -ForegroundColor $Colors.Highlight
        for ($i = 0; $i -lt $Function.Examples.Count; $i++) {
            Write-Host "  $($i + 1). $($Function.Examples[$i])" -ForegroundColor $Colors.Info
        }
        Write-Host ""
    }
    
    $actions = @("Run Function", "View Source Code", "← Back to Functions")
    $selectedAction = Show-Menu -Title "What would you like to do?" -Options $actions
    
    return $selectedAction
}

# Execute function
function Invoke-Function {
    param($Function)
    
    Show-Header "Executing: $($Function.Name)"
    
    try {
        # Load and execute the function
        . $Function.Path
        
        Write-Host "Running $($Function.Name)..." -ForegroundColor $Colors.Info
        Write-Host ""
        
        # Execute the function
        & $Function.Name
        
        Write-Host ""
        Write-Host "Function completed." -ForegroundColor $Colors.Success
    }
    catch {
        Write-Host "Error executing function: $($_.Exception.Message)" -ForegroundColor $Colors.Error
    }
    
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor $Colors.Prompt
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# View source code
function Show-SourceCode {
    param($Function)
    
    Show-Header "Source Code: $($Function.Name)"
    
    try {
        $content = Get-Content $Function.Path
        
        Write-Host "File: $($Function.Path)" -ForegroundColor $Colors.SubHeader
        Write-Host ("-" * 60) -ForegroundColor $Colors.SubHeader
        Write-Host ""
        
        # Display with line numbers
        for ($i = 0; $i -lt $content.Count; $i++) {
            $lineNumber = ($i + 1).ToString().PadLeft(3)
            Write-Host "$lineNumber : $($content[$i])" -ForegroundColor $Colors.Info
        }
    }
    catch {
        Write-Host "Error reading source file: $($_.Exception.Message)" -ForegroundColor $Colors.Error
    }
    
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor $Colors.Prompt
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Main application loop
function Start-VTSToolsTUI {
    # Initialize
    if (-not (Initialize-Functions)) {
        Write-Host "Press any key to exit..." -ForegroundColor $Colors.Prompt
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }
    
    # Main loop
    while ($true) {
        $selectedCategory = Show-MainMenu
        
        if ($selectedCategory -eq "Exit") {
            Show-Header "Thank you for using VTS Tools!"
            Start-Sleep -Milliseconds 1000
            break
        }
        
        # Category loop
        while ($true) {
            $selectedFunction = Show-CategoryMenu -CategoryName $selectedCategory
            
            if (-not $selectedFunction) {
                break # Back to main menu
            }
            
            # Function details loop
            while ($true) {
                $action = Show-FunctionDetails -Function $selectedFunction
                
                switch ($action) {
                    "Run Function" {
                        Invoke-Function -Function $selectedFunction
                    }
                    "View Source Code" {
                        Show-SourceCode -Function $selectedFunction
                    }
                    "← Back to Functions" {
                        break
                    }
                }
                
                if ($action -eq "← Back to Functions") {
                    break
                }
            }
        }
    }
}

# Start the application
Start-VTSToolsTUI