#Requires -Version 5.1
<#
.SYNOPSIS
    Terminal User Interface for VTS Tools
.DESCRIPTION
    A beautiful terminal interface for browsing and executing VTS PowerShell tools
    organized by category. Uses gum and psCandy for enhanced visual appeal.
#>

# Check for required modules and tools
function Test-Requirements {
    $missingRequirements = @()
    $hasGum = $false
    $hasPsCandy = $false
    
    # Check for psCandy module
    if (-not (Get-Module -ListAvailable -Name psCandy)) {
        $missingRequirements += "psCandy PowerShell module (optional)"
    } else {
        $hasPsCandy = $true
    }
    
    # Check for gum (after trying to initialize it)
    $gumPath = Get-Command gum -ErrorAction SilentlyContinue
    if (-not $gumPath) {
        $missingRequirements += "gum CLI tool"
    } else {
        $hasGum = $true
    }
    
    if ($missingRequirements.Count -gt 0) {
        Write-Host "`n⚠️  Missing Requirements:" -ForegroundColor Yellow
        $missingRequirements | ForEach-Object { Write-Host "  • $_" -ForegroundColor Red }
        Write-Host "`nInstallation Instructions:" -ForegroundColor Cyan
        Write-Host "  • psCandy: Install-Module psCandy -Scope CurrentUser" -ForegroundColor White
        Write-Host "  • gum: Run .\Install-VTSTUIRequirements.ps1" -ForegroundColor White
        Write-Host ""
        
        # If gum is missing, offer to run in fallback mode
        if (-not $hasGum) {
            Write-Host "Alternative: Run in fallback mode without gum (basic functionality)" -ForegroundColor Cyan
            $choice = Read-Host "Continue without gum? (y/n)"
            if ($choice -eq 'y' -or $choice -eq 'Y') {
                return $true
            }
        }
        return $false
    }
    return $true
}

# Import psCandy if available
if (Get-Module -ListAvailable -Name psCandy) {
    Import-Module psCandy -ErrorAction SilentlyContinue
}

# Function to find and add gum to PATH
function Initialize-GumPath {
    # Check if gum is already available
    if (Get-Command gum -ErrorAction SilentlyContinue) {
        return $true
    }
    
    # Common gum installation paths
    $gumPaths = @(
        "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\charmbracelet.gum_Microsoft.Winget.Source_8wekyb3d8bbwe\gum_0.16.2_Windows_x86_64",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\charmbracelet.gum_Microsoft.Winget.Source_8wekyb3d8bbwe",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Links",
        "$env:PROGRAMFILES\gum",
        "$env:PROGRAMFILES(X86)\gum",
        "$env:USERPROFILE\.local\bin"
    )
    
    foreach ($path in $gumPaths) {
        if (Test-Path "$path\gum.exe") {
            Write-Host "Found gum.exe at: $path" -ForegroundColor Green
            # Add to PATH for this session (ensure no duplicates)
            if ($env:Path -notlike "*$path*") {
                $env:Path = "$env:Path;$path"
                Write-Host "Added $path to PATH for this session" -ForegroundColor Green
            }
            # Test if gum is now accessible
            if (Get-Command gum -ErrorAction SilentlyContinue) {
                Write-Host "gum is now accessible!" -ForegroundColor Green
                return $true
            } else {
                Write-Host "gum found but still not accessible, trying direct path..." -ForegroundColor Yellow
                # Set a global variable for direct path access
                $global:GumPath = "$path\gum.exe"
                return $true
            }
        }
    }
    
    return $false
}

# Helper function to call gum with fallback to direct path
function Invoke-Gum {
    param(
        [string[]]$Arguments
    )
    
    if (Get-Command gum -ErrorAction SilentlyContinue) {
        return & gum @Arguments
    } elseif ($global:GumPath -and (Test-Path $global:GumPath)) {
        return & $global:GumPath @Arguments
    } else {
        throw "gum not available"
    }
}

# Fallback functions for when gum is not available
function Show-MenuFallback {
    param(
        [string]$Header,
        [array]$Options
    )
    
    Write-Host "`n$Header" -ForegroundColor Cyan
    Write-Host ("=" * $Header.Length) -ForegroundColor DarkCyan
    
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "  $($i + 1). $($Options[$i])" -ForegroundColor White
    }
    
    do {
        $choice = Read-Host "`nEnter your choice (1-$($Options.Count))"
        $index = $choice -as [int]
    } while ($index -lt 1 -or $index -gt $Options.Count)
    
    return $Options[$index - 1]
}

function Get-InputFallback {
    param(
        [string]$Prompt,
        [string]$Placeholder = ""
    )
    
    $displayPrompt = $Prompt
    if ($Placeholder) {
        $displayPrompt += " [$Placeholder]"
    }
    $displayPrompt += ": "
    
    return Read-Host $displayPrompt
}

# Function to parse PowerShell function files
function Get-FunctionInfo {
    param([string]$Path)
    
    $content = Get-Content $Path -Raw
    $functionInfo = @{
        Name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        Path = $Path
        Category = "Uncategorized"
        Synopsis = ""
        Description = ""
        Examples = @()
        Parameters = @()
    }
    
    # Extract category from .LINK
    if ($content -match '\.LINK\s*\r?\n\s*([^\r\n]+)') {
        $functionInfo.Category = $matches[1].Trim()
    }
    
    # Extract synopsis
    if ($content -match '\.SYNOPSIS\s*\r?\n\s*([^\r\n]+)') {
        $functionInfo.Synopsis = $matches[1].Trim()
    }
    
    # Extract description
    if ($content -match '\.DESCRIPTION\s*\r?\n([\s\S]*?)(?=\r?\n\s*\.|$)') {
        $desc = $matches[1] -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $functionInfo.Description = ($desc -join ' ').Trim()
    }
    
    # Extract examples
    $exampleMatches = [regex]::Matches($content, '\.EXAMPLE\s*\r?\n([\s\S]*?)(?=\r?\n\s*\.|$)')
    foreach ($match in $exampleMatches) {
        $exampleText = $match.Groups[1].Value.Trim()
        if ($exampleText) {
            $functionInfo.Examples += $exampleText
        }
    }
    
    # Extract parameters
    $paramMatches = [regex]::Matches($content, '\.PARAMETER\s+(\w+)\s*\r?\n\s*([^\r\n]+)')
    foreach ($match in $paramMatches) {
        $paramInfo = @{
            Name = $match.Groups[1].Value
            Description = $match.Groups[2].Value.Trim()
            Mandatory = $false
            Type = 'string'
        }
        
        # Skip complex parameter parsing to avoid regex issues
        # Just use the defaults: Mandatory = $false, Type = 'string'
        
        $functionInfo.Parameters += $paramInfo
    }
    
    return $functionInfo
}

# Function to display the main menu
function Show-MainMenu {
    param($Categories)
    
    Clear-Host
    
    # Create the header - simplified to avoid encoding issues
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "        VTS Tools Terminal UI          " -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    Write-Host ""
    
    # Use gum to select category or fallback
    $categoryList = $Categories + "Exit"
    
    try {
        $selected = $categoryList | Invoke-Gum -Arguments @("choose", "--header", "Select a category:", "--height", "15")
    } catch {
        $selected = Show-MenuFallback -Header "Select a category:" -Options $categoryList
    }
    
    return $selected
}

# Function to display functions in a category
function Show-CategoryFunctions {
    param(
        [string]$Category,
        [array]$Functions
    )
    
    Clear-Host
    
    # Header - simplified
    Write-Host "Category: $Category" -ForegroundColor Cyan
    Write-Host ("=" * 50) -ForegroundColor DarkCyan
    
    Write-Host ""
    
    # Prepare function list with descriptions
    $functionList = @()
    foreach ($func in $Functions) {
        $displayText = "$($func.Name)"
        if ($func.Synopsis) {
            $synopsis = $func.Synopsis
            if ($synopsis.Length -gt 50) {
                $synopsis = $synopsis.Substring(0, 47) + "..."
            }
            $displayText += " - $synopsis"
        }
        $functionList += $displayText
    }
    $functionList += "← Back to Categories"
    
    # Use gum to select function or fallback
    try {
        $selected = $functionList | Invoke-Gum -Arguments @("choose", "--header", "Select a function:", "--height", "20")
    } catch {
        $selected = Show-MenuFallback -Header "Select a function:" -Options $functionList
    }
    
    if ($selected -eq "← Back to Categories") {
        return $null
    }
    
    # Extract function name from selection
    $functionName = ($selected -split ' - ')[0]
    return $Functions | Where-Object { $_.Name -eq $functionName } | Select-Object -First 1
}

# Function to display function details
function Show-FunctionDetails {
    param($Function)
    
    Clear-Host
    
    # Header with function name - simplified
    Write-Host "Function: $($Function.Name)" -ForegroundColor Green
    Write-Host ("=" * 50) -ForegroundColor DarkGreen
    
    Write-Host ""
    
    # Synopsis
    if ($Function.Synopsis) {
        Write-Host "Synopsis:" -ForegroundColor Yellow
        Write-Host "  $($Function.Synopsis)" -ForegroundColor White
        Write-Host ""
    }
    
    # Description
    if ($Function.Description) {
        Write-Host "Description:" -ForegroundColor Yellow
        $descLines = $Function.Description -split '(?<=\.|!|\?)(?=\s)' | Where-Object { $_.Trim() }
        foreach ($line in $descLines) {
            Write-Host "  $($line.Trim())" -ForegroundColor White
        }
        Write-Host ""
    }
    
    # Parameters
    if ($Function.Parameters.Count -gt 0) {
        Write-Host "Parameters:" -ForegroundColor Yellow
        foreach ($param in $Function.Parameters) {
            $mandatoryText = if ($param.Mandatory) { "[Required]" } else { "[Optional]" }
            Write-Host "  -$($param.Name) " -ForegroundColor Cyan -NoNewline
            Write-Host "$mandatoryText " -ForegroundColor $(if ($param.Mandatory) { "Red" } else { "DarkGray" }) -NoNewline
            Write-Host "<$($param.Type)>" -ForegroundColor Magenta
            if ($param.Description) {
                Write-Host "    $($param.Description)" -ForegroundColor Gray
            }
        }
        Write-Host ""
    }
    
    # Examples
    if ($Function.Examples.Count -gt 0) {
        Write-Host "Examples:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $Function.Examples.Count; $i++) {
            Write-Host "  Example $($i + 1):" -ForegroundColor DarkYellow
            $exampleLines = $Function.Examples[$i] -split '\r?\n'
            foreach ($line in $exampleLines) {
                if ($line.Trim()) {
                    Write-Host "    $($line.Trim())" -ForegroundColor White
                }
            }
            Write-Host ""
        }
    }
    
    # Action menu
    $actions = @("Execute Function", "Execute with Parameters", "View Source", "← Back to Functions")
    
    try {
        $selectedAction = $actions | Invoke-Gum -Arguments @("choose", "--header", "Choose an action:")
    } catch {
        $selectedAction = Show-MenuFallback -Header "Choose an action:" -Options $actions
    }
    
    return $selectedAction
}

# Function to execute a PowerShell function
function Invoke-VTSFunction {
    param(
        $Function,
        [switch]$WithParameters
    )
    
    # Load the function
    . $Function.Path
    
    if ($WithParameters -and $Function.Parameters.Count -gt 0) {
        $parameters = @{}
        
        Write-Host "`nEnter parameter values:" -ForegroundColor Cyan
        foreach ($param in $Function.Parameters) {
            $prompt = "  -$($param.Name) [$($param.Type)]"
            if ($param.Mandatory) {
                $prompt += " (Required)"
            }
            $prompt += ": "
            
            # Use gum input for parameter collection or fallback
            try {
                $value = Invoke-Gum -Arguments @("input", "--placeholder", "Enter value for $($param.Name)", "--prompt", "$prompt")
            } catch {
                $value = Get-InputFallback -Prompt "$prompt" -Placeholder "Enter value for $($param.Name)"
            }
            
            if ($value -and $value.Trim()) {
                # Try to convert value to appropriate type
                switch -Regex ($param.Type) {
                    'int|int32' { $parameters[$param.Name] = [int]$value }
                    'bool|boolean' { $parameters[$param.Name] = [bool]$value }
                    'switch' { if ($value -eq 'true' -or $value -eq '1') { $parameters[$param.Name] = $true } }
                    default { $parameters[$param.Name] = $value }
                }
            }
        }
        
        Write-Host "`nExecuting: $($Function.Name) with parameters..." -ForegroundColor Green
        & $Function.Name @parameters
    } else {
        Write-Host "`nExecuting: $($Function.Name)..." -ForegroundColor Green
        & $Function.Name
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Main execution
function Start-VTSToolsTUI {
    # Try to initialize gum path first
    Initialize-GumPath | Out-Null
    
    if (-not (Test-Requirements)) {
        return
    }
    
    # Get all function files
    $functionPath = Join-Path $PSScriptRoot "functions"
    if (-not (Test-Path $functionPath)) {
        Write-Host "Functions directory not found at: $functionPath" -ForegroundColor Red
        return
    }
    
    Write-Host "Loading functions..." -ForegroundColor Yellow
    $allFunctions = Get-ChildItem -Path $functionPath -Filter "*.ps1" | ForEach-Object {
        Get-FunctionInfo -Path $_.FullName
    }
    
    # Group functions by category
    $categories = $allFunctions | Group-Object -Property Category | Sort-Object Name
    
    # Main loop
    while ($true) {
        $selectedCategory = Show-MainMenu -Categories ($categories | ForEach-Object { $_.Name })
        
        if ($selectedCategory -eq "Exit" -or -not $selectedCategory) {
            Write-Host "`nThank you for using VTS Tools TUI!" -ForegroundColor Green
            break
        }
        
        # Get functions in selected category
        $categoryFunctions = $allFunctions | Where-Object { $_.Category -eq $selectedCategory } | Sort-Object Name
        
        while ($true) {
            $selectedFunction = Show-CategoryFunctions -Category $selectedCategory -Functions $categoryFunctions
            
            if (-not $selectedFunction) {
                break # Back to main menu
            }
            
            while ($true) {
                $action = Show-FunctionDetails -Function $selectedFunction
                
                switch ($action) {
                    "Execute Function" {
                        Invoke-VTSFunction -Function $selectedFunction
                    }
                    "Execute with Parameters" {
                        Invoke-VTSFunction -Function $selectedFunction -WithParameters
                    }
                    "View Source" {
                        Clear-Host
                        Write-Host "Source: $($selectedFunction.Path)" -ForegroundColor Cyan
                        Write-Host ("=" * 70) -ForegroundColor DarkCyan
                        
                        try {
                            Get-Content $selectedFunction.Path | Invoke-Gum -Arguments @("pager")
                        } catch {
                            Get-Content $selectedFunction.Path | Out-Host
                            Write-Host "`nPress any key to continue..." -ForegroundColor DarkGray
                            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                        }
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

# Start the TUI
Start-VTSToolsTUI