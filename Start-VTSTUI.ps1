#Requires -Version 5.1
<#
.SYNOPSIS
    Launcher for VTS Tools Terminal User Interface
.DESCRIPTION
    Quick launcher script for the VTS Tools TUI
#>

# Get the directory where this script is located
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Launch the TUI
& "$scriptPath\VTSToolsTUI.ps1"