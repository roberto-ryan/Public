function Get-vtsFileContentMatch {
  <#
  .SYNOPSIS
  This function searches for a specific pattern in the content of files in a given directory and optionally exports the results to a CSV file.
  
  .DESCRIPTION
  The Get-vtsFileContentMatch function takes a directory path, a pattern to match, and an optional array of file types to exclude. It recursively searches through all files in the specified directory (excluding the specified file types), and returns a custom object for each line in each file that matches the specified pattern. The custom object contains the full path of the file and the line that matched the pattern. The function can also export the results to a CSV file if the path to the CSV file is provided.
  
  .PARAMETER Path
  The path to the directory to search. This parameter is mandatory.
  
  .PARAMETER Pattern
  The pattern or word to match in the file content. This parameter is mandatory.
  
  .PARAMETER Exclude
  An array of file types to exclude from the search. The default value is '*.exe', '*.dll'. This parameter is optional.
  
  .PARAMETER ExportToCsv
  A boolean value indicating whether to export the results to a CSV file. This parameter is optional.
  
  .PARAMETER CsvPath
  The path to the CSV file to export the results to. This parameter is optional.
  
  .EXAMPLE
  Get-vtsFileContentMatch -Path "C:\Users\Username\Documents" -Pattern "error" -ExportToCsv $true -CsvPath "C:\Users\Username\Documents\results.csv"
  
  This example searches for the word "error" in all files in the "C:\Users\Username\Documents" directory, excluding .exe and .dll files, and exports the results to a CSV file.
  
  .LINK
  Utilities
  #>
  param (
    [Parameter(Mandatory = $true, HelpMessage = "Please provide the path to the directory.")]
    [string]$Path,
    [Parameter(Mandatory = $true, HelpMessage = "Please provide the pattern or word to match.")]
    [string]$Pattern,
    [Parameter(Mandatory = $false, HelpMessage = "Please provide the file types to exclude. Default is '*.exe', '*.dll'.")]
    [string[]]$Exclude = @("*.exe", "*.dll"),
    [Parameter(Mandatory = $false, HelpMessage = "Please indicate if you want to export the results to a CSV file.")]
    [bool]$ExportToCsv = $false,
    [Parameter(Mandatory = $false, HelpMessage = "Please provide the path to the CSV file.")]
    [string]$CsvPath = "C:\temp\$(Get-Date -f yyyy-MM-dd-HH-mm)-$($env:COMPUTERNAME)-FilesIncluding-$Pattern.csv"
  )

  $Results = @()
  Get-ChildItem -Path $Path -Recurse -File -Exclude $Exclude | ForEach-Object {
    $filePath = $_.FullName
    Get-Content -Path $filePath | ForEach-Object {
      if ($_ -match $Pattern) {
        $Result = [pscustomobject]@{
          FilePath = $filePath
          Match    = $_
        }
        $Result | Format-List
        $Results += $Result
      }
    }
  }

  if ($ExportToCsv) {
    $Results | Export-Csv -Path $CsvPath -NoTypeInformation -Force
    if ($?) { Write-Host "Results exported to $CsvPath" -ForegroundColor Yellow }
  }
}

