function Search-vtsAllLogs {
  <#
  .SYNOPSIS
  Searches Windows Event Logs for a specific term and exports the results to a CSV or HTML report.
  
  .DESCRIPTION
  The Search-vtsAllLogs function searches through Windows Event Logs for a specific term. The function allows the user to specify the logs to search, the term to search for, the type of report to generate (CSV or HTML), and the output directory for the report. If a date is provided, the function will only search logs from that date. If the date is set to 'today', the function will search logs from the current date.
  
  .PARAMETER SearchTerm
  The term to search for in the logs. This parameter is mandatory.
  
  .PARAMETER ReportType
  The type of report to generate. Valid options are 'csv' and 'html'. If not specified, the function will default to 'html' for results with less than or equal to 150 entries, and 'csv' for results with more than 150 entries.
  
  .PARAMETER OutputDirectory
  The directory where the report will be saved. If not specified, the function will default to 'C:\temp'.
  
  .PARAMETER Date
  The date of the logs to search. The date should be entered in the format: month/day/year. For example, 12/14/2023. You can also enter 'today' as a value to search logs from the current date.
  
  .EXAMPLE
  Search-vtsAllLogs -SearchTerm "Error" -ReportType "csv" -OutputDirectory "C:\temp" -Date "12/14/2023"
  
  This example searches all logs from 12/14/2023 for the term "Error" and generates a CSV report in the 'C:\temp' directory.
  
  .EXAMPLE
  Search-vtsAllLogs -SearchTerm "Warning" -Date "today"
  
  This example searches all logs from the current date for the term "Warning" and generates a report based on the number of results found.
  
  .LINK
  Log Management
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $SearchTerm,
    [ValidateSet("csv", "html")]
    [string]$ReportType,
    [string]$OutputDirectory = "C:\temp",
    [Parameter(HelpMessage = "Please enter the date in the format: month/day/year. For example, 12/14/2023. You can also enter 'today' as a value.")]
    $Date
  )

  if ($Date -eq "today") {
    $Date = Get-Date -f MM/dd/yy
  }
    
  # Create Output Dir if not exist
  if (!(Test-Path -Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
  }
  
  # Validate the log name
  $validLogNames = (Get-WinEvent -ListLog *).LogName 2>$null
  
  $LogTable = @()
  $key = 1

  foreach ($log in $validLogNames) {
    $LogTable += [pscustomobject]@{
      Key = $key
      Log = $log
    }
    $key++
  }
  
  if ($(whoami) -eq "nt authority\system") {
    $LogTable | Out-Host
    $userInput = Read-Host "Please input the log numbers you wish to search, separated by commas. Alternatively, input '*' to search all logs."
    if ("$userInput" -eq '*') {
      Write-Host "Searching all available logs..."
      $SelectedLogs = $LogTable.Log
    }
    else {
      Write-Host "Searching selected logs..."
      $SelectedLogs = $LogTable | Where-Object Key -in ("$userInput" -split ",") | Select-Object -ExpandProperty Log
    }
  }
  else {
    $SelectedLogs = $LogTable | Out-GridView -OutputMode Multiple | Select-Object -ExpandProperty Log
  }

  # Get the logs from the Event Viewer based on the provided log name
  $Results = @()
  $LogCount = $SelectedLogs.Count * $SearchTerm.Count
  $CurrentLog = 0
  foreach ($LogName in $SelectedLogs) {
    foreach ($Term in $SearchTerm) {
      $CurrentLog++
      Write-Host "Searching $LogName log for $Term... ($CurrentLog/$LogCount)" -ForegroundColor Yellow
      if ($Date) {
        $Results += Get-WinEvent -LogName "$LogName" -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated.Date -eq $Date } |
        Where-Object Message -like "*$Term*" |
        Tee-Object -Variable temp
      }
      else {
        $Results += Get-WinEvent -LogName "$LogName" -ErrorAction SilentlyContinue |
        Where-Object Message -like "*$Term*" |
        Tee-Object -Variable temp
      }
    }
  }
  
  if ("" -eq $ReportType) {
    if ($Results.Count -le 150) {
      $ReportType = "html"
    }
    else {
      $ReportType = "csv"
    }
  }
  
  $ReportPath = "C:\temp\$($env:COMPUTERNAME)-$(Get-Date -Format MM-dd-yy-mm-ss).$ReportType"

  switch ($ReportType) {
    csv { 
      $Results | Export-Csv $ReportPath -NoTypeInformation
      $openFile = Read-Host "Do you want to open the file? (Y/N)"
      if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ReportPath
      }
    }
    html { 
      # Check if PSWriteHTML module is installed, if not, install it
      if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
        Install-Module -Name PSWriteHTML -Force -Confirm:$false
      }
      
      # Export the results to an HTML file using the PSWriteHTML module
      $Results | Out-HtmlView -FilePath $ReportPath

    }
    Default {
      Write-Host "Invalid ReportType selection. Defaulting to csv."
      $Results | Export-Csv $ReportPath
      $openFile = Read-Host "Do you want to open the file? (Y/N)"
      if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        Invoke-Item $ReportPath
      }
    }
  }
  
  Write-Host "Report saved at $ReportPath"
}

