function Get-vts365PowerBIUsageReport {
  <#
  .SYNOPSIS
      Performs various operations on a folder structure based on specified configuration settings.

  .DESCRIPTION
      This script processes folders according to the configuration specified in a JSON file.
      It can:
      - Move folders to appropriate destinations based on their names
      - Rename folders according to defined patterns
      - Archive folders that match specified criteria
      - Handle nested folder structures
      
      The script reads its configuration from a JSON file that defines rules for processing
      folders, including pattern matching, destination paths, and archiving settings.

  .PARAMETER ConfigFile
      Path to the JSON configuration file that defines processing rules.

  .PARAMETER RootFolder
      Path to the root folder where processing should begin.

  .PARAMETER ArchiveFolder
      Path to the folder where archived items should be stored.

  .PARAMETER LogFile
      Path to the log file where processing activities will be recorded.

  .PARAMETER Force
      If specified, forces the script to execute without confirmation prompts.

  .EXAMPLE
      .\YourScript.ps1 -ConfigFile "config.json" -RootFolder "C:\Data" -ArchiveFolder "C:\Archive" -LogFile "C:\Logs\process.log"
      
      Processes folders in C:\Data according to rules in config.json, archives to C:\Archive, and logs to the specified file.

  .EXAMPLE
      .\YourScript.ps1 -ConfigFile "config.json" -RootFolder "C:\Data" -Force
      
      Processes folders in C:\Data using the default archive location and logging, without any confirmation prompts.

  .NOTES
      File Name      : YourScript.ps1
      Author         : Your Name
      Prerequisite   : PowerShell 5.1 or later
      Copyright      : Your Copyright Information
      
  .LINK
      M365
  #>

  # Check if the required module is installed, if not, install it
  if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
      Write-Host "MicrosoftPowerBIMgmt module not found. Installing..."
      Install-Module -Name MicrosoftPowerBIMgmt -Force -AllowClobber
  } else {
      Write-Host "MicrosoftPowerBIMgmt module is already installed."
  }

  Connect-PowerBIServiceAccount

  # Define the number of days to retrieve (max 30)
  $daysToExtract = 30

  # Initialize an array to store all activities
  $allActivities = @()

  # Loop through each day
  for ($i = 0; $i -lt $daysToExtract; $i++) {
      # Calculate start and end times for the current day
      $startDate = (Get-Date).AddDays(-$i).ToString("yyyy-MM-ddT00:00:00Z")
      $endDate = (Get-Date).AddDays(-$i).ToString("yyyy-MM-ddT23:59:59Z")
  
      Write-Host "Fetching activities for $startDate to $endDate"

      # Fetch activities for this day
      $activities = Get-PowerBIActivityEvent -StartDateTime $startDate -EndDateTime $endDate -Verbose

      # Parse JSON response (if returned as JSON)
      if ($activities) {
          $activityData = $activities | ConvertFrom-Json
          if ($activityData) {
              $allActivities += $activityData
          }
      }
  }

  # Export to CSV if there’s data
  if ($allActivities.Count -gt 0) {
      $allActivities | Select-Object -Property CreationTime, UserId, Activity, Operation | Export-Csv -Path "C:\Reports\PowerBIActivity_30Days.csv" -NoTypeInformation
      Write-Host "Data exported to C:\Reports\PowerBIActivity_30Days.csv"
  }
  else {
      Write-Host "No activity data found for the specified 30-day period."
  }
}

