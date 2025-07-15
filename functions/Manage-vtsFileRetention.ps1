function Manage-vtsFileRetention {
  <#
  .SYNOPSIS
      Manages file retention by resetting, marking, moving, and deleting files based on access times.
  
  .DESCRIPTION
      This script manages file retention in a specified directory. It performs the following actions:
      - Resets the LastWriteTime for recently accessed files.
      - Marks old files by setting an artificial LastWriteTime.
      - Moves old files to a "ToBeDeleted" directory.
      - Deletes files from the "ToBeDeleted" directory after a specified grace period.
  
  .PARAMETER BasePath
      The base path where the files are located. Default is "C:\qdstm*\out".
  
  .PARAMETER DaysBeforeSoftDelete
      The number of days before a file is considered old and marked for deletion. Default is 30 days.
  
  .PARAMETER DaysBeforeHardDelete
      The number of days before a file is permanently deleted from the "ToBeDeleted" directory. Default is 45 days.
  
  .PARAMETER ArtificialLastWriteTimeYears
      The number of years to subtract from the LastWriteTime to mark a file as old. Default is 100 years.
  
  .PARAMETER Extension
      The file extension to filter files. Default is "*.pdf".
  
  .EXAMPLE
      PS> .\Manage-vtsFileRetention.ps1 -BasePath "C:\example\path" -DaysBeforeSoftDelete 60 -DaysBeforeHardDelete 90 -Extension "*.txt"
      This example manages file retention for .txt files in the specified path, marking files as old after 60 days and deleting them after 90 days.
  
  .LINK
      File Management
  
  #>
  param (
      [string]$BasePath = "C:\qdstm*\out",
      [int]$DaysBeforeSoftDelete = 30,
      [int]$DaysBeforeHardDelete = 45,
      [int]$ArtificialLastWriteTimeYears = 100,
      [string]$Extension = ".pdf"
  )

  $OUT = Get-Item -Path $BasePath | Select-Object -ExpandProperty FullName

  $ToBeDeleted = Join-Path $OUT "ToBeDeleted"

  function Write-ActionLog {
      param(
          [string]$Message,
          [string]$Action,
    [string]$ToBeDeleted = $ToBeDeleted
      )
      $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
      $logMessage = "[$timestamp] $Action : $Message"
      $logPath = Join-Path $ToBeDeleted "file_management.log"
      Add-Content -Path $logPath -Value $logMessage
      Write-Host $logMessage
  }

  if ($OUT) {
      $ToBeDeleted = Join-Path $OUT "ToBeDeleted"
  
      if (-not (Test-Path $ToBeDeleted)) {
          try {
              New-Item -ItemType Directory -Path $ToBeDeleted
              Write-ActionLog -Message "Created ToBeDeleted directory at $ToBeDeleted" -Action "INIT"
          }
          catch {
              Write-ActionLog -Message "Failed to create directory: $ToBeDeleted. Error: $_" -Action "ERROR"
              exit
          }
      }
  
      # Reset LastWriteTime for recently accessed files
      try {
          Get-ChildItem -Path $OUT | 
          Where-Object { 
      $_.Extension -eq $Extension -and
      $_.LastAccessTime -gt (Get-Date).AddDays(-1) 
    } | 
          ForEach-Object { 
              $_.LastWriteTime = Get-Date
              Write-ActionLog -Message "Reset LastWriteTime for recently accessed file: $($_.FullName)" -Action "RESET"
          }
      }
      catch {
          Write-ActionLog -Message "Failed to update LastWriteTime for accessed files in $OUT. Error: $_" -Action "ERROR"
          exit
      }
  
      # Process old files
      try {
          Get-ChildItem -Path $OUT | 
          Where-Object {
      $_.Extension -eq $Extension -and
      $_.LastAccessTime -lt (Get-Date).AddDays( - ($DaysBeforeSoftDelete))
    } | 
          ForEach-Object { 
              $_.LastWriteTime = ($_.LastWriteTime).AddYears( - ($ArtificialLastWriteTimeYears))
              Write-ActionLog -Message "Marked file as old: $($_.FullName)" -Action "MARK"
          }
      }
      catch {
          Write-ActionLog -Message "Failed to update LastWriteTime for old files in $OUT. Error: $_" -Action "ERROR"
          exit
      }
  
      # Move old files block - modified to set new CreationTime
      try {
          Get-ChildItem -Path $OUT | 
          Where-Object { 
      $_.Extension -eq $Extension -and
      $_.LastAccessTime -lt (Get-Date).AddDays( - ($DaysBeforeSoftDelete)) 
    } | 
          ForEach-Object {
              $fileName = $_.FullName
              # Set CreationTime to now for the moved file
              $_.CreationTime = Get-Date
              Move-Item -Path $_.FullName -Destination $ToBeDeleted
              Write-ActionLog -Message "Moved to ToBeDeleted: $fileName (not accessed for $DaysBeforeSoftDelete days)" -Action "MOVE"
          }
      }
      catch {
          Write-ActionLog -Message "Failed to move files to $ToBeDeleted. Error: $_" -Action "ERROR"
          exit
      }
  
      # Delete expired files - modified block
      try {
          Get-ChildItem -Path $ToBeDeleted -Recurse | 
    Where-Object Extension -eq ".pdf" |
          Where-Object { 
              $movedToDeletedDate = $_.CreationTime
              $daysInToBeDeleted = ((Get-Date) - $movedToDeletedDate).Days
              $isOldFile = $_.LastWriteTime -lt (Get-Date).AddYears( - ($ArtificialLastWriteTimeYears))
              $daysInToBeDeleted -ge ($DaysBeforeHardDelete - $DaysBeforeSoftDelete) -and $isOldFile
          } | 
          ForEach-Object {
              $fileName = $_.FullName
              $daysInToBeDeleted = ((Get-Date) - $_.CreationTime).Days
              Remove-Item -Path $_.FullName -Force
              Write-ActionLog -Message "Permanently deleted: $fileName (was in ToBeDeleted for $daysInToBeDeleted days)" -Action "DELETE"
          }
      }
      catch {
          Write-Error "Failed to remove files from $ToBeDeleted. Error: $_"
          exit
      }
  }
}

