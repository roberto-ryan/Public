function Start-vtsRepair {
  <#
  .SYNOPSIS
  This function invokes a system check using DISM and System File Checker.
  
  .DESCRIPTION
  The Start-vtsRepair function initiates a system check by first running the DISM restore health process, followed by a System File Checker scan. It provides status updates at each stage of the process.
  
  .EXAMPLE
  Start-vtsRepair
  This command will initiate the system check process.
  
  .NOTES
  The DISM restore health process can help fix Windows corruption errors. The System File Checker scan will scan all protected system files, and replace corrupted files with a cached copy.
  
  .LINK
  Utilities
  #>
  Write-Host "Starting DISM restore health process..."
  dism /online /cleanup-image /restorehealth
  Write-Host "DISM restore health process completed."

  Write-Host "Starting System File Checker scan now..."
  sfc /scannow
  Write-Host "System File Checker scan completed."
}

