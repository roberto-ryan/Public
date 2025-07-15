function Stop-vtsScreenRecording {
  <#
  .SYNOPSIS
      The Stop-vtsScreenRecording function stops the ongoing screen recording.
  
  .DESCRIPTION
      The Stop-vtsScreenRecording function stops the screen recording initiated by the Start-vtsScreenRecording function. It disables the scheduled task that was created to start the recording and stops the FFmpeg and PowerShell processes that were running the recording.
  
  .PARAMETER No Parameters
  
  .EXAMPLE
      PS C:\> Stop-vtsScreenRecording
  
      This command stops the ongoing screen recording.
  
  .NOTES
      This function should be used to stop a screen recording that was started with the Start-vtsScreenRecording function.
  
  .LINK
      Utilities
  #>
  Disable-ScheduledTask -TaskName "RecordSession"
  Get-Process ffmpeg, powershell | Stop-Process -Force -Confirm:$false
}

