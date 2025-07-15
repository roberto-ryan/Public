function Show-vtsToastNotification {
  <#
  .DESCRIPTION
  Returns users toast notifications. Duplicates notifications are removed for brevity.
  .EXAMPLE
  Show notifications for all users
  PS> Show-vtsToastNotification
  .EXAMPLE
  Show notifications for a selected user
  PS> Show-vtsToastNotification -user john.doe
  
  .LINK
  Device Management
  #>
  param(
    $user = (Get-ChildItem C:\Users\ | Select-Object -ExpandProperty Name)
  )

  $db = foreach ($u in $user) {
    Get-Content "C:\Users\$u\AppData\Local\Microsoft\Windows\Notifications\wpndatabase.db-wal" 2>$null
  }

  $tags = @(
    'text>'
    'text id="1">'
    'text id="2">'
  )

  $notification = foreach ($tag in $tags) {
    ($db -split '<' |
    Select-String $tag |
    Select-Object -ExpandProperty Line) -replace "$tag", "" -replace "</text>", "" |
    Select-String -NotMatch '/'
  }

  Write-Host "Duplicates removed for brevity." -ForegroundColor Yellow
  $notification | Select-Object -Unique
}

