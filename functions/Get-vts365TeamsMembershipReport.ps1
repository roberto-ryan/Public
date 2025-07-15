function Get-vts365TeamsMembershipReport {
  <#
  .SYNOPSIS
      This script generates a report of Microsoft Teams' memberships.
  
  .DESCRIPTION
      The Get-vts365TeamsMembershipReport function generates a report of all Microsoft Teams' memberships. 
      It lists all the teams along with their owners, members, and guests. 
      If a team does not have any members or guests, it will be noted in the report. 
      The report is then copied to the clipboard.
  
  .PARAMETER None
      This function does not take any parameters.
  
  .EXAMPLE
      PS C:\> .\Get-vts365TeamsMembershipReport.ps1
      This command runs the script and generates the report.
  
  .NOTES
      You need to have the MicrosoftTeams module installed and be connected to Microsoft Teams for this script to work.
  
  .LINK
      M365
  #>
  if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Install-Module -Name MicrosoftTeams
  }
  import-module MicrosoftTeams
  Connect-MicrosoftTeams

  $teamsToAudit = get-team | Out-GridView -OutputMode Multiple -Title "Select one or more Teams, then click OK."

  $report = @()

  foreach ($team in $teamsToAudit) {
    $owners = $team | Get-TeamUser -Role Owner | Select-Object -ExpandProperty User
    $members = $team | Get-TeamUser -Role Member | Select-Object -ExpandProperty User
    $guests = $team | Get-TeamUser -Role Guest | Select-Object -ExpandProperty User
    $report += "====================`n"
    $report += "Team Name: $($team.Displayname)`n"
    $report += "`nOwners:`n"
    foreach ($owner in $owners) {
      $report += "`t• $owner`n"
    }
    $report += "`nMembers:`n"
    if ($members) {
      foreach ($member in $members) {
        $report += "`t• $member`n"
      }
    }
    else {
      $report += "`tNo Members`n"
    }
    $report += "`nGuests:`n"
    if ($guests) {
      foreach ($guest in $guests) {
        $report += "`t• $guest`n"
      }
    }
    else {
      $report += "`tNo Guests`n"
    }
  }

  $report | Set-Clipboard

  Write-Host "Results copied to clipboard." -f Yellow
}

