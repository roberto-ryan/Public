function Export-vts365DistributionListConfig {
  <#
  .SYNOPSIS
  Exports the configuration of selected distribution lists to a CSV file.
  
  .DESCRIPTION
  The Export-vts365DistributionListConfig function connects to Exchange Online and exports detailed configuration information about selected distribution lists to a CSV file. It includes information such as managed by users, members, moderators, and various distribution list settings.
  
  .PARAMETER FilePath
  The path where the CSV file will be saved. If not specified, defaults to the Downloads folder with filename format 'yyyy-MM-dd_HHmm_DistributionListConfig.csv'.
  
  .EXAMPLE
  Export-vts365DistributionListConfig
  
  This example prompts for distribution list selection and exports their configuration to the default location.
  
  .EXAMPLE
  Export-vts365DistributionListConfig -FilePath "C:\Reports\DLConfig.csv"
  
  This example exports the selected distribution lists' configuration to the specified file path.
  
  .NOTES
  Requires connection to Exchange Online PowerShell.
  
  .LINK
  M365 Migration Scripts
  #>
  [CmdletBinding()]
  param (
      [Parameter()]
      [String]
      $FilePath = (Join-Path -Path $env:USERPROFILE -ChildPath "Downloads\$(Get-Date -f 'yyyy-MM-dd_HHmm')_DistributionListConfig.csv")
  )
  
  # Get all distribution groups
  $Groups = Get-DistributionGroup -ResultSize Unlimited |
  Out-GridView -OutputMode Multiple -Title "Select Distribution Groups to Export"
  
  # Count the total number of distribution groups
  $totalGroups = $Groups.Count
  $currentGroupIndex = 0
  
  # Initialize a List to store the data
  $Report = [System.Collections.Generic.List[Object]]::new()
  
  # Loop through distribution groups
  foreach ($Group in $Groups) {
      $currentGroupIndex++
      $GroupDN = $Group.DistinguishedName
  
      # Get ManagedBy names and SMTP addresses properly
      $ManagedByNames = @()
      $ManagedBySmtp = @()
      if ($Group.ManagedBy) {
          $Group.ManagedBy | ForEach-Object {
              try {
                  $owner = Get-Recipient $_ -ErrorAction Stop
                  $ManagedByNames += $owner.DisplayName
                  $ManagedBySmtp += $owner.PrimarySmtpAddress
              }
              catch {
                  $ManagedByNames += $_
                  $ManagedBySmtp += $_
              }
          }
      }

      # Update progress bar
      $progressParams = @{
          Activity        = "Processing Distribution Groups"
          Status          = "Processing Group $currentGroupIndex of $totalGroups"
          PercentComplete = ($currentGroupIndex / $totalGroups) * 100
      }
  
      Write-Progress @progressParams
  
      $GroupMembers = Get-DistributionGroupMember $GroupDN -ResultSize Unlimited
      
      # Get required attributes directly within the output object
      $ReportLine = [PSCustomObject]@{
          DisplayName                            = $Group.DisplayName
          Name                                   = $Group.Name
          PrimarySmtpAddress                     = $Group.PrimarySmtpAddress
          EmailAddresses                         = ($Group.EmailAddresses -join ',')
          Domain                                 = $Group.PrimarySmtpAddress.ToString().Split("@")[1]
          Alias                                  = $Group.Alias
          GroupType                              = $Group.GroupType
          RecipientTypeDetails                   = $Group.RecipientTypeDetails
          Members                                = $GroupMembers.Name -join ','
          MembersPrimarySmtpAddress              = $GroupMembers.PrimarySmtpAddress -join ','
          ManagedBy                              = ($ManagedByNames -join ',')
          ManagedBySmtpAddress                   = ($ManagedBySmtp -join ',')
          HiddenFromAddressLists                 = $Group.HiddenFromAddressListsEnabled
          MemberJoinRestriction                  = $Group.MemberJoinRestriction
          MemberDepartRestriction                = $Group.MemberDepartRestriction
          AcceptMessagesOnlyFrom                 = ($Group.AcceptMessagesOnlyFrom.Name -join ',')
          AcceptMessagesOnlyFromDLMembers        = ($Group.AcceptMessagesOnlyFromDLMembers -join ',')
          AcceptMessagesOnlyFromSendersOrMembers = ($Group.AcceptMessagesOnlyFromSendersOrMembers -join ',')
          ModeratedBy                            = ($Group.ModeratedBy -join ',')
          BypassModerationFromSendersOrMembers   = ($Group.BypassModerationFromSendersOrMembers -join ',')
          ModerationEnabled                      = $Group.ModerationEnabled
          SendModerationNotifications            = $Group.SendModerationNotifications
          GrantSendOnBehalfTo                    = ($Group.GrantSendOnBehalfTo.Name -join ',')
      }
      $Report.Add($ReportLine)
  }
  
  # Clear progress bar
  Write-Progress -Activity "Processing Distribution Groups" -Completed
  
  # Sort the output by DisplayName and export to CSV file
  $Report | Sort-Object DisplayName | Export-Csv -Path $FilePath -NoTypeInformation

  $Report | Out-Host

  Write-Host "`n`nCSV file has been created at: $FilePath `n`n"

  # Ask the user if they want to open the file after it's created
  $OpenFile = Read-Host "Do you want to open the CSV file? (Y/N)"
  if ($OpenFile -eq 'Y') {
      # Open the CSV file in the default application
      Invoke-Item $FilePath
  }
}

