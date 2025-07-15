function Import-vts365DistributionListConfig {
  <#
  .SYNOPSIS
  Imports and creates distribution lists from a CSV configuration file.
  
  .DESCRIPTION
  The Import-vts365DistributionListConfig function creates new distribution lists in Exchange Online based on configuration data from a CSV file. It supports creating security groups, distribution groups, and room lists with their respective settings and members.
  
  The function will prefix all new distribution list names with "C-" and will automatically handle different types of groups including:
  - MailUniversalSecurityGroup
  - MailUniversalDistributionGroup
  - RoomList
  
  .PARAMETER FilePath
  The path to the CSV file containing the distribution list configurations to import. The CSV file should include columns for DisplayName, Name, Alias, PrimarySmtpAddress, and other distribution list properties.
  
  .EXAMPLE
  Import-vts365DistributionListConfig -FilePath "C:\DistributionLists.csv"
  
  This example imports distribution list configurations from the specified CSV file and creates the corresponding groups in Exchange Online.
  
  .NOTES
  Requires an active connection to Exchange Online PowerShell.
  Make sure the CSV file contains all required columns and properly formatted data.
  
  .LINK
  M365 Migration Scripts
  #>
  [CmdletBinding()]
  param (
      [Parameter(Mandatory)]
      [String]$FilePath,
      [Parameter()]
      [String]$AppendToDisplayName,
      [Parameter()]
      [String]$DefaultDomain
  )

  if (-not($AppendToDisplayName)) {
      $AppendToDisplayName = Read-Host "Enter Company Abbreviation to Append to DisplayName (so we can find the groups later.)"
  }

  # Import CSV and get default domain
  $GroupsData = Import-Csv $FilePath
  if(-not($DefaultDomain)){
      $DefaultDomain = Get-AcceptedDomain | Where-Object Default -eq $True | Select-Object -ExpandProperty DomainName
  }

  foreach ($GroupData in $GroupsData) {
      $LocalPart = ($GroupData.PrimarySmtpAddress) -replace '@.*$'
      $DisplayName = $GroupData.DisplayName + " - $AppendToDisplayName"
      
      # Check if group exists
      $ExistingGroup = Get-DistributionGroup -Identity $DisplayName -ErrorAction SilentlyContinue
      if ($ExistingGroup) {
          Write-Host "Group exists: $DisplayName" -ForegroundColor Yellow
          continue
      }

      try {
          Write-Host "Creating group: $DisplayName" -ForegroundColor Cyan
          
          # Create group first without owners
          $NewGroupParams = @{
              DisplayName = $DisplayName
              Name = $GroupData.Name
              Alias = $GroupData.Alias
              PrimarySMTPAddress = "$LocalPart@$DefaultDomain"
          }

          # Add type-specific parameters
          switch ($GroupData.RecipientTypeDetails) {
              "MailUniversalSecurityGroup" { 
                  $NewGroupParams.Type = "Security" 
              }
              "RoomList" { 
                  $NewGroupParams.Roomlist = $true 
              }
          }

          # Create the group
          $NewGroup = New-DistributionGroup @NewGroupParams

          # Set basic properties
          $SetGroupParams = @{
              Identity = $NewGroup.DisplayName
              # HiddenFromAddressListsEnabled = $True
              MemberJoinRestriction = $GroupData.MemberJoinRestriction
              MemberDepartRestriction = $GroupData.MemberDepartRestriction
              RequireSenderAuthenticationEnabled = [System.Convert]::ToBoolean($GroupData.RequireSenderAuthenticationEnabled)
          }

          if (-not [string]::IsNullOrWhiteSpace($GroupData.Notes)) {
              $SetGroupParams.Description = $GroupData.Notes
          }

          Set-DistributionGroup @SetGroupParams

          # Add owners separately
          if (-not [string]::IsNullOrWhiteSpace($GroupData.ManagedBy)) {
              $Owners = $GroupData.ManagedBy -split ',' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
              Write-Host "Adding owners..." -ForegroundColor Cyan
              
              foreach ($Owner in $Owners) {
                  try {
                      Set-DistributionGroup -Identity $NewGroup.DisplayName -ManagedBy @{Add=$Owner} -ErrorAction Stop
                      Write-Host "Added owner: $Owner" -ForegroundColor Green
                  }
                  catch {
                      Write-Host "Failed to add owner $Owner : $_" -ForegroundColor Yellow
                  }
              }
          }

          Start-Sleep -Seconds 5

          # Add members
          if (-not [string]::IsNullOrWhiteSpace($GroupData.MembersPrimarySmtpAddress)) {
              $Members = $GroupData.MembersPrimarySmtpAddress -split ',' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
              foreach ($Member in $Members) {
                  try {
                      Add-DistributionGroupMember -Identity $NewGroup.PrimarySmtpAddress -Member $Member -BypassSecurityGroupManagerCheck -ErrorAction Stop
                      Write-Host "Added member: $Member" -ForegroundColor Green
                  }
                  catch {
                      Write-Host "Failed to add member $Member : $_" -ForegroundColor Red
                  }
              }
          }

          Write-Host "Distribution group $DisplayName created successfully." -ForegroundColor Green
      }
      catch {
          Write-Host "Failed to create group $DisplayName : $_" -ForegroundColor Red
      }
  }
}

