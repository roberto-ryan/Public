function New-vtsBulkADUsers {
  <#
  .SYNOPSIS
      Creates multiple Active Directory users from a CSV file.
  
  .DESCRIPTION
      The New-vtsBulkADUsers function automates the creation of multiple Active Directory users
      using data from a CSV file. It supports creating users in either a new or existing OU,
      sets up email properties, and triggers an AD sync after completion.
  
  .PARAMETER CsvPath
      Path to the CSV file containing user information. Required columns:
      - FirstName
      - LastName
      - PrimaryEmail
  
  .PARAMETER OUOption
      Specifies whether to create a new OU ('CreateNew') or select an existing one ('Select').
  
  .PARAMETER NewOUName
      Name of the new OU to create. Required when OUOption is 'CreateNew'.
  
  .PARAMETER NewOUPath
      Distinguished path where the new OU will be created. Required when OUOption is 'CreateNew'.
  
  .PARAMETER Password
      SecureString containing the default password for new users.
      Defaults to "WelcomeRocky123!".
  
  .PARAMETER LogPath
      Path where the operation log will be saved.
      Defaults to "C:\temp\User_Creation-<current-date>.csv".
  
  .EXAMPLE
      PS> New-vtsBulkADUsers -CsvPath "C:\Users.csv" -OUOption Select
      Creates users from Users.csv in an existing OU selected via GUI.
  
  .EXAMPLE
      PS> $params = @{
          CsvPath = "C:\Users.csv"
          OUOption = "CreateNew"
          NewOUName = "NewEmployees"
          NewOUPath = "DC=contoso,DC=com"
      }
      PS> New-vtsBulkADUsers @params
      Creates users from Users.csv in a new OU named "NewEmployees".
  
  .NOTES
      Author: VTS Systems
      Required Modules: ActiveDirectory
      CSV Format Example:
      FirstName,LastName,PrimaryEmail
      John,Doe,john.doe@contoso.com
  
  .LINK
      M365
  #>
  [CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string]$CsvPath,

  [Parameter(Mandatory)]
  [ValidateSet('CreateNew', 'Select')]
  [string]$OUOption,

  [Parameter()]
  [string]$NewOUName,

  [Parameter()]
  [string]$NewOUPath,

  [Parameter()]
  [SecureString]$Password = ("WelcomeRocky123!" | ConvertTo-SecureString -AsPlainText -Force),

  [Parameter()]
  [string]$LogPath = "C:\temp\User_Creation-$(get-date -Format 'yyyy-MM-dd').csv"
)

  function Write-Log {
    [CmdletBinding()]
    param(
      [Parameter()]
      [ValidateNotNullOrEmpty()]
      [string]$Message,

      [Parameter()]
      [ValidateNotNullOrEmpty()]
      [ValidateSet('Information', 'Warning', 'Error')]
      [string]$Severity = 'Information'
    )

    [pscustomobject]@{
      Time     = (Get-Date -f g)
      Message  = $Message
      Severity = $Severity
    } | Export-Csv -Path $LogPath -Append -NoTypeInformation

    Write-Host -Object $Message
  }

  # Create log directory if it doesn't exist
  if (-not (Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType File -Force
  }

  # Handle OU Selection
  if ($OUOption -eq 'CreateNew') {
    if (-not $NewOUName -or -not $NewOUPath) {
      throw "NewOUName and NewOUPath are required when creating a new OU"
    }
    try {
      $SelectedOU = New-ADOrganizationalUnit -Name $NewOUName -Path $NewOUPath -PassThru
      Write-Log -Message "Created new OU: $($SelectedOU.DistinguishedName)" -Severity Information
    }
    catch {
      throw "Failed to create new OU: $_"
    }
  }
  else {
    $SelectedOU = Get-ADOrganizationalUnit -Filter * | 
    Out-GridView -Title "Select Target OU" -OutputMode Single
    if (-not $SelectedOU) {
      throw "No OU selected"
    }
  }

  # Import user list
  $UserList = Import-Csv $CsvPath

  # Create users
  foreach ($user in $UserList) {
    $splat = @{
      Path              = $SelectedOU.DistinguishedName
      Name              = "$($user.FirstName) $($user.LastName)"
      UserPrincipalName = $user.PrimaryEmail
      SamAccountName    = "$($user.FirstName).$($user.LastName)"
      GivenName         = $user.FirstName
      Surname           = $user.LastName
      AccountPassword   = $Password
      Enabled           = $true 
    }

    try {
      New-ADUser @splat -Verbose -Confirm:$false
      Set-ADUser "$($user.FirstName).$($user.LastName)" -Add @{ProxyAddresses = "SMTP:$($user.PrimaryEmail)" }
      Set-ADUser "$($user.FirstName).$($user.LastName)" -EmailAddress $user.PrimaryEmail
      Write-Log -Message "Created user: $($user.FirstName) $($user.LastName)" -Severity Information
    }
    catch {
      Write-Log -Message "Failed to create user $($user.FirstName) $($user.LastName): $_" -Severity Error
    }
  }

  # Run delta sync
  try {
    Start-ADSyncSyncCycle -PolicyType Delta
    while (-not $?) { 
      Start-Sleep 5
      Start-ADSyncSyncCycle -PolicyType Delta 
    }
    Write-Log -Message "AD Sync completed successfully" -Severity Information
  }
  catch {
    Write-Log -Message "AD Sync failed: $_" -Severity Error
  }
  
}

