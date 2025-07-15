function New-vts365User {
  <#
  .SYNOPSIS
  Creates new Microsoft 365 users from a CSV file.
  
  .DESCRIPTION
  The New-vts365User function creates new Microsoft 365 users with the specified domain and usage location. It requires a CSV file path as input, which should contain the user details. The function also allows setting a password length.
  
  The CSV file must contain the following columns:
  "FirstName","LastName","PrimaryEmail","RecoveryPhone","RecoveryEmail"
  
  Example CSV content:
  "FirstName","LastName","PrimaryEmail","RecoveryPhone","RecoveryEmail"
  "John","Doe","john.doe@contoso.com","+1 234567890","john.recovery@email.com"
  "Jane","Smith","jane.smith@contoso.com","+1 987654321","jane.recovery@email.com"
  
  .PARAMETER CsvPath
  The path to the CSV file containing user information to be imported.
  
  .PARAMETER UsageLocation
  The usage location for the new user. Default is 'US'.
  
  .PARAMETER PasswordLength
  The length of the password to be generated for the new user. Default is 16 characters.
  
  .PARAMETER CustomPassword
  Specifies a custom password for the new user. If not provided, a random password will be generated.
  
  .PARAMETER CreateTemplate
  Switch to create a CSV template for user input.
  
  .EXAMPLE
  New-vts365User -CsvPath "C:\Users\example\userlist.csv" -Domain "contoso.com"
  
  This command creates new users from the specified CSV file with the domain 'contoso.com' and the default usage location and password length.
  
  .EXAMPLE
  New-vts365User -CsvPath "C:\Users\example\userlist.csv" -Domain "contoso.com" -UsageLocation "GB" -PasswordLength 20
  
  This command creates new users with the domain 'contoso.com', sets the usage location to 'GB', and generates passwords with a length of 20 characters.
  
  .NOTES
  Requires the MSOnline and AzureAD modules.
  
  .LINK
  M365
  
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({
      if (Test-Path $_ -PathType Leaf) {
        $true
      }
      else {
        throw "File not found or not a valid file: $_"
      }
    })]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("US", "CA", "MX", "GB", "DE", "FR", "JP", "AU", "BR", "IN")]
    [string]$UsageLocation = "US",

    [Parameter(Mandatory = $false)]
    [ValidateRange(12, 32)]
    [int]$PasswordLength = 16,

    [Parameter(Mandatory = $false)]
    [string]$CustomPassword,

    [switch]$CreateTemplate,

    [switch]$UseLast4PhoneDigitsAsPassword,

    $keyPath = "$($env:USERPROFILE)\Documents\$(Get-Date -f MMddyymmss)_temp.txt"
  )

  if($UseLast4PhoneDigitsAsPassword){
    $codeword = Read-Host "Enter a code word to use for generating passwords"
  }

  # Function to retry operations
  function Invoke-WithRetry {
    param(
      [ScriptBlock]$Action,
      [int]$MaxAttempts = 3,
      [int]$DelaySeconds = 2
    )

    $attempts = 0
    do {
      $attempts++
      try {
        return & $Action
      }
      catch {
        if ($attempts -eq $MaxAttempts) { throw }
        Start-Sleep -Seconds $DelaySeconds
      }
    } while ($attempts -lt $MaxAttempts)
  }

  # Function to create a random password
  function Get-RandomPassword {
    param ([int]$Length = 16)
    $charSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_-+=[]{}|;:,.<>?'.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($Length)
    $rng.GetBytes($bytes)

    $result = New-Object char[]($Length)
    for ($i = 0 ; $i -lt $Length ; $i++) {
      $result[$i] = $charSet[$bytes[$i] % $charSet.Length]
    }
    return (-join $result)
  }

  # Check and install required modules
  function Install-RequiredModules {
    $requiredModules = @("MSOnline", "Microsoft.Graph")
    foreach ($module in $requiredModules) {
      if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Verbose "Installing $module module..."
        Install-Module -Name $module -Force -Scope CurrentUser
      }
    }
  }

  # Connect to Microsoft 365 and Microsoft Graph
  function Connect-Services {
    try {
      if ((Get-MsolCompanyInformation).DisplayName) {
        Write-Verbose "Already connected to $((Get-MsolCompanyInformation).DisplayName)'s Microsoft 365 tenant."
      }
      else {
        Connect-MsolService
      }
      Connect-MgGraph -Scopes "User.ReadWrite.All", "UserAuthenticationMethod.ReadWrite.All", "Directory.AccessAsUser.All", "Directory.ReadWrite.All" -UseDeviceAuthentication -ForceRefresh
    }
    catch {
      Write-Error "Failed to connect to Microsoft 365 or Microsoft Graph. Error: $_"
      throw
    }
  }

  # Create CSV Template if requested
  function Create-CSVTemplate {
    $templatePath = Join-Path $env:USERPROFILE "Documents\$(Get-Date -f MMddyymmss)_M365NewUser.csv"
    $templateData = @(
      [PSCustomObject]@{
        FirstName     = ""
        LastName      = ""
        PrimaryEmail  = ""
        RecoveryPhone = ""
        RecoveryEmail = ""
      }
    )

    try {
      $templateData | Export-Csv -Path $templatePath -NoTypeInformation
      while(-not(Test-Path $templatePath)) {
        Start-Sleep -Seconds 1
      }
      Invoke-Item $templatePath
      Write-Host "The csv template is being created, and will be automatically opened momentarily. Enter the user details in the new csv template. Once finished, save the template, then type 'next' to proceed."
      do {
        $response = Read-Host
      } while ($response -ne "next")
      return $templatePath
    }
    catch {
      Write-Error "Failed to create template CSV file: $_"
      throw
    }
  }

  # Import and validate CSV data
  function Import-CSVData {
    param(
      [string]$Path
    )
    try {
      $users = Import-Csv $Path
      $requiredColumns = @("FirstName", "LastName", "PrimaryEmail", "RecoveryPhone", "RecoveryEmail")
      $missingColumns = $requiredColumns | Where-Object { $_ -notin $users[0].PSObject.Properties.Name }
      if ($missingColumns) {
        throw "CSV is missing required columns: $($missingColumns -join ', ')"
      }
      return $users
    }
    catch {
      Write-Error "Failed to parse CSV data: $_"
      throw
    }
  }

  # Process each user
  function Process-Users {
    param(
      [array]$Users
    )

    foreach ($user in $Users) {
      if ($user.LastName) {
        $displayName = "$($user.FirstName) $($user.LastName)"
      }
      else {
        $displayName = $user.FirstName
      }

      $userPrincipalName = $user.PrimaryEmail

      if ($CustomPassword) {
        $password = $CustomPassword
        "$($user.PrimaryEmail) $password" | out-file -filepath $keypath -Append -Force
      } elseif ($UseLast4PhoneDigitsAsPassword) {
        do {
          $last4 = ($user.RecoveryPhone -replace "[^0-9]", "")[-4..-1] -join ""
          if (-not($last4)){
            $last4 = 1234
          }
          $password = $codeword+$last4

        } until (
          $password.Length -ge 12
        )
        "$($user.PrimaryEmail) $password" | out-file -filepath $keypath -Append -Force
      }
      else {
        $password = Get-RandomPassword -Length $PasswordLength
        "$($user.PrimaryEmail) $password" | out-file -filepath $keypath -Append -Force
      }
      $PasswordProfile = @{
        Password                      = $password
        ForceChangePasswordNextSignIn = $true
      }

      try {
        # Check if user exists in Microsoft Graph
        $existingUser = Get-MgUser -Filter "userPrincipalName eq '$($user.PrimaryEmail)'" -ErrorAction SilentlyContinue

        if ($existingUser) {
          Write-Verbose "User $($user.PrimaryEmail) already exists. Updating details..."

          if ($CustomPassword) {
            if ($user.LastName) {
              Update-MgUser -UserId $existingUser.Id -DisplayName "$displayName" -GivenName $user.FirstName -Surname $user.LastName -UsageLocation $UsageLocation -PasswordProfile $PasswordProfile
            }
            else {
              Update-MgUser -UserId $existingUser.Id -DisplayName "$displayName" -GivenName $user.FirstName -UsageLocation $UsageLocation -PasswordProfile $PasswordProfile
            }
          }
          else {
            if ($user.LastName) {
              Update-MgUser -UserId $existingUser.Id -DisplayName "$displayName" -GivenName $user.FirstName -Surname $user.LastName -UsageLocation $UsageLocation
            }
            else {
              Update-MgUser -UserId $existingUser.Id -DisplayName "$displayName" -GivenName $user.FirstName -UsageLocation $UsageLocation
            }
          }

          # Update aliases (still using MSOnline as Graph doesn't have a direct equivalent)
          $currentAliases = Get-MsolUser -UserPrincipalName $user.PrimaryEmail | Select-Object -ExpandProperty ProxyAddresses
          if ($currentAliases -notcontains "smtp:$userPrincipalName") {
            # Set-MsolUser -UserPrincipalName $user.PrimaryEmail -EmailAddresses @($currentAliases + "smtp:$userPrincipalName")
          }

          Write-Output "User updated: $displayName ($($user.PrimaryEmail))"
        }
        else {
          Write-Verbose "Creating new user: $displayName"

          try {
            if ($user.LastName) {
              # Create new user in Microsoft Graph
              $newUser = New-MgUser -DisplayName "$displayName" `
                -GivenName $user.FirstName `
                -Surname $user.LastName `
                -UserPrincipalName $user.PrimaryEmail `
                -UsageLocation $UsageLocation `
                -PasswordProfile $PasswordProfile `
                -AccountEnabled:$true `
                -MailNickname ($user.FirstName.ToLower())
            }
            else {
              # Create new user in Microsoft Graph
              $newUser = New-MgUser -DisplayName "$displayName" `
                -GivenName $user.FirstName `
                -UserPrincipalName $user.PrimaryEmail `
                -UsageLocation $UsageLocation `
                -PasswordProfile $PasswordProfile `
                -AccountEnabled:$true `
                -MailNickname ($user.FirstName.ToLower())
            }

            Write-Output "User created: $displayName ($($user.PrimaryEmail))"
          }
          catch {
            Write-Error "Failed to create or update user $displayName. Error: $_"
          }
        }
      }
      catch {
        Write-Error "Failed to create or update user $displayName. Error: $_"
      }
    }

    Write-Output "All users have been processed. Waiting for user sync to complete..."

    do {
      Start-Sleep -Seconds 30
      Get-MgUser -Filter "userPrincipalName eq '$($user.PrimaryEmail)'" -ErrorAction SilentlyContinue
    } until (
      $?
    )

    foreach ($user in $Users) {
      $newlyCreatedUser = Get-MgUser -Filter "userPrincipalName eq '$($user.PrimaryEmail)'"
      if ($user.RecoveryPhone) {
        Invoke-WithRetry -Action {
          # Set phone authentication method
          $phoneParams = @{
            PhoneNumber = "$($user.RecoveryPhone)"
            PhoneType   = "mobile"
          }
          New-MgUserAuthenticationPhoneMethod -UserId $newlyCreatedUser.Id -BodyParameter $phoneParams
        }
      }

      if ($user.RecoveryEmail) {
        Invoke-WithRetry -Action {
          # Set email authentication method
          $emailParams = @{
            EmailAddress = $user.RecoveryEmail
          }
          New-MgUserAuthenticationEmailMethod -UserId $newlyCreatedUser.Id -BodyParameter $emailParams
        }
      }
    }
  }

  # Disconnect from services
  function Disconnect-Services {
    # Disconnect-MgGraph
    Write-Verbose "Disconnected from Microsoft 365 and Microsoft Graph services."
  }

  # Main logic starts here
  if (-not($CsvPath) -and -not($CreateTemplate)) {
    Write-Error "Please provide a CSV file path or use the -CreateTemplate switch to generate a template CSV file."
    return
  }

  if ($CreateTemplate) {
    $CsvPath = Create-CSVTemplate
  }

  Install-RequiredModules
  Connect-Services
  $users = Import-CSVData -Path $CsvPath
  Process-Users -Users $users
  Disconnect-Services
  Write-Verbose "All users have been processed."
  ii $keyPath
}

