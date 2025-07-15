function Get-vtsMappedDrive {
  <#
  .Description
  Displays Mapped Drives information from the Windows Registry.
  .EXAMPLE
  PS> Get-vtsMappedDrive
  
  Output:
  Username            : VTS-ROBERTO\rober
  DriveLetter         : Y
  RemotePath          : https://live.sysinternals.com
  ConnectWithUsername : rober
  SID                 : S-1-5-21-376445358-2603134888-3166729622-1001
  
  .LINK
  Drive Management
  #>
  [CmdletBinding()]
  param ()

  # On most OSes, HKEY_USERS only contains users that are logged on.
  # There are ways to load the other profiles, but it can be problematic.
  $Drives = Get-ItemProperty "Registry::HKEY_USERS\*\Network\*" 2>$null

  # See if any drives were found
  if ( $Drives ) {

    ForEach ( $Drive in $Drives ) {

      # PSParentPath looks like this: Microsoft.PowerShell.Core\Registry::HKEY_USERS\S-1-5-21-##########-##########-##########-####\Network
      $SID = ($Drive.PSParentPath -split '\\')[2]

      [PSCustomObject]@{
        # Use .NET to look up the username from the SID
        Username            = ([System.Security.Principal.SecurityIdentifier]"$SID").Translate([System.Security.Principal.NTAccount])
        DriveLetter         = $Drive.PSChildName
        RemotePath          = $Drive.RemotePath

        # The username specified when you use "Connect using different credentials".
        # For some reason, this is frequently "0" when you don't use this option. I remove the "0" to keep the results consistent.
        ConnectWithUsername = $Drive.UserName -replace '^0$', $null
        SID                 = $SID
      }

    }

  }
  else {

    Write-Verbose "No mapped drives were found"

  }
}

