function Get-vtsReadablePermissions {
  <#
  .SYNOPSIS
      Retrieves readable permissions for a given directory path.
  
  .DESCRIPTION
      The Get-vtsReadablePermissions function takes a file system path as input and returns a custom object array with the permissions of each subfolder. It includes details such as the user or group with access, the type of access, and the inheritance and propagation of the permissions.
  
  .EXAMPLE
      PS C:\> Get-vtsReadablePermissions -Path "C:\MyFolder"
      This example retrieves the permissions for all the immediate child folders of "C:\MyFolder".
  
  .NOTES
      This function requires at least PowerShell version 3.0 to run properly due to the usage of the Get-ChildItem -Directory parameter.
  
  .LINK
      File Management
  #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Check if the path exists
    if (!(Test-Path $Path)) {
        Write-Error "The specified path does not exist."
        return
    }

    # Initialize an empty array to store the permission information
    $permissions = @()

    $ChildFolders = Get-ChildItem -Path $Path -Directory -Depth 1 | Select-Object -expand fullname

    foreach ($Folder in $ChildFolders) {
        # Retrieve the ACL for the specified path
        $acl = Get-Acl -Path $Folder

        # Loop through each Access Control Entry (ACE)
        foreach ($ace in $acl.Access) {
            # Translate the IdentityReference (user or group)
            $user = $ace.IdentityReference
            $accessRights = $ace.FileSystemRights
            $accessType = $ace.AccessControlType

            # Translate Inheritance Flags
            $inheritance = switch ($ace.InheritanceFlags) {
                "None" { "This Folder Only" }
                "ContainerInherit" { "This Folder and Subfolders" }
                "ObjectInherit" { "This Folder and Files" }
                "ContainerInherit, ObjectInherit" { "This Folder, Subfolders, and Files" }
                default { "Unknown type" }
            }

            # Translate Propagation Flags
            $propagation = switch ($ace.PropagationFlags) {
                "None" { "None" }
                "NoPropagateInherit" { "Does not pass down" }
                "InheritOnly" { "Only affects children" }
                default { "Unknown propagation type" }
            }

            # Add to the permissions array
            $permissions += [PSCustomObject]@{
                Folder      = $Folder
                UserOrGroup = $user
                AccessType  = $accessType
                Rights      = $accessRights
                Inheritance = $inheritance
                Propagation = $propagation
            }
        }
    }

    # Output the permission information in a readable format
    $FilteredPermissions = $permissions |
    Where-Object UserOrGroup -ne "NT AUTHORITY\SYSTEM" |
    Where-Object UserOrGroup -ne "BUILTIN\Administrators" |
    Where-Object UserOrGroup -ne "CREATOR OWNER" |
    Where-Object UserOrGroup -notlike "*S-1-5*" |
    Sort-Object Folder, Rights, AccessType, Inheritance, UserOrGroup 
    
    $FilteredPermissions |
    Format-Table -AutoSize

    # Ask user if they want to export a report
    $exportReport = Read-Host -Prompt "Do you want to export a report? (Y/N)"
  
    if ($exportReport -eq "Y" -or $exportReport -eq "y") {
        # Set TLS1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        # Prompt for filepath
        $Filepath = Read-Host "Enter a file path to save the report (e.g., C:\Reports)"

        # Set alternate $Filepath if null or whitepace
        if ([string]::IsNullOrWhiteSpace($Filepath)) {
            Write-Host "`$Filepath is empty. Defaulting to C:\Windows\TEMP"
            $Filepath = "C:\Windows\TEMP" 
        }

        # Set alternate $Filepath if not valid
        if (-not(Test-Path $Filepath)) {
            Write-Host "$Filepath is not a valid path. Defaulting to C:\Windows\TEMP"
            $Filepath = "C:\Windows\TEMP" 
        }
        # Check if PSWriteHTML module is installed, if not, install it
        if (!(Get-InstalledModule -Name PSWriteHTML 2>$null)) {
            Install-Module -Name PSWriteHTML -Force -Confirm:$false
        }
        
        # Export the results to an HTML file using the PSWriteHTML module
        $FilteredPermissions | Out-HtmlView -Title "Permission Report - $(Get-Date -f "dddd MM-dd-yyyy HHmm")" -Filepath "$Filepath\Permission Report - $(Get-Date -f MM_dd_yyyy).html"
    }
}

