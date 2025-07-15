function Get-vtsWlanProfilesAndKeys {
  <#
  .SYNOPSIS
  Retrieves the network profiles and their associated keys on the local machine.
  
  .DESCRIPTION
  This function lists all wireless network profiles stored on the local machine along with their clear text keys (passwords). It uses the 'netsh' command-line utility to query the profiles and extract the information.
  
  .EXAMPLE
  PS C:\> Get-vtsWlanProfilesAndKeys
  
  This command will display a list of all wireless network profiles and their associated keys.
  
  .NOTES
  This function requires administrative privileges to reveal the keys for the network profiles.
  
  .LINK
  Network
  #>
    netsh wlan show profiles | 
    Where-Object {$_ -match ' : '} | 
    ForEach-Object {$_.split(':')[1].trim()} | 
    ForEach-Object {
        $networkName = $_
        netsh wlan show profile name="$_" key=clear
    } | 
    Where-Object {$_ -match 'key content'} | 
    Select-Object @{Name='Network'; Expression={$networkName}}, @{Name='Key'; Expression={$_.split(':')[1].trim()}}
}

