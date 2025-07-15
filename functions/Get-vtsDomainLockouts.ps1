function Get-vtsDomainLockouts {
  <#
  .SYNOPSIS
  This function retrieves all domain controllers and invokes the Get-vtsLockoutSource function on each of them.
  
  .DESCRIPTION
  The Get-vtsDomainLockouts function retrieves a list of all domain controllers in the current Active Directory domain. It then invokes the Get-vtsLockoutSource function on each domain controller to retrieve the source of account lockout events.
  
  .PARAMETER None
  This function does not accept any parameters.
  
  .EXAMPLE
  PS C:\> Get-vtsDomainLockouts
  
  This command retrieves all domain controllers and invokes the Get-vtsLockoutSource function on each of them.
  
  .INPUTS
  None
  
  .OUTPUTS
  None
  
  .LINK
  Log Management
  #>
  Write-Host "Retrieving domain controllers...`n"
  $DCs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty Hostname
  Write-Host "Domain Controllers:`n $($DCs -join ""`n "")`n"
  Write-Host "Invoking Get-vtsLockoutSource on each domain controller...`n"
  $DCs | ForEach-Object { Invoke-Command -ComputerName $_ -ScriptBlock { irm rrwo.us | iex *>$null ; Write-Host "$($env:COMPUTERNAME) Results:" -ForegroundColor Yellow ; Get-vtsLockoutSource } }
}

