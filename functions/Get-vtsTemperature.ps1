function Get-vtsTemperature {
  <#
  .DESCRIPTION
  Returns temperature of thermal sensor on motherboard. Not accurate for CPU temp.
  .EXAMPLE
  PS> Get-vtsTemperature
  
  Output:
  27.85 C : 82.1300000000001 F : 301K
  
  .LINK
  System Information
  #>
  $t = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi"
  $returntemp = @()

  foreach ($temp in $t.CurrentTemperature) {


    $currentTempKelvin = $temp / 10
    $currentTempCelsius = $currentTempKelvin - 273.15

    $currentTempFahrenheit = (9 / 5) * $currentTempCelsius + 32

    $returntemp += $currentTempCelsius.ToString() + " C : " + $currentTempFahrenheit.ToString() + " F : " + $currentTempKelvin + "K"  
  }
  return $returntemp
}

