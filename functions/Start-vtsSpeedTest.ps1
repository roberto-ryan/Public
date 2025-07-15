function Start-vtsSpeedTest {
  <#
  .DESCRIPTION
  Runs speedtest by Ookla. Installs via chocolatey.
  .EXAMPLE
  PS> Start-vtsSpeedTest
  
  Output:
     Speedtest by Ookla
  
       Server: Sparklight - Anniston, AL (id = 8829)
          ISP: Spectrum Business
      Latency:    25.11 ms   (0.18 ms jitter)
     Download:   236.80 Mbps (data used: 369.4 MB )
       Upload:   309.15 Mbps (data used: 526.0 MB )
  Packet Loss:     0.0%
   Result URL: https://www.speedtest.net/result/c/23d057dd-8de5-4d62-aef9-72beb122d7a4
  
   .LINK
  Network
   #>
  if (Test-Path "C:\ProgramData\chocolatey\bin\speedtest.exe") {
    C:\ProgramData\chocolatey\bin\speedtest.exe
  }
  elseif (Test-Path C:\ProgramData\chocolatey\lib\speedtest\tools\speedtest.exe) {
    C:\ProgramData\chocolatey\lib\speedtest\tools\speedtest.exe
  }
  elseif (Test-Path "C:\ProgramData\chocolatey\choco.exe") {
    choco install speedtest -y
    speedtest
  }
  else {
    Install-vtsChoco
    choco install speedtest -y
    speedtest
  }
}

