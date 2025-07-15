function Start-vtsPacketCapture {
  <#
  .SYNOPSIS
  Starts a packet capture using Wireshark's tshark.
  
  .DESCRIPTION
  The Start-vtsPacketCapture function starts a packet capture on the specified network interface. If Wireshark is not installed, it will install it along with Chocolatey and Npcap.
  
  .PARAMETER interface
  The network interface to capture packets from. Defaults to the first Ethernet interface that is up.
  
  .PARAMETER output
  The path to the output file. Defaults to a .pcap file in C:\temp with the computer name and current date and time.
  
  .EXAMPLE
  Start-vtsPacketCapture -interface "Ethernet" -output "C:\temp\capture.pcap"
  Starts a packet capture on the Ethernet interface, with the output saved to C:\temp\capture.pcap.
  
  .LINK
  Network
  #>
  param (
    [string]$interface = (get-netadapter | Where-Object Name -like "*ethernet*" | Where-Object Status -eq Up | Select-Object -expand name),
    [string]$output = "C:\temp\$($env:COMPUTERNAME)-$(Get-Date -f hhmm-MM-dd-yyyy)-capture.pcap"
  )

  Write-Host "Starting packet capture..."

  if (!(Test-Path "C:\Program Files\Wireshark\tshark.exe")) {
    Write-Host "Wireshark not found. Installing necessary components..."

    #Install choco
    Write-Host "Installing Chocolatey..."
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

    #Install Wireshark
    Write-Host "Installing Wireshark..."
    choco install wireshark -y

    #Install npcap
    Write-Host "Installing Npcap..."
    mkdir C:\temp
    Invoke-WebRequest -Uri "https://npcap.com/dist/npcap-1.79.exe" -UseBasicParsing -OutFile "C:\temp\npcap-1.79.exe"

    & "C:\temp\npcap-1.79.exe"

    $wshell = New-Object -ComObject wscript.shell
    Start-sleep -Milliseconds 250
    $wshell.AppActivate('Npcap 1.79 Setup')
    $wshell.SendKeys('%a');
    Start-sleep -Milliseconds 250
    $wshell.AppActivate('Npcap 1.79 Setup')
    $wshell.SendKeys('%i');
    Start-sleep -Milliseconds 250
    $wshell.AppActivate('Npcap 1.79 Setup')
    $wshell.SendKeys('%n');
    Start-sleep -Milliseconds 250
    $wshell.AppActivate('Npcap 1.79 Setup')
    $wshell.SendKeys('{enter}');
  }
    
  #Start packet capture using tshark on ethernet NIC
  $tsharkPath = "C:\Program Files\Wireshark\tshark.exe"
    
  Write-Host "Starting packet capture on interface $interface, output to $output"
  Start-Process -FilePath $tsharkPath -ArgumentList "-i $interface -w $output" -WindowStyle Hidden
}

