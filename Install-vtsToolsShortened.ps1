[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; iex (iwr "https://raw.githubusercontent.com/roberto-ryan/Public/main/Install-vtsTools.ps1" -useb).Content
