function GPTFollowUp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $OpenAIKey
    )

    $ErrorActionPreference = 'SilentlyContinue'

    $dir = (Get-ChildItem "C:\Windows\Temp\VTS\PSDOCS\" |
        Sort-Object Name |
        Select-Object -ExpandProperty FullName -last 1)
        
    function WriteResultsToHost {
        Clear-Host
        (Get-Content "$dir\result_header.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        (Get-Content "$dir\gpt_result.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    }

    While ($true) {
        $alterations = Read-Host "GPT3.5>>>"
        
        switch ($alterations) {
            steps {
                "`n"
                Get-Content $dir\cleaned_steps.txt 
                "`n"
            }
            Default {
                $prompt = @"
Here are 3 examples of the output I am looking for:

Example1 = (
Issue Reported: Screen flickering

Customer Actions Taken: None

Troubleshooting Methods:
- Called the user and established a remote session using Teamviewer.
- Checked for Windows Updates.
- Navigated to the Device Manager, located Display Adapters and right-clicked on the NVIDIA GeForce GTX 1050, selecting Update Driver.
- Clicked on Search Automatically for Drivers, followed by Search for Updated Drivers on Windows Update.
- Searched for 'gtx 1050 drivers' and clicked on the first result.
- Clicked on the Official Drivers link and downloaded the driver.
- Updated the graphics driver, resolving the issue.

Resolution: Updating the graphics driver resolved the issue.

Additional Comments: None


Message to End User: 

[User Name],

We have successfully resolved the screen flickering issue you were experiencing by updating the graphics driver. At your earliest convenience, please test your system to confirm that the issue with your screen has been rectified. Should you encounter any additional issues or require further assistance, do not hesitate to reach out to us.

Respectfully,
[Technician Name]
)

Example2 = (
Issue Reported: Scanner not working

Customer Actions Taken: None

Troubleshooting Methods:
- Called the user and established a remote session using Teamviewer.
- Accessed the 'Printers and Scanners' Settings menu, located the scanner in the list of devices and right-clicked on it, selecting Properties.
- Determined the scanner was not being recognized by the computer.
- Had the user unplug the scanner from the computer and plug it back in.
- Reinstalled the scanner driver and rebooted.

Resolution: Unfortunately, the issue remains unresolved, as the computer is unable to recognize the scanner. This ticket will now be transferred to the onsite support queue for further assistance.

Additional Comments: None

Message to End User: 

[User Name],

We have been unable to resolve the issue with your scanner, as the computer is not recognizing the device. To further investigate and address this problem, we will require an onsite technician to visit your location. Should you have any questions or concerns, please feel free to reach out.

Respectfully,
[Technician Name]
)

Example3 = (
Issue Reported: Slow internet connection

Customer Actions Taken: None

Troubleshooting Methods:
- Called the user and established a remote session using Teamviewer.
- Opened Command Prompt by searching for 'cmd' in the Start menu and running it as Administrator.
- Typed 'ipconfig /flushdns' and pressed Enter to flush the DNS cache.
- Closed Command Prompt and opened the Start menu, navigating to Settings.
- Chose Network & Internet, and clicked on Change Adapter Options.
- Right-clicked on the active network connection and selected Properties.
- Clicked on Internet Protocol Version 4 (TCP/IPv4) and selected Properties.
- Changed the Preferred DNS server to 8.8.8.8 (Google DNS) and the Alternate DNS server to 8.8.4.4, then clicked OK.
- Closed the Network Connections window and restarted the computer.
- Tested internet connectivity, confirming that the issue was resolved.

Resolution: Flushing the DNS cache and changing DNS servers resolved the issue.

Additional Comments: None

Message to End User:

[User Name],

We have successfully addressed the slow internet connection issue you reported by clearing the DNS cache and updating the DNS servers. When you have a moment, please check your internet connection to verify that the issue has been resolved. If you come across any further problems or need additional support, please don't hesitate to contact us.

Respectfully,
[Technician Name]
)


Rewrite the form below. Keep all information the same, but take into account this additional information: 
$alterations.


$(Get-Content $dir\gpt_result.txt -Encoding utf8)

`"`"`"
"@
    
                $body = @{
                    'prompt'            = $prompt;
                    'temperature'       = 0;
                    'max_tokens'        = 500;
                    'top_p'             = 1.0;
                    'frequency_penalty' = 0.0;
                    'presence_penalty'  = 0.0;
                    'stop'              = @('"""');
                }
        
                $JsonBody = $body | ConvertTo-Json -Compress
                $EncodedJsonBody = [System.Text.Encoding]::UTF8.GetBytes($JsonBody)

                $response = Invoke-RestMethod -Uri "https://api.openai.com/v1/engines/text-davinci-003/completions" -Method Post -Body $EncodedJsonBody -Headers @{ Authorization = "Bearer $OpenAIKey" } -ContentType "application/json; charset=utf-8"
                "$($response.choices.text)" | Out-File "$dir\gpt_result.txt" -Force -Encoding utf8
                WriteResultsToHost
            }
        }
    }
}