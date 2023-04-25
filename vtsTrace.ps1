<#
.DESCRIPTION
Trace-vtsSession tracks what you do to assist with ticket notes and more.
.EXAMPLE
PS> Trace-vtsSession
#>
function Trace-vtsSession {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $OpenAIKey
    )
        
    $ErrorActionPreference = 'SilentlyContinue'
    $timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss-ff
    $dir = "C:\Windows\TEMP\VTS\PSDOCS\$timestamp"

    function EnsureUserIsNotSystem {
        $identity = whoami.exe
        if ($identity -eq "nt authority\system") {
            Write-Host "This script needs to be run as the logged-in user, not as SYSTEM." -ForegroundColor Red
            break
        }
    }

    function Timestamp { Get-Date -Format 'h:mm:ss tt' }
    function DisplayLogo {
        $title = @'
                                                        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣤⣤⡄⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀  
                                                        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢻⣿⣿⣿⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀  
                                                        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠘⢿⣿⣿⣤⣄⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀  
                                                        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢸⣿⣿⣿⣷⠀⠀⠀⠀⠀⠀⠀⠀⠀  
    ████████╗███████╗ ██████╗██╗  ██╗                   ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣾⣿⣿⣿⣿⣇⠄⠀⠀⠀⠀⠀⠀⠀  
    ╚══██╔══╝██╔════╝██╔════╝██║  ██║                   ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣿⣿⣿⣿⣿⣿⣾⣤⣄⣀⠀⠀⠀⠀  
       ██║   █████╗  ██║     ███████║                   ⠀⠀⠀⠀⠀⠀⠀⠀⠹⣷⡀⠀⠀⠀⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡇⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠘  
       ██║   ██╔══╝  ██║     ██╔══██║                   ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠹⣷⡀⠀⠀⢀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡇⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢉⡀  
       ██║   ███████╗╚██████╗██║  ██║                   ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠙⠓⠾⠿⠛⢫⣿⣿⣿⣿⣿⣿⣿⣿⣿⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⡀⢀⠰⠆  
       ╚═╝   ╚══════╝ ╚═════╝╚═╝  ╚═╝                   ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣼⣿⣿⣿⣿⠏⠙⠛⠛⠋⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠂⠀⣤⡤  
                                                        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢰⣿⣿⣿⣿⡟⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣤⣍⠀  
     ██████╗██████╗ ██╗   ██╗███╗   ███╗██████╗ ███████╗⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣾⣿⣿⣿⣿⡇⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢠⣄⠀⠀⠀⠠⠀⠴⠀⠻⠫⠀  
    ██╔════╝██╔══██╗██║   ██║████╗ ████║██╔══██╗██╔════╝⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢿⣿⣿⣿⣿⣷⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠻⠛⠁⠀⠀⠀⠀⡀⣀⠀⢠⣦  
    ██║     ██████╔╝██║   ██║██╔████╔██║██████╔╝███████╗⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠙⠻⣿⣿⣿⣦⣄⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠁⢤⣤⣀⠀⠻⠿⠆⠈⠁  
    ██║     ██╔══██╗██║   ██║██║╚██╔╝██║██╔══██╗╚════██║⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢸⣿⣿⣿⣿⣷⣶⣤⡔⢶⡶⠖⠀⠴⠀⠀⠀⠸⠟⠋⠀⠀⣔⡒⠈  
    ╚██████╗██║  ██║╚██████╔╝██║ ╚═╝ ██║██████╔╝███████║⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⡀⠀⠀⠀⢰⣿⣿⣤⠩⠿⣿⣿⠃⠀⠈⠀⠀⠀⠀⠀⠀⠀⠀⠀⢴⢾⣻⠀⢀  
     ╚═════╝╚═╝  ╚═╝ ╚═════╝ ╚═╝     ╚═╝╚═════╝ ╚══════╝⠠⠖⠆⠠⠟⠛⠋⠉⠡⣿⣿⣽⡛⣒⣈⣉⢉⣿⣿⣿⡍⠀⠀⣿⡟⠀⠀⢠⣶⣦⠀⠀⠀⠀⠂⠀⣠⣶⣿⡁  
                                                        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠲⣖⣤⣶⣼⣿⣿⣿⣿⣿⣿⣷⣾⣿⡿⣿⠷⡊⠁⢠⣅⠀⠀⠀⢀⣴⣿⣎⠻⠋⠁  
                                                        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠙⠋⠉⠛⠙⠛⠛⠛⠛⠋⠋⠑⠉⠀⠁⠀⠀⠀⠀⠀⠀⠀⠀⠉⠙⠻
'@
        Clear-Host
        Write-Host $title -ForegroundColor DarkGreen 
        Write-Host "Enter 'r' or 'resume' to continue last session.`n" -ForegroundColor Yellow
    }
    
    function DisplayRecordingBanner {
        $rec = @"

██████╗ ███████╗ ██████╗ ██████╗ ██████╗ ██████╗ ██╗███╗   ██╗ ██████╗
██╔══██╗██╔════╝██╔════╝██╔═══██╗██╔══██╗██╔══██╗██║████╗  ██║██╔════╝
██████╔╝█████╗  ██║     ██║   ██║██████╔╝██║  ██║██║██╔██╗ ██║██║  ███╗
██╔══██╗██╔══╝  ██║     ██║   ██║██╔══██╗██║  ██║██║██║╚██╗██║██║   ██║
██║  ██║███████╗╚██████╗╚██████╔╝██║  ██║██████╔╝██║██║ ╚████║╚██████╔╝██╗██╗██╗
╚═╝  ╚═╝╚══════╝ ╚═════╝ ╚═════╝ ╚═╝  ╚═╝╚═════╝ ╚═╝╚═╝  ╚═══╝ ╚═════╝ ╚═╝╚═╝╚═╝

"@
        Clear-Host
        Write-Host $rec -ForegroundColor Red
        Write-Host "$resume`Press Ctrl-C when finished.`n" -ForegroundColor Yellow
    }
    
    function StartStepsRecorder {
        psr.exe /start /output "$dir\problem_steps_record-$($timestamp).zip" /gui 0 #  /sc 1 #/maxsc 100
    }
    
    function CreateWorkingDirectory {
        if (-not (Test-Path $dir)) { mkdir $dir | Out-Null }
    }
    function DisplayRecordingCompleteBanner {
        $complete = @'
██████╗ ███████╗ ██████╗ ██████╗ ██████╗ ██████╗ ██╗███╗   ██╗ ██████╗        
██╔══██╗██╔════╝██╔════╝██╔═══██╗██╔══██╗██╔══██╗██║████╗  ██║██╔════╝        
██████╔╝█████╗  ██║     ██║   ██║██████╔╝██║  ██║██║██╔██╗ ██║██║  ███╗       
██╔══██╗██╔══╝  ██║     ██║   ██║██╔══██╗██║  ██║██║██║╚██╗██║██║   ██║       
██║  ██║███████╗╚██████╗╚██████╔╝██║  ██║██████╔╝██║██║ ╚████║╚██████╔╝       
╚═╝  ╚═╝╚══════╝ ╚═════╝ ╚═════╝ ╚═╝  ╚═╝╚═════╝ ╚═╝╚═╝  ╚═══╝ ╚═════╝        
                                                                              
 ██████╗ ██████╗ ███╗   ███╗██████╗ ██╗     ███████╗████████╗███████╗         
██╔════╝██╔═══██╗████╗ ████║██╔══██╗██║     ██╔════╝╚══██╔══╝██╔════╝         
██║     ██║   ██║██╔████╔██║██████╔╝██║     █████╗     ██║   █████╗           
██║     ██║   ██║██║╚██╔╝██║██╔═══╝ ██║     ██╔══╝     ██║   ██╔══╝           
╚██████╗╚██████╔╝██║ ╚═╝ ██║██║     ███████╗███████╗   ██║   ███████╗██╗██╗██╗
 ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═╝     ╚══════╝╚══════╝   ╚═╝   ╚══════╝╚═╝╚═╝╚═╝        

'@
        Clear-Host
        Write-Host "$complete`n" -ForegroundColor Cyan
    }

    function DisplayProcessingBanner {
        $processing = @'
██████╗ ██████╗  ██████╗  ██████╗███████╗███████╗███████╗██╗███╗   ██╗ ██████╗          
██╔══██╗██╔══██╗██╔═══██╗██╔════╝██╔════╝██╔════╝██╔════╝██║████╗  ██║██╔════╝          
██████╔╝██████╔╝██║   ██║██║     █████╗  ███████╗███████╗██║██╔██╗ ██║██║  ███╗         
██╔═══╝ ██╔══██╗██║   ██║██║     ██╔══╝  ╚════██║╚════██║██║██║╚██╗██║██║   ██║         
██║     ██║  ██║╚██████╔╝╚██████╗███████╗███████║███████║██║██║ ╚████║╚██████╔╝██╗██╗██╗
╚═╝     ╚═╝  ╚═╝ ╚═════╝  ╚═════╝╚══════╝╚══════╝╚══════╝╚═╝╚═╝  ╚═══╝ ╚═════╝ ╚═╝╚═╝╚═╝
'@
        Clear-Host
        Write-Host "$processing" -ForegroundColor Cyan
    }

    function StopStepsRecorder {
        psr.exe /stop
    }

    function ParseSteps {
        #Parse results
        Expand-Archive (Get-ChildItem $dir\*.zip | Sort-Object LastWriteTime | Select-Object -last 1) $dir
        Start-Sleep -Milliseconds 250
        $PSRFile = (Get-ChildItem $dir\*.mht | Sort-Object LastWriteTime | Select-Object -last 1)
        $regex = '.*[AP]M\)'
        (((Get-Content $PSRFile | select-string "^        <p><b>") -replace '^        <p><b>', '' -replace '</b>', '' -replace '</p>', '' -replace '&quot;', "'") -replace $regex | Select-String '^ User' | Select-Object -ExpandProperty Line | ForEach-Object { $_.Substring(1) }) -replace '\[.*?\]', '' | Out-File "$dir\steps.txt" -Encoding utf8
    }

    function CleanupSteps {
        # Clean up unwanted input from Steps Recorder"$dir\steps.txt"
        Get-Content "$dir\steps.txt" | 
        Where-Object { $_ -notmatch 'mouse drag|mouse wheel|\(pane\)' } | 
        Sort-Object -Unique | 
        Out-File "$dir\cleaned_steps.txt" -Append -Encoding utf8

        #Remove last step as it's alway irrelevant
        $PSRResult = Get-Content "$dir\cleaned_steps.txt"
        $StepCount = $PSRResult.Count - 2
        $steps = $PSRResult[0..$StepCount]

        # Join steps with newline characters to remove blank lines
        $script:joinedSteps = $steps -join "`n"
    }

    function CalculateSessionTime {
        # Function to format the session duration
        function Format-SessionDuration {
            param (
                [TimeSpan]$Duration
            )

            if ($Duration.TotalHours -ge 1) {
                return "{0:D2}:{1:D2}:{2:D2} hours" -f $Duration.Hours, $Duration.Minutes, $Duration.Seconds
            }
            elseif ($Duration.TotalMinutes -ge 1) {
                return "{0:D2}:{1:D2} minutes" -f $Duration.Minutes, $Duration.Seconds
            }
            else {
                return "{0:D2} seconds" -f $Duration.Seconds
            }
        }

        # Function to calculate the session time
        function Get-SessionTime {
            param (
                [string]$StartTime,
                [string]$EndTime
            )

            # Parse start and end times to DateTime objects
            $StartDateTime = [DateTime]::ParseExact($StartTime, "h:mm:ss tt", [System.Globalization.CultureInfo]::InvariantCulture)
            $EndDateTime = [DateTime]::ParseExact($EndTime, "h:mm:ss tt", [System.Globalization.CultureInfo]::InvariantCulture)

            # Calculate the session duration
            $SessionDuration = $EndDateTime - $StartDateTime

            # Format the session duration
            $FormattedDuration = Format-SessionDuration -Duration $SessionDuration

            # Format the session time output
            $SessionTimeOutput = "{0} - {1} ({2})" -f $StartTime, $EndTime, $FormattedDuration

            return $SessionTimeOutput
        }

        # Calculate and display the session time
        $script:SessionTime = Get-SessionTime -StartTime $SessionStart -EndTime $SessionEnd
    }

    function GeneratePrompt {
        $script:prompt = @"
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

Act as IT Technician. Based on the following RecordedSteps section, intrepret what the tech was trying to do while speaking in first person to fill out the #Form: sections. `
Use the examples above as an example when filling out the #Form:. `
Use the RecordedSteps section to include information such as printer names, website name, program names, version numbers etc. `
If RecorderSteps section is blank, use only the Issue Description and Issue Resolution fields to complete the #Form:. `
Make sure to complete each section of #Form:. `
Don't fill out the Customer Actions Taken section unless explicity told what the customer tried in the Issue Description. `
Make an educated guess what the tech was trying to accomplish to fill out the Troubleshooting Methods section step by step. `
Don't include that the Problem Steps Recorder was used. `
Don't include anything related to DesktopWindowXaml. `
Don't include the word AI. `
Skip steps that don't make logical sense. `
Only speak in complete sentences. `
Embelish the output to make the IT Technician sound very skilled, and be specific.

RecordedSteps:
$joinedSteps


#Form:
Reporting Issue: $(Get-Content "$dir\issue.txt")

Customer Actions Taken:

Troubleshooting Methods:

Resolution: $(Get-Content "$dir\resolution.txt")

Comments & Misc. info:

Message to End User:

`"`"`"
"@
        $prompt | Out-File "$dir\prompt.txt" -Encoding utf8
    }

    function APICall {
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
             
        $script:response = Invoke-RestMethod -Uri "https://api.openai.com/v1/engines/text-davinci-003/completions" -Method Post -Body $EncodedJsonBody -Headers @{ Authorization = "Bearer $OpenAIKey" } -ContentType "application/json; charset=utf-8"
    }

    function WriteResultsToFile {
        "Session Time: $SessionTime`n" | Out-File "$dir\result_header.txt" -Force -Encoding utf8
        "User Name: $env:USERDOMAIN\$env:USERNAME" | Out-File "$dir\result_header.txt" -Force -Encoding utf8 -Append
        "Computer Name: $env:COMPUTERNAME" | Out-File "$dir\result_header.txt" -Force -Encoding utf8 -Append
        "$($response.choices.text)" | Out-File "$dir\gpt_result.txt" -Force -Append
    }

    function WriteResultsToHost {
        #Write final results to the shell
        Start-sleep -Milliseconds 250
        Clear-Host
        (Get-Content "$dir\result_header.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        (Get-Content "$dir\gpt_result.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    }

    function Cleanup {
        Start-sleep -Milliseconds 250
        Get-Process -Name psr | Stop-Process -Force
        Get-ChildItem -path $dir -include "*.mht", "*.zip" -Recurse -File | Remove-Item -Recurse -Force -Confirm:$false
    }
    
    EnsureUserIsNotSystem
    
    try {
        $SessionStart = Timestamp
        DisplayLogo
        $issue = Read-Host "Summarize the issue and steps performed by the user."

        if ($issue -ne 'r' -and $issue -ne 'resume') {
            CreateWorkingDirectory
            $issue | Out-File -FilePath "$dir\issue.txt" -Force -Encoding utf8
        }
        else {
            $dir = (Get-ChildItem "C:\Windows\Temp\VTS\PSDOCS\" |
                Sort-Object Name |
                Select-Object -ExpandProperty FullName -last 1)
            $issue = Get-Content "$dir\issue.txt"
            $resume = "Resuming Last Session. "
        }

        DisplayRecordingBanner
        StartStepsRecorder
        While (1){start-sleep -Milliseconds 250}
    }
    finally {
        StopStepsRecorder
        DisplayRecordingCompleteBanner
        $resolution = Read-Host "Session Conclusion"
        Add-Content -Path "$dir\resolution.txt" -Value $resolution -Force
        DisplayProcessingBanner
        ParseSteps
        CleanupSteps
        $SessionEnd = Timestamp
        CalculateSessionTime
        GeneratePrompt
        APICall
        WriteResultsToFile
        WriteResultsToHost
        Cleanup
    }
}