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
    $script:clipboard = New-Object -TypeName "System.Collections.ArrayList"

    function EnsureUserIsNotSystem {
        $identity = whoami.exe
        if ($identity -eq "nt authority\system") {
            Write-Host "`n`nThis script needs to be run as the logged-in user, not as SYSTEM.`n`n" -ForegroundColor Red
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
        Write-Host "`"Oh look, another tech genius expecting AI to do all the work."
        Write-Host "This tool? It's an assistant, not your magical unicorn."
        Write-Host "It helps jot down ticket notes, but isn't blessed with divine perfection."
        Write-Host "It is your responsibility to ensure your notes are accurate and thorough."
        Write-Host "Usage of this tool confirms you agree to this pact.`" -ChatGPT"

        Write-Host "`nFor optimal results, run this tool on the computer where the work is being performed."
        Write-Host "`nEnter 'r' to resume last session.`n" -ForegroundColor Yellow
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
        Write-Host "Tip: Any text copied to the logged-in user's clipboard during this session will be used to improve these notes.`n`n" -ForegroundColor Cyan
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
        (((Get-Content $PSRFile | select-string "^        <p><b>") -replace '^        <p><b>', '' -replace '</b>', '' -replace '</p>', '' -replace '&quot;', "'") -replace $regex | Select-String '^ User' | Select-Object -ExpandProperty Line | ForEach-Object { $_.Substring(1) }) -replace '\[.*?\]', '' -replace 'â€‹','' -replace 'User ','I ' | Out-File "$dir\steps.txt" -Encoding utf8
    }

    function CleanupSteps {
        # Clean up unwanted input from Steps Recorder"$dir\steps.txt"
        Get-Content "$dir\steps.txt" | 
        Where-Object { $_ -notmatch 'mouse drag|mouse wheel|\(pane\)' } | 
        Select-Object -Unique | 
        Out-File "$dir\cleaned_steps.txt" -Append -Encoding utf8

        #Remove last step as it's alway irrelevant
        $PSRResult = Get-Content "$dir\cleaned_steps.txt"
        $StepCount = $PSRResult.Count - 2
        $steps = $PSRResult[0..$StepCount]

        # Join steps with newline characters to remove blank lines
        $script:joinedSteps = ($steps | Select-Object -last 30) -join "`n"
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
Create ticket notes based on the following information:

Recorded Steps:
$($script:joinedSteps)

Clipped:
$(Get-Content "$dir\clipboard.txt")

Issue:
$(Get-Content "$dir\issue.txt")

Resolution:
$(Get-Content "$dir\resolution.txt")
"@ | ConvertTo-Json
        $prompt | Out-File "$dir\prompt.txt" -Encoding utf8
    }

    function APICall {
        $Headers = @{
            "Content-Type"  = "application/json"
            "Authorization" = "Bearer $OpenAIKey"
        }
        $Body = @{
            "model"             = "gpt-3.5-turbo"
            "messages"          = @( @{
                    "role"    = "system"
                    "content" = "You are a helpful IT technician that creates comprehensive ticket notes for IT support issues."
                },
                @{
                    "role"    = "system"
                    "content" = "You always respond in the first person, in the following format:\n\nReported Issue:<text describing issue\n\nCustomer Actions Taken:<customer actions here\n\nTroubleshooting Methods:\n- <bulletted troubleshooting steps here>\n\nResolution:<resolution here>\n\nComments & Misc. info:<miscellaneous info here>\n\nMessage to End User:\n<email to end user here>"
                },
                @{
                    "role"    = "system"
                    "content" = "Use the data in the Recorded Steps section to include printer names, website names, program names, software version numbers, etc., in the Troubleshooting Methods section."
                },
                @{
                    "role"    = "system"
                    "content" = "Use the Clipped section to add more detail to the notes. Add details to the Comments & Misc. section if they don't make sense in the Troubleshooting Methods section."
                },
                @{
                    "role"    = "system"
                    "content" = "Don't fill out the Customer Actions Taken section."
                },
                @{
                    "role"    = "system"
                    "content" = "The Troubleshooting Methods are written in the first-person by (you) the technician only."
                },
                @{
                    "role"    = "system"
                    "content" = "Always start the email with 'Hi <name>' and sign off with 'Respectfully,'"
                },
                @{
                    "role"    = "user"
                    "content" = "Here's an example of the output I want:\n\nIssue Reported: Screen flickering\n\nCustomer Actions Taken: None\n\nTroubleshooting Methods:\n- Checked for Windows Updates.\n- Navigated to the Device Manager, located Display Adapters and right-clicked on the NVIDIA GeForce GTX 1050, selecting Update Driver.\n- Clicked on Search Automatically for Drivers, followed by Search for Updated Drivers on Windows Update.\n- Searched for 'gtx 1050 drivers' and clicked on the first result.\n- Clicked on the Official Drivers link and downloaded the driver.\n- Updated the graphics driver, resolving the issue.\n\nResolution: Updating the graphics driver resolved the issue.\n\nAdditional Comments: None\n\n\nMessage to End User: \n\n[User Name],\n\nWe have successfully resolved the screen flickering issue you were experiencing by updating the graphics driver. At your earliest convenience, please test your system to confirm that the issue with your screen has been rectified. Should you encounter any additional issues or require further assistance, do not hesitate to reach out to us.\n\nRespectfully,"
                },
                @{
                    "role"    = "user"
                    "content" = "Here's another example of the output I want:\n\nIssue Reported: Scanner not working\n\nCustomer Actions Taken: None\n\nTroubleshooting Methods:\n- Accessed the 'Printers and Scanners' Settings menu, located the scanner in the list of devices and right-clicked on it, selecting Properties.\n- Determined the scanner was not being recognized by the computer.\n- Had the user unplug the scanner from the computer and plug it back in.\n- Reinstalled the scanner driver and rebooted.\n\nResolution: Unfortunately, the issue remains unresolved, as the computer is unable to recognize the scanner. This ticket will now be transferred to the onsite support queue for further assistance.\n\nAdditional Comments: None\n\nMessage to End User: \n\n[User Name],\n\nWe have been unable to resolve the issue with your scanner, as the computer is not recognizing the device. To further investigate and address this problem, we will require an onsite technician to visit your location. Should you have any questions or concerns, please feel free to reach out.\n\nRespectfully,"
                },
                @{
                    "role"    = "user"
                    "content" = "Here's another example of the output I want:\n\nIssue Reported: Slow internet connection\n\nCustomer Actions Taken: None\n\nTroubleshooting Methods:\n- Opened Command Prompt and enterted 'ipconfig /flushdns' to flush the DNS cache.\n- Closed Command Prompt and opened the Start menu, navigating to Settings, chose Network & Internet, and clicked on Change Adapter Options.\n- Right-clicked on the active network connection and selected Properties.\n- Clicked on Internet Protocol Version 4 (TCP/IPv4) and selected Properties.\n- Changed the Preferred DNS server to 8.8.8.8 (Google DNS) and the Alternate DNS server to 8.8.4.4, then clicked OK.\n- Tested internet connectivity, confirming that the issue was resolved.\n\nResolution: Flushing the DNS cache and changing DNS servers resolved the issue.\n\nAdditional Comments: None\n\nMessage to End User:\n\n[User Name],\n\nWe have successfully addressed the slow internet connection issue you reported by clearing the DNS cache and updating the DNS servers. When you have a moment, please check your internet connection to verify that the issue has been resolved. If you come across any further problems or need additional support, please don't hesitate to contact us.\n\nRespectfully,"
                },
                @{
                    "role"    = "user"
                    "content" = "$script:prompt"
                },
                @{
                    "role"    = "assistant"
                    "content" = ""
                })
            "temperature"       = 0
            'top_p'             = 1.0
            'frequency_penalty' = 0.0
            'presence_penalty'  = 0.0
            'stop'              = @('"""')
        } | ConvertTo-Json
        
        try {
            
            $script:response = Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Method Post -Body $Body -Headers $Headers
        }
        catch {
            Write-Error "$($_.Exception.Message)"
        }
    }

    function WriteResultsToFile {
        "Session Time: $SessionTime`n" | Out-File "$dir\result_header.txt" -Force -Encoding utf8
        "User Name: $env:USERDOMAIN\$env:USERNAME" | Out-File "$dir\result_header.txt" -Force -Encoding utf8 -Append
        "Computer Name: $env:COMPUTERNAME" | Out-File "$dir\result_header.txt" -Force -Encoding utf8 -Append
        "$($script:response.choices.message.content)" | Out-File "$dir\gpt_result.txt" -Force
    }

    function WriteResultsToHost {
        #Write final results to the shell
        Start-sleep -Milliseconds 250
        #Clear-Host
        (Get-Content "$dir\result_header.txt") | ForEach-Object { Write-Host $_ }
        (Get-Content "$dir\gpt_result.txt") | ForEach-Object { Write-Host $_ }
        Write-Host "`nToken Usage: Prompt=$($response.usage.prompt_tokens) Completion=$($response.usage.completion_tokens) Total=$($response.usage.total_tokens) Cost=`$$(($response.usage.total_tokens / 1000) * 0.002)" -ForegroundColor Gray
    }

    function Cleanup {
        Start-sleep -Milliseconds 250
        Get-Process -Name psr | Stop-Process -Force
        Get-ChildItem -path $dir -include "*.mht", "*.zip" -Recurse -File | Remove-Item -Recurse -Force -Confirm:$false
    }

    function GetClipboard {
        if (" " -ne $(Get-Clipboard)) { ($script:clipboard).add("$(Get-Clipboard)`n") | Out-Null }
    }
    
    EnsureUserIsNotSystem
    
    try {
        $SessionStart = Timestamp
        DisplayLogo
        $issue = Read-Host "Enter Ticket Description"

        if ($issue -ne 'r') {
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
        set-clipboard " "
        While ($true) {
            start-sleep -Milliseconds 250
            GetClipboard
        }
    }
    finally {
        StopStepsRecorder
        DisplayRecordingCompleteBanner
        $resolution = Read-Host "Enter Session Conclusion"
        Add-Content -Path "$dir\resolution.txt" -Value $resolution -Force
        DisplayProcessingBanner
        ParseSteps
        CleanupSteps
        $script:clipboard | Select-Object -unique | Out-File -FilePath "$dir\clipboard.txt" -Force -Encoding utf8 -Append
        $SessionEnd = Timestamp
        CalculateSessionTime
        GeneratePrompt
        APICall
        Clear-Host
        WriteResultsToFile
        WriteResultsToHost
        Cleanup
    }
}