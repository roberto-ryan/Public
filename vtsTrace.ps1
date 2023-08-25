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
        $OpenAIKey,
        [switch]$RecordSession
    )
        
    $ErrorActionPreference = 'SilentlyContinue'
    $timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss-ff
    $script:dir = "C:\Windows\TEMP\VTS\PSDOCS\$timestamp"
    $script:clipboard = New-Object -TypeName "System.Collections.ArrayList"

    function Read-String ($maxLength = 65536) {
        $str = ""
        $inputStream = [System.Console]::OpenStandardInput($maxLength);
        $bytes = [byte[]]::new($maxLength);
        while ($true) {
            $len = $inputStream.Read($bytes, 0, $maxLength);
            $str += [string]::new($bytes, 0, $len)
            if ($str.EndsWith("`r`n")) {
                $str = $str.Substring(0, $str.Length - 2)
                return $str
            }
        }
    }   

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
        psr.exe /start /output "$script:dir\problem_steps_record-$($timestamp).zip" /gui 0 #  /sc 1 #/maxsc 100
    }
    
    function CreateWorkingDirectory {
        if (-not (Test-Path $dir)) { mkdir $script:dir | Out-Null }
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
        Expand-Archive (Get-ChildItem $script:dir\*.zip | Sort-Object LastWriteTime | Select-Object -last 1) $dir
        Start-Sleep -Milliseconds 250
        $PSRFile = (Get-ChildItem $script:dir\*.mht | Sort-Object LastWriteTime | Select-Object -last 1)
        $regex = '.*[AP]M\)'
        (((Get-Content $PSRFile | select-string "^        <p><b>") -replace '^        <p><b>', '' -replace '</b>', '' -replace '</p>', '' -replace '&quot;', "'") -replace $regex | Select-String '^ User' | Select-Object -ExpandProperty Line | ForEach-Object { $_.Substring(1) }) -replace '\[.*?\]', '' -replace 'â€‹', '' -replace 'User ', 'Technician ' | Out-File "$script:dir\steps.txt" -Encoding utf8
    }

    function CleanupSteps {
        # Clean up unwanted input from Steps Recorder"$script:dir\steps.txt"
        Get-Content "$script:dir\steps.txt" | 
        Where-Object { $_ -notmatch 'mouse drag|mouse wheel|\(pane\)' } | 
        Select-Object -Unique | 
        Out-File "$script:dir\cleaned_steps.txt" -Append -Encoding utf8

        #Remove last step as it's alway irrelevant
        $PSRResult = Get-Content "$script:dir\cleaned_steps.txt"
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
$(Get-Content "$script:dir\clipboard.txt" -Raw)

Issue:
$(Get-Content "$script:dir\issue.txt")

Resolution:
$(Get-Content "$script:dir\resolution.txt")
"@ | ConvertTo-Json
        $prompt | Out-File "$script:dir\prompt.txt" -Encoding utf8
    }

    function APICall {
        $Headers = @{
            "Content-Type"  = "application/json"
            "Authorization" = "Bearer $OpenAIKey"
        }
        $Body = @{
            "model"             = "gpt-3.5-turbo-16k-0613"
            "messages"          = @( @{
                    "role"    = "system"
                    "content" = "You are a helpful IT technician that creates comprehensive ticket notes for IT support issues."
                },
                @{
                    "role"    = "system"
                    "content" = "You always respond in the first person, in the following format:\n\nReported Issue:<text describing issue>\n\nCustomer Actions Taken:<customer actions here>\n\nTroubleshooting Methods:\n- <bulletted troubleshooting steps here>\n\nResolution:<resolution here>\n\nComments & Misc. info:<miscellaneous info here>\n\nMessage to End User:\n<email to end user here>"
                },
                @{
                    "role"    = "system"
                    "content" = "Use the data in the Recorded Steps and clipped sections to include printer names, website names, program names, software version numbers, etc."
                },
                @{
                    "role"    = "system"
                    "content" = "Skip Recorded Steps that are duplicated or not relevant to the issue."
                },
                @{
                    "role"    = "system"
                    "content" = "The Troubleshooting Methods are written in the first-person by (you) the technician only. The Troubleshooting Methods section is concise as possible, while still including information such as website names, software names and printer names."
                },
                @{
                    "role"    = "system"
                    "content" = "Always start the email with 'Hi <name>' and sign off with 'Respectfully,'"
                },
                @{
                    "role"    = "system"
                    "content" = "If Recorded Steps and Clipped sections are blank, use only the Issue Description and Issue Resolution fields to complete the Troubleshooting Methods section."
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
        "Session Time: $SessionTime`n" | Out-File "$script:dir\result_header.txt" -Force -Encoding utf8
        "User Name: $env:USERDOMAIN\$env:USERNAME" | Out-File "$script:dir\result_header.txt" -Force -Encoding utf8 -Append
        "Computer Name: $env:COMPUTERNAME" | Out-File "$script:dir\result_header.txt" -Force -Encoding utf8 -Append
        "$($script:response.choices.message.content)" | Out-File "$script:dir\gpt_result.txt" -Force
    }

    function WriteResultsToHost {
        #Write final results to the shell
        Start-sleep -Milliseconds 250
        #Clear-Host
        (Get-Content "$script:dir\result_header.txt") | ForEach-Object { Write-Host $_ }
        (Get-Content "$script:dir\gpt_result.txt") | ForEach-Object { Write-Host $_ }
        Write-Host "`nToken Usage: Prompt=$($response.usage.prompt_tokens) Completion=$($response.usage.completion_tokens) Total=$($response.usage.total_tokens) Cost=`$$(($response.usage.total_tokens / 1000) * 0.002)" -ForegroundColor Gray
    }

    function Cleanup {
        Start-sleep -Milliseconds 250
        Get-Process -Name psr | Stop-Process -Force
        Get-ChildItem -path $script:dir -include "*.mht", "*.zip" -Recurse -File | Remove-Item -Recurse -Force -Confirm:$false
    }

    function GetClipboard {
        if (" " -ne $(Get-Clipboard -Raw)) { ($script:clipboard).add("$(Get-Clipboard -Raw)`n") | Out-Null }
    }
    
    EnsureUserIsNotSystem
    
    try {
        $SessionStart = Timestamp
        DisplayLogo
        Write-Host "Enter Ticket Description" -ForegroundColor DarkBlue
        $issue = Read-String

        if ($issue -ne 'r') {
            CreateWorkingDirectory
            $issue | Out-File -FilePath "$script:dir\issue.txt" -Force -Encoding utf8
        }
        else {
            $script:dir = (Get-ChildItem "C:\Windows\Temp\VTS\PSDOCS\" |
                Sort-Object Name |
                Select-Object -ExpandProperty FullName -last 1)
            $issue = Get-Content "$script:dir\issue.txt"
            $resume = "Resuming Last Session. "
        }
        if ($RecordSession -eq $true) {
            DisplayRecordingBanner
            StartStepsRecorder
            set-clipboard " "
            While ($true) {
                start-sleep -Milliseconds 250
                GetClipboard
            }
        }
    }
    finally {
        if ($RecordSession -eq $true) {
            StopStepsRecorder
            DisplayRecordingCompleteBanner
        }
        Write-Host "Enter Session Conclusion"
        $resolution = Read-String
        Add-Content -Path "$script:dir\resolution.txt" -Value $resolution -Force
        DisplayProcessingBanner
        ParseSteps
        CleanupSteps
        $script:clipboard | Select-Object -unique | Out-File -FilePath "$script:dir\clipboard.txt" -Force -Encoding utf8 -Append
        $SessionEnd = Timestamp
        CalculateSessionTime
        GeneratePrompt
        APICall
        Clear-Host
        WriteResultsToFile
        WriteResultsToHost
        Cleanup
        
        function prompt {
            "GPT>>"
        
            $ErrorActionPreference = 'Continue'
                
            function WriteResultsToHost {
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ BEGIN >>" -ForegroundColor Green
                (Get-Content "$script:dir\result_header.txt") | ForEach-Object { Write-Host $_ }
                (Get-Content "$script:dir\gpt_result.txt") | ForEach-Object { Write-Host $_ }
                Write-Host "`nToken Usage: Prompt=$($script:response.usage.prompt_tokens) Completion=$($script:response.usage.completion_tokens) Total=$($script:response.usage.total_tokens) Cost=`$$(($script:response.usage.total_tokens / 1000) * 0.002)" -ForegroundColor Gray
                Write-Host "/////////////////////////////////////////////////////////////// END >>" -ForegroundColor Red
            }
        
            While ($true) {
                Write-Host "`nType:`n's' - review recorded actions`n'c' - review copied text.`n'ctrl-c' to exit.`nOtherwise, you can ask ChatGPT to make alterations to the notes above.`n`n" -ForegroundColor Yellow
                $alterations = Read-Host "GPT-3.5-Turbo>>>"
        
                switch ($alterations) {
                    s {
                        Write-Host "\\\\\\\\ STEPS >" -ForegroundColor Green
                        Get-Content $script:dir\cleaned_steps.txt 
                    }
                    c {
                        Write-Host "\\\\\\\\ CLIPBOARD >" -ForegroundColor Green
                        Get-Content $script:dir\clipboard.txt 
                    }
                    $null {
        
                    }
                    Default {
                        if ($null -ne $script:response.choices.message.content) {
                            $ticket = $script:response.choices.message.content
                        }
                        else {
                            $ticket = $(Get-Content $script:dir\gpt_result.txt -Encoding utf8 -Raw)
                        }
        
                        $prompt = @"
        $ticket
                        
        Adjust the ticket notes above taking into account the following: 
        
        $alterations.
"@ | ConvertTo-Json
        
                        $Headers = @{
                            "Content-Type"  = "application/json"
                            "Authorization" = "Bearer $OpenAIKey"
                        }
                        $Body = @{
                            "model"             = "gpt-3.5-turbo-16k-0613"
                            "messages"          = @( @{
                                    "role"    = "system"
                                    "content" = "You are a helpful assistant that rewrites IT Support ticket notes using updated information."
                                },
                                @{
                                    "role"    = "system"
                                    "content" = "You always respond in the first person, in the following format:\n\nReported Issue:<text describing issue>\n\nCustomer Actions Taken:<customer actions here>\n\nTroubleshooting Methods:\n- <bulletted troubleshooting steps here>\n\nResolution:<resolution here>\n\nComments & Misc. info:<miscellaneous info here>\n\nMessage to End User:\n<email to end user here>"
                                },
                                @{
                                    "role"    = "system"
                                    "content" = "Only add or subtract from the ticket notes based on the information provided by the user."
                                },
                                @{
                                    "role"    = "system"
                                    "content" = "The Message to End User email is intended for the end user. All other sections are internal notes for technician review only."
                                },
                                @{
                                    "role"    = "system"
                                    "content" = "Format your response properly. Do not return messages in JSON format."
                                },
                                @{
                                    "role"    = "user"
                                    "content" = "$prompt"
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
                        "$($script:response.choices.message.content)" | Out-File "$script:dir\gpt_result.txt" -Force -Encoding utf8
                        Start-sleep -Milliseconds 250
                        WriteResultsToHost
                    }
                }
            }
        }
        prompt
    }
    prompt
}


# function passwordValidates($pass) {
#     $count = 0
 
#     if(($pass.length -ge 8) -and ($pass.length -le 32)) {
#        if($pass -match ".*\d.*") {
#           $count++
#        }
#        if($pass -match ".*[a-z].*") {
#           $count++
#        }
#        if($pass -match ".*[A-Z].*") {
#           $count++
#        }
#        if($pass -match ".*[*.!@#$%^&(){}\[\]:;'<>,.?/~`_+-=|\\].*") {
#           $count++
#        }
#     }
 
#     return $count -ge 3
#  }