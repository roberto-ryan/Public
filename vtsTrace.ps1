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
        
        
    try {
        $SessionStart = Get-Date -Format 'h:mm:ss tt'
        $timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss-ff
        $dir = "$env:LOCALAPPDATA\VTS\PSDOCS\$timestamp"
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
    
        $issue = Read-Host "Enter a short description of the issue"
        
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
        Write-Host "Press Ctrl-C when finished.`n" -ForegroundColor Yellow
        
        While ($true) {

            # Define Funtions
            $KeyLoggerBase64 = "ZnVuY3Rpb24gU3RhcnQtS2V5TG9nZ2VyKCRQYXRoID0gIiRkaXJca2V5bG9nZ2VyLnR4dCIpIHsKICAgICMgcmVjb3JkcyBhbGwga2V5IHByZXNzZXMgdW50aWwgc2NyaXB0IGlzIGFib3J0ZWQKCiAgICAjIFNpZ25hdHVyZXMgZm9yIEFQSSBDYWxscwogICAgJHNpZ25hdHVyZXMgPSBAIgpbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIiwgQ2hhclNldD1DaGFyU2V0LkF1dG8sIEV4YWN0U3BlbGxpbmc9dHJ1ZSldIApwdWJsaWMgc3RhdGljIGV4dGVybiBzaG9ydCBHZXRBc3luY0tleVN0YXRlKGludCB2aXJ0dWFsS2V5Q29kZSk7IApbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIiwgQ2hhclNldD1DaGFyU2V0LkF1dG8pXQpwdWJsaWMgc3RhdGljIGV4dGVybiBpbnQgR2V0S2V5Ym9hcmRTdGF0ZShieXRlW10ga2V5c3RhdGUpOwpbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIiwgQ2hhclNldD1DaGFyU2V0LkF1dG8pXQpwdWJsaWMgc3RhdGljIGV4dGVybiBpbnQgTWFwVmlydHVhbEtleSh1aW50IHVDb2RlLCBpbnQgdU1hcFR5cGUpOwpbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIiwgQ2hhclNldD1DaGFyU2V0LkF1dG8pXQpwdWJsaWMgc3RhdGljIGV4dGVybiBpbnQgVG9Vbmljb2RlKHVpbnQgd1ZpcnRLZXksIHVpbnQgd1NjYW5Db2RlLCBieXRlW10gbHBrZXlzdGF0ZSwgU3lzdGVtLlRleHQuU3RyaW5nQnVpbGRlciBwd3N6QnVmZiwgaW50IGNjaEJ1ZmYsIHVpbnQgd0ZsYWdzKTsKIkAKCiMgbG9hZCBzaWduYXR1cmVzIGFuZCBtYWtlIG1lbWJlcnMgYXZhaWxhYmxlCiRBUEkgPSBBZGQtVHlwZSAtTWVtYmVyRGVmaW5pdGlvbiAkc2lnbmF0dXJlcyAtTmFtZSAnV2luMzInIC1OYW1lc3BhY2UgQVBJIC1QYXNzVGhydQoKIyBjcmVhdGUgb3V0cHV0IGZpbGUKJG51bGwgPSBOZXctSXRlbSAtUGF0aCAkUGF0aCAtSXRlbVR5cGUgRmlsZSAtRm9yY2UKCiMgQ3JlYXRlcyBsb29wIHRoYXQgZXhpdHMgd2hlbiBQU1IgaXMgbm8gbG9uZ2VyIHJ1bm5pbmcuCndoaWxlIChnZXQtcHJvY2VzcyBwc3IpIHsKICAgIFN0YXJ0LVNsZWVwIC1NaWxsaXNlY29uZHMgMTAgIzIwICM0MAoKICAgICMgc2NhbiBhbGwgQVNDSUkgY29kZXMgYWJvdmUgOAogICAgZm9yICgkYXNjaWkgPSA5OyAkYXNjaWkgLWxlIDI1NDsgJGFzY2lpKyspIHsKICAgICAgICAjIGdldCBjdXJyZW50IGtleSBzdGF0ZQogICAgICAgICRzdGF0ZSA9ICRBUEk6OkdldEFzeW5jS2V5U3RhdGUoJGFzY2lpKQoKICAgICAgICAjIGlzIGtleSBwcmVzc2VkPwogICAgICAgIGlmICgkc3RhdGUgLWVxIC0zMjc2NykgewogICAgICAgICAgICAkbnVsbCA9IFtjb25zb2xlXTo6Q2Fwc0xvY2sKCiAgICAgICAgICAgICMgdHJhbnNsYXRlIHNjYW4gY29kZSB0byByZWFsIGNvZGUKICAgICAgICAgICAgJHZpcnR1YWxLZXkgPSAkQVBJOjpNYXBWaXJ0dWFsS2V5KCRhc2NpaSwgMykKCiAgICAgICAgICAgICMgZ2V0IGtleWJvYXJkIHN0YXRlIGZvciB2aXJ0dWFsIGtleXMKICAgICAgICAgICAgJGtic3RhdGUgPSBOZXctT2JqZWN0IEJ5dGVbXSAyNTYKICAgICAgICAgICAgJGNoZWNra2JzdGF0ZSA9ICRBUEk6OkdldEtleWJvYXJkU3RhdGUoJGtic3RhdGUpCgogICAgICAgICAgICAjIHByZXBhcmUgYSBTdHJpbmdCdWlsZGVyIHRvIHJlY2VpdmUgaW5wdXQga2V5CiAgICAgICAgICAgICRteWNoYXIgPSBOZXctT2JqZWN0IC1UeXBlTmFtZSBTeXN0ZW0uVGV4dC5TdHJpbmdCdWlsZGVyCgogICAgICAgICAgICAjIHRyYW5zbGF0ZSB2aXJ0dWFsIGtleQogICAgICAgICAgICAkc3VjY2VzcyA9ICRBUEk6OlRvVW5pY29kZSgkYXNjaWksICR2aXJ0dWFsS2V5LCAka2JzdGF0ZSwgJG15Y2hhciwgJG15Y2hhci5DYXBhY2l0eSwgMCkKCiAgICAgICAgICAgIGlmICgkc3VjY2VzcykgewogICAgICAgICAgICAgICAgIyBhZGQga2V5IHRvIGxvZ2dlciBmaWxlCiAgICAgICAgICAgICAgICBbU3lzdGVtLklPLkZpbGVdOjpBcHBlbmRBbGxUZXh0KCRQYXRoLCAkbXljaGFyLCBbU3lzdGVtLlRleHQuRW5jb2RpbmddOjpVbmljb2RlKSAKICAgICAgICAgICAgfQogICAgICAgIH0KICAgIH0KfQoKfQ=="
            $Bytes = [System.Convert]::FromBase64String($KeyLoggerBase64)
            iex ( [System.Text.Encoding]::UTF8.GetString($Bytes) )
            # Create $path directory if it doesn't exist
            if (-not (Test-Path $dir)) { mkdir $dir | Out-Null }


            # Start PSR
            psr.exe /start /output "$dir\problem_steps_record-$($timestamp).zip" /gui 0 /sc 1 #/maxsc 100

            # Start Keylogger
            Start-KeyLogger
        }
    }
    finally {
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
        $SessionEnd = Get-Date -Format 'h:mm:ss tt'
        $resolution = Read-Host "If issue is resolved, write a brief description of the fix"
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
        if ($null -eq $OpenAIKey) {
            $OpenAIKey = Read-Host -Prompt "Enter OpenAI API Key" -AsSecureString
        }

        # Stop PSR
        psr.exe /stop

        Start-Sleep 3

        # Remove invalid characters from keylogger file
        $inputFile = "$dir\keylogger.txt"
        $outputFile = "$dir\keylogger.txt"

        # Read the content of the input file
        $content = Get-Content $inputFile

        # Function to determine if a character is valid UTF-8 and not a control character, excluding line breaks
        function Is-ValidUTF8AndNotControlChar($char) {
            try {
                $isControlChar = [Char]::IsControl($char)
                if ($isControlChar -and $char -ne "`r" -and $char -ne "`n") { return $false }

                [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes($char)) -eq $char
            }
            catch {
                return $false
            }
            return $true
        }

        # Filter the content to keep only valid UTF-8 characters and not control characters, preserving line breaks
        $filteredContent = $content | ForEach-Object {
            $line = $_
            -join ($line.ToCharArray() | Where-Object { Is-ValidUTF8AndNotControlChar $_ })
        }

        # Write the filtered content to the output file
        Set-Content $outputFile $filteredContent

        # Get PSR Results
        Expand-Archive (Get-ChildItem $dir\*.zip | Sort-Object LastWriteTime | Select-Object -last 1) $dir
        Start-Sleep -Milliseconds 250
        $PSRFile = (Get-ChildItem $dir\*.mht | Sort-Object LastWriteTime | Select-Object -last 1)
        $regex = '.*[AP]M\)'
        (((Get-Content $PSRFile | select-string "^        <p><b>") -replace '^        <p><b>', '' -replace '</b>', '' -replace '</p>', '' -replace '&quot;', "'") -replace $regex | Select-String '^ User' | Select-Object -ExpandProperty Line | ForEach-Object { $_.Substring(1) }) -replace '\[.*?\]', '' | Out-File "$dir\steps.txt" 
        
        $PSRResult = Get-Content "$dir\steps.txt"

        # Get Keylogger Results
        $KeyloggerResult = Get-Content "$dir\keylogger.txt"

        $StepCount = $PSRResult.Count - 2
        $steps = $PSRResult[0..$StepCount]

        # Join steps with newline characters
        $joinedSteps = $steps -join "`n"


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
        $SessionTime = Get-SessionTime -StartTime $SessionStart -EndTime $SessionEnd

        # Compile Results
        $Result = @"
Issue Description:
$issue

Issue Resolution:
$resolution

RecordedSteps:
$joinedSteps
            
Keylogger:
$KeyloggerResult
"@

        $prompt = "#EXAMPLE OUTPUT:
User Name: SB2\rober
Computer Name: SB2

Issue Reported: Screen flickering

Customer Actions Taken: None

Troubleshooting Methods:

1.	Accessed the Start menu and navigated to Settings.
2.	Selected Windows Update and clicked on Check for Updates.
3.	Closed the Settings window, right-clicked on the Start button, and chose Device Manager.
4.	Located Display Adapters and right-clicked on the NVIDIA GeForce GTX 1050, selecting Update Driver.
5.	Clicked on Search Automatically for Drivers, followed by Search for Updated Drivers on Windows Update.
6.	Closed the Settings window, right-clicked on the Microsoft Edge button, and selected New Window.
7.	Searched for 'gtx 1050 drivers' and clicked on the first result.
8.	Clicked on the Official Drivers link and selected the Download Drivers button.
9.	Navigated to the Downloads folder and double-clicked on the Name field.
10.	Updated the graphics driver, resolving the issue.

Resolution: Updating the graphics driver resolved the issue.

Additional Comments: None


Message to End User: 

[User Name],

I am pleased to inform you that we have successfully resolved the issue you were experiencing by flushing the DNS.
At your earliest convenience, please test your system to confirm that the problem has been rectified. Should you encounter any additional issues or require further assistance, do not hesitate to reach out to us.

Respectfully,
[Technician Name]
#END EXAMPLE OUTPUT

#EXAMPLE OUTPUT:
User Name: SB2\rober
Computer Name: SB2

Issue Reported: Scanner not working

Customer Actions Taken: None

Troubleshooting Methods:
1. Accessed the Start menu and navigated to Settings.
2. Selected Devices and clicked on Printers & Scanners.
3. Located the scanner in the list of devices and right-clicked on it, selecting Troubleshoot.
4. Followed the on-screen instructions to troubleshoot the scanner.
5. Unplugged the scanner from the computer and plugged it back in.
6. Reinstalled the scanner driver.

Resolution: Unfortunately, the issue remains unresolved, as the computer is unable to recognize the scanner. This ticket will now be transferred to the onsite support queue for further assistance.

Additional Comments: None

Message to End User: 

[User Name],

I regret to inform you that we have been unable to resolve the issue with your scanner, as the computer is not recognizing the device. To further investigate and address this problem, we will require an onsite technician to visit your location. Should you have any questions or concerns, please feel free to reach out.

Respectfully,
[Technician Name]
#END EXAMPLE OUTPUT

Act as IT Technician. Based on the following Keyloger and RecordedSteps sections, intrepret what the tech was trying to do while speaking in first person to fill out the #Form: sections. `
Use the EXAMPLE OUTPUT above as an example for filling out the #Form:. `
Make sure to complete each section of #Form:. `
Don't fill out the Customer Actions Taken section unless explicity told what the customer tried in the Issue Description. `
Guess what the tech was trying to accomplish to fill out the Troubleshooting Methods section step by step. `
Don't include that the Problem Steps Recorder was used. `
Don't include anything related to DesktopWindowXaml. `
Don't include the word AI. `
Skip steps that don't make logical sense. `
Only speak in complete sentences. `
Embelish the output to make the IT Technician sound very skilled, and be specific.

$Result


#Form:
User Name: $env:USERDOMAIN\$env:USERNAME
Computer Name: $env:COMPUTERNAME

Reporting Issue:
Customer Actions Taken:
Troubleshooting Methods:
Resolution:
Comments & Misc. info:
Message to End User:
"

        
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
        "Session Time: $SessionTime

$($response.choices.text)" | Out-File "$dir\gpt_result.txt" -Force

        Start-sleep -Milliseconds 250
        #Write final results to the shell
        Clear-Host
        (Get-Content "$dir\gpt_result.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        # Write-Host "$(Get-Content "$dir\gpt_result.txt")`n`n" -ForegroundColor Yellow
        Start-sleep -Milliseconds 250
        
        #Cleanup
        Get-Process -Name psr | Stop-Process -Force
        #Remove-Item -Path "$dir\*" -recurse -Force -ErrorAction SilentlyContinue
        
    }
}