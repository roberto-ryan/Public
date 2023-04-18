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
    }
    
    function DisplayRecording {
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
    }
    
    function StartStepsRecorder {
        psr.exe /start /output "$dir\problem_steps_record-$($timestamp).zip" /gui 0 #  /sc 1 #/maxsc 100
    }
    
    function CreateWorkingDirectory {
        if (-not (Test-Path $dir)) { mkdir $dir | Out-Null }
    }
    
    #function Start-KeyLogger { THE FUNCTION IS CONTAINED WITHIN THE BASE64 ENCODED STRING, JUST CALL Start-Keylogger
    $KeyLoggerBase64 = "ZnVuY3Rpb24gU3RhcnQtS2V5TG9nZ2VyKCRQYXRoID0gIiRkaXJca2V5bG9nZ2VyLnR4dCIpIHsKICAgICMgcmVjb3JkcyBhbGwga2V5IHByZXNzZXMgdW50aWwgc2NyaXB0IGlzIGFib3J0ZWQKCiAgICAjIFNpZ25hdHVyZXMgZm9yIEFQSSBDYWxscwogICAgJHNpZ25hdHVyZXMgPSBAIgpbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIiwgQ2hhclNldD1DaGFyU2V0LkF1dG8sIEV4YWN0U3BlbGxpbmc9dHJ1ZSldIApwdWJsaWMgc3RhdGljIGV4dGVybiBzaG9ydCBHZXRBc3luY0tleVN0YXRlKGludCB2aXJ0dWFsS2V5Q29kZSk7IApbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIiwgQ2hhclNldD1DaGFyU2V0LkF1dG8pXQpwdWJsaWMgc3RhdGljIGV4dGVybiBpbnQgR2V0S2V5Ym9hcmRTdGF0ZShieXRlW10ga2V5c3RhdGUpOwpbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIiwgQ2hhclNldD1DaGFyU2V0LkF1dG8pXQpwdWJsaWMgc3RhdGljIGV4dGVybiBpbnQgTWFwVmlydHVhbEtleSh1aW50IHVDb2RlLCBpbnQgdU1hcFR5cGUpOwpbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIiwgQ2hhclNldD1DaGFyU2V0LkF1dG8pXQpwdWJsaWMgc3RhdGljIGV4dGVybiBpbnQgVG9Vbmljb2RlKHVpbnQgd1ZpcnRLZXksIHVpbnQgd1NjYW5Db2RlLCBieXRlW10gbHBrZXlzdGF0ZSwgU3lzdGVtLlRleHQuU3RyaW5nQnVpbGRlciBwd3N6QnVmZiwgaW50IGNjaEJ1ZmYsIHVpbnQgd0ZsYWdzKTsKIkAKCiMgbG9hZCBzaWduYXR1cmVzIGFuZCBtYWtlIG1lbWJlcnMgYXZhaWxhYmxlCiRBUEkgPSBBZGQtVHlwZSAtTWVtYmVyRGVmaW5pdGlvbiAkc2lnbmF0dXJlcyAtTmFtZSAnV2luMzInIC1OYW1lc3BhY2UgQVBJIC1QYXNzVGhydQoKIyBjcmVhdGUgb3V0cHV0IGZpbGUKJG51bGwgPSBOZXctSXRlbSAtUGF0aCAkUGF0aCAtSXRlbVR5cGUgRmlsZSAtRm9yY2UKCiMgQ3JlYXRlcyBsb29wIHRoYXQgZXhpdHMgd2hlbiBQU1IgaXMgbm8gbG9uZ2VyIHJ1bm5pbmcuCndoaWxlIChnZXQtcHJvY2VzcyBwc3IpIHsKICAgIFN0YXJ0LVNsZWVwIC1NaWxsaXNlY29uZHMgMTAgIzIwICM0MAoKICAgICMgc2NhbiBhbGwgQVNDSUkgY29kZXMgYWJvdmUgOAogICAgZm9yICgkYXNjaWkgPSA5OyAkYXNjaWkgLWxlIDI1NDsgJGFzY2lpKyspIHsKICAgICAgICAjIGdldCBjdXJyZW50IGtleSBzdGF0ZQogICAgICAgICRzdGF0ZSA9ICRBUEk6OkdldEFzeW5jS2V5U3RhdGUoJGFzY2lpKQoKICAgICAgICAjIGlzIGtleSBwcmVzc2VkPwogICAgICAgIGlmICgkc3RhdGUgLWVxIC0zMjc2NykgewogICAgICAgICAgICAkbnVsbCA9IFtjb25zb2xlXTo6Q2Fwc0xvY2sKCiAgICAgICAgICAgICMgdHJhbnNsYXRlIHNjYW4gY29kZSB0byByZWFsIGNvZGUKICAgICAgICAgICAgJHZpcnR1YWxLZXkgPSAkQVBJOjpNYXBWaXJ0dWFsS2V5KCRhc2NpaSwgMykKCiAgICAgICAgICAgICMgZ2V0IGtleWJvYXJkIHN0YXRlIGZvciB2aXJ0dWFsIGtleXMKICAgICAgICAgICAgJGtic3RhdGUgPSBOZXctT2JqZWN0IEJ5dGVbXSAyNTYKICAgICAgICAgICAgJGNoZWNra2JzdGF0ZSA9ICRBUEk6OkdldEtleWJvYXJkU3RhdGUoJGtic3RhdGUpCgogICAgICAgICAgICAjIHByZXBhcmUgYSBTdHJpbmdCdWlsZGVyIHRvIHJlY2VpdmUgaW5wdXQga2V5CiAgICAgICAgICAgICRteWNoYXIgPSBOZXctT2JqZWN0IC1UeXBlTmFtZSBTeXN0ZW0uVGV4dC5TdHJpbmdCdWlsZGVyCgogICAgICAgICAgICAjIHRyYW5zbGF0ZSB2aXJ0dWFsIGtleQogICAgICAgICAgICAkc3VjY2VzcyA9ICRBUEk6OlRvVW5pY29kZSgkYXNjaWksICR2aXJ0dWFsS2V5LCAka2JzdGF0ZSwgJG15Y2hhciwgJG15Y2hhci5DYXBhY2l0eSwgMCkKCiAgICAgICAgICAgIGlmICgkc3VjY2VzcykgewogICAgICAgICAgICAgICAgIyBhZGQga2V5IHRvIGxvZ2dlciBmaWxlCiAgICAgICAgICAgICAgICBbU3lzdGVtLklPLkZpbGVdOjpBcHBlbmRBbGxUZXh0KCRQYXRoLCAkbXljaGFyLCBbU3lzdGVtLlRleHQuRW5jb2RpbmddOjpVbmljb2RlKSAKICAgICAgICAgICAgfQogICAgICAgIH0KICAgIH0KfQoKfQ=="
    $Bytes = [System.Convert]::FromBase64String($KeyLoggerBase64)
    Invoke-Expression ( [System.Text.Encoding]::UTF8.GetString($Bytes) )
    #}

    function DisplayRecordingComplete {
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

    function DisplayProcessing {
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
        Start-Sleep 3
    }

    function RemoveInvalidCharacters {
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
    }

    function ParseSteps {
        #Parse results
        Expand-Archive (Get-ChildItem $dir\*.zip | Sort-Object LastWriteTime | Select-Object -last 1) $dir
        Start-Sleep -Milliseconds 250
        $PSRFile = (Get-ChildItem $dir\*.mht | Sort-Object LastWriteTime | Select-Object -last 1)
        $regex = '.*[AP]M\)'
        (((Get-Content $PSRFile | select-string "^        <p><b>") -replace '^        <p><b>', '' -replace '</b>', '' -replace '</p>', '' -replace '&quot;', "'") -replace $regex | Select-String '^ User' | Select-Object -ExpandProperty Line | ForEach-Object { $_.Substring(1) }) -replace '\[.*?\]', '' | Out-File "$dir\steps.txt" 
    }

    function CleanupSteps {
        # Clean up unwanted input from Steps Recorder"$dir\steps.txt"
        Get-Content "$dir\steps.txt" | 
        Where-Object { $_ -notmatch 'mouse drag|mouse wheel|\(pane\)' } | 
        Sort-Object -Unique | 
        Set-Content "$dir\cleaned_steps.txt"
    }
    
    try {
        $SessionStart = Timestamp
        DisplayLogo
        $issue = Read-Host "Summarize the issue and steps performed by the user."
        DisplayRecording
        CreateWorkingDirectory
        StartStepsRecorder
        Start-KeyLogger
    }
    finally {
        DisplayRecordingComplete
        $SessionEnd = Timestamp
        $resolution = Read-Host "Session Conclusion"
        DisplayProcessing
        StopStepsRecorder
        RemoveInvalidCharacters
        ParseSteps
        CleanupSteps

        $PSRResult = Get-Content "$dir\cleaned_steps.txt"

        # Remove-Passwords.ps1
        $InputFile = "$dir\keylogger.txt"
        $OutputFile = "$dir\keylog-cleaned.txt"

        # Define a regex pattern to detect common password patterns
        $passwordPattern = "^(?:(?:(?=.*[0-9])(?=.*[a-z])(?=.*[A-Z]))|(?:(?=.*[a-z])(?=.*[A-Z])(?=.*[*.!@$%^&(){}[]:;<>,.?/~_+-=|\]))|(?:(?=.*[0-9])(?=.*[A-Z])(?=.*[*.!@$%^&(){}[]:;<>,.?/~_+-=|\]))|(?:(?=.*[0-9])(?=.*[a-z])(?=.*[*.!@$%^&(){}[]:;<>,.?/~_+-=|\]))).{8,32}"

        # Process input file
        Get-Content -Path $InputFile | ForEach-Object {
            $line = $_
            $cleanLine = [regex]::Replace($line, $passwordPattern, '')
            if (-not [string]::IsNullOrWhiteSpace($cleanLine)) {
                $cleanLine
            }
        } | Set-Content -Path $OutputFile


        # Get Keylogger Results
        $KeyloggerResult = Get-Content "$dir\keylog-cleaned.txt"

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

        $prompt = "Example1 = (
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

We have successfully resolved the screen flickering issue you were experiencing by updating the graphics driver.
At your earliest convenience, please test your system to confirm that the issue with your screen has been rectified. Should you encounter any additional issues or require further assistance, do not hesitate to reach out to us.

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

Act as IT Technician. Based on the following Keyloger and RecordedSteps sections, intrepret what the tech was trying to do while speaking in first person to fill out the #Form: sections. `
Use the examples above as an example when filling out the #Form:. `
Make sure to complete each section of #Form:. `
Don't fill out the Customer Actions Taken section unless explicity told what the customer tried in the Issue Description. `
Make an educated guess what the tech was trying to accomplish to fill out the Troubleshooting Methods section step by step. `
Don't include that the Problem Steps Recorder was used. `
Don't include anything related to DesktopWindowXaml. `
Don't include the word AI. `
Skip steps that don't make logical sense. `
Only speak in complete sentences. `
Embelish the output to make the IT Technician sound very skilled, and be specific.

$Result

#Form:
Reporting Issue:
Customer Actions Taken:
Troubleshooting Methods:
Resolution:
Comments & Misc. info:
Message to End User:

`"`"`"
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

User Name: $env:USERDOMAIN\$env:USERNAME
Computer Name: $env:COMPUTERNAME

$($response.choices.text)" | Out-File "$dir\gpt_result.txt" -Force

        Start-sleep -Milliseconds 250
        #Write final results to the shell
        Clear-Host
        (Get-Content "$dir\gpt_result.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        # Write-Host "$(Get-Content "$dir\gpt_result.txt")`n`n" -ForegroundColor Yellow
        Start-sleep -Milliseconds 250
        
        #Cleanup
        Get-Process -Name psr | Stop-Process -Force
        Get-ChildItem -path $dir -include "*.mht", "*.zip", "*keylogger.txt" -Recurse -File | Remove-Item -Recurse -Force -Confirm:$false
    }
}