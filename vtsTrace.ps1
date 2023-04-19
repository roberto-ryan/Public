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
        Write-Host "Press Ctrl-C when finished.`n" -ForegroundColor Yellow
    }
    
    function StartStepsRecorder {
        psr.exe /start /output "$dir\problem_steps_record-$($timestamp).zip" /gui 0 #  /sc 1 #/maxsc 100
    }
    
    function CreateWorkingDirectory {
        if (-not (Test-Path $dir)) { mkdir $dir | Out-Null }
    }
    
    #function Start-KeyLogger { THE FUNCTION IS CONTAINED WITHIN THE BASE64 ENCODED STRING, JUST CALL Start-Keylogger
    $KeyLoggerBase64 = "ZnVuY3Rpb24gU3RhcnQtS2V5TG9nZ2VyKCRQYXRoID0gIiRkaXJca2V5bG9nZ2VyLnR4dCIpIHsNCiAgICAjIHJlY29yZHMgYWxsIGtleSBwcmVzc2VzIHVudGlsIHNjcmlwdCBpcyBhYm9ydGVkDQoNCiAgICAjIFNpZ25hdHVyZXMgZm9yIEFQSSBDYWxscw0KICAgICRzaWduYXR1cmVzID0gQCINCltEbGxJbXBvcnQoInVzZXIzMi5kbGwiLCBDaGFyU2V0PUNoYXJTZXQuQXV0bywgRXhhY3RTcGVsbGluZz10cnVlKV0gDQpwdWJsaWMgc3RhdGljIGV4dGVybiBzaG9ydCBHZXRBc3luY0tleVN0YXRlKGludCB2aXJ0dWFsS2V5Q29kZSk7IA0KW0RsbEltcG9ydCgidXNlcjMyLmRsbCIsIENoYXJTZXQ9Q2hhclNldC5BdXRvKV0NCnB1YmxpYyBzdGF0aWMgZXh0ZXJuIGludCBHZXRLZXlib2FyZFN0YXRlKGJ5dGVbXSBrZXlzdGF0ZSk7DQpbRGxsSW1wb3J0KCJ1c2VyMzIuZGxsIiwgQ2hhclNldD1DaGFyU2V0LkF1dG8pXQ0KcHVibGljIHN0YXRpYyBleHRlcm4gaW50IE1hcFZpcnR1YWxLZXkodWludCB1Q29kZSwgaW50IHVNYXBUeXBlKTsNCltEbGxJbXBvcnQoInVzZXIzMi5kbGwiLCBDaGFyU2V0PUNoYXJTZXQuQXV0byldDQpwdWJsaWMgc3RhdGljIGV4dGVybiBpbnQgVG9Vbmljb2RlKHVpbnQgd1ZpcnRLZXksIHVpbnQgd1NjYW5Db2RlLCBieXRlW10gbHBrZXlzdGF0ZSwgU3lzdGVtLlRleHQuU3RyaW5nQnVpbGRlciBwd3N6QnVmZiwgaW50IGNjaEJ1ZmYsIHVpbnQgd0ZsYWdzKTsNCiJADQoNCiAgICAjIGxvYWQgc2lnbmF0dXJlcyBhbmQgbWFrZSBtZW1iZXJzIGF2YWlsYWJsZQ0KICAgICRBUEkgPSBBZGQtVHlwZSAtTWVtYmVyRGVmaW5pdGlvbiAkc2lnbmF0dXJlcyAtTmFtZSAnV2luMzInIC1OYW1lc3BhY2UgQVBJIC1QYXNzVGhydQ0KDQogICAgIyBjcmVhdGUgb3V0cHV0IGZpbGUNCiAgICAkbnVsbCA9IE5ldy1JdGVtIC1QYXRoICRQYXRoIC1JdGVtVHlwZSBGaWxlIC1Gb3JjZQ0KDQogICAgU3RhcnQtU2xlZXAgLU1pbGxpc2Vjb25kcyAxMCAjMjAgIzQwDQoNCiAgICAjIENyZWF0ZXMgbG9vcCB0aGF0IGV4aXRzIHdoZW4gUFNSIGlzIG5vIGxvbmdlciBydW5uaW5nLg0KICAgIHdoaWxlICgkdHJ1ZSkgew0KICAgICAgICAjIHNjYW4gYWxsIEFTQ0lJIGNvZGVzIGFib3ZlIDgNCiAgICAgICAgZm9yICgkYXNjaWkgPSA5OyAkYXNjaWkgLWxlIDI1NDsgJGFzY2lpKyspIHsNCiAgICAgICAgICAgICMgZ2V0IGN1cnJlbnQga2V5IHN0YXRlDQogICAgICAgICAgICAkc3RhdGUgPSAkQVBJOjpHZXRBc3luY0tleVN0YXRlKCRhc2NpaSkNCg0KICAgICAgICAgICAgIyBpcyBrZXkgcHJlc3NlZD8NCiAgICAgICAgICAgIGlmICgkc3RhdGUgLWVxIC0zMjc2Nykgew0KICAgICAgICAgICAgICAgICRudWxsID0gW2NvbnNvbGVdOjpDYXBzTG9jaw0KDQogICAgICAgICAgICAgICAgIyB0cmFuc2xhdGUgc2NhbiBjb2RlIHRvIHJlYWwgY29kZQ0KICAgICAgICAgICAgICAgICR2aXJ0dWFsS2V5ID0gJEFQSTo6TWFwVmlydHVhbEtleSgkYXNjaWksIDMpDQoNCiAgICAgICAgICAgICAgICAjIGdldCBrZXlib2FyZCBzdGF0ZSBmb3IgdmlydHVhbCBrZXlzDQogICAgICAgICAgICAgICAgJGtic3RhdGUgPSBOZXctT2JqZWN0IEJ5dGVbXSAyNTYNCiAgICAgICAgICAgICAgICAkY2hlY2trYnN0YXRlID0gJEFQSTo6R2V0S2V5Ym9hcmRTdGF0ZSgka2JzdGF0ZSkNCg0KICAgICAgICAgICAgICAgICMgcHJlcGFyZSBhIFN0cmluZ0J1aWxkZXIgdG8gcmVjZWl2ZSBpbnB1dCBrZXkNCiAgICAgICAgICAgICAgICAkbXljaGFyID0gTmV3LU9iamVjdCAtVHlwZU5hbWUgU3lzdGVtLlRleHQuU3RyaW5nQnVpbGRlcg0KDQogICAgICAgICAgICAgICAgIyB0cmFuc2xhdGUgdmlydHVhbCBrZXkNCiAgICAgICAgICAgICAgICAkc3VjY2VzcyA9ICRBUEk6OlRvVW5pY29kZSgkYXNjaWksICR2aXJ0dWFsS2V5LCAka2JzdGF0ZSwgJG15Y2hhciwgJG15Y2hhci5DYXBhY2l0eSwgMCkNCg0KICAgICAgICAgICAgICAgIGlmICgkc3VjY2Vzcykgew0KICAgICAgICAgICAgICAgICAgICAjIGFkZCBrZXkgdG8gbG9nZ2VyIGZpbGUNCiAgICAgICAgICAgICAgICAgICAgW1N5c3RlbS5JTy5GaWxlXTo6QXBwZW5kQWxsVGV4dCgkUGF0aCwgJG15Y2hhciwgW1N5c3RlbS5UZXh0LkVuY29kaW5nXTo6VW5pY29kZSkgDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgfQ0KICAgICAgICB9DQogICAgfQ0KfQ=="
    $Bytes = [System.Convert]::FromBase64String($KeyLoggerBase64)
    Invoke-Expression ( [System.Text.Encoding]::UTF8.GetString($Bytes) )
    #}

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
            $line = $_.TrimStart(' ')
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

        #Remove last step as it's alway irrelevant
        $PSRResult = Get-Content "$dir\cleaned_steps.txt"
        $StepCount = $PSRResult.Count - 2
        $steps = $PSRResult[0..$StepCount]

        # Join steps with newline characters to remove blank lines
        $script:joinedSteps = $steps -join "`n"
    }

    function RemovePasswords {
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

        $script:KeyloggerResult = Get-Content $OutputFile
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
If Keylogger and RecorderSteps sections are blank, use only the Issue Description and Issue Resolution fields to complete the #Form:. `
Make sure to complete each section of #Form:. `
Don't fill out the Customer Actions Taken section unless explicity told what the customer tried in the Issue Description. `
Make an educated guess what the tech was trying to accomplish to fill out the Troubleshooting Methods section step by step. `
Don't include that the Problem Steps Recorder was used. `
Don't include anything related to DesktopWindowXaml. `
Don't include the word AI. `
Skip steps that don't make logical sense. `
Only speak in complete sentences. `
Embelish the output to make the IT Technician sound very skilled, and be specific.

Issue Description:
$issue

Issue Resolution:
$resolution

RecordedSteps:
$joinedSteps

Keylogger:
$KeyloggerResult

#Form:
Reporting Issue:
Customer Actions Taken:
Troubleshooting Methods:
Resolution:
Comments & Misc. info:
Message to End User:

`"`"`"
"@
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
        "Session Time: $SessionTime`n" | Out-File "$dir\gpt_result.txt" -Force
        "User Name: $env:USERDOMAIN\$env:USERNAME" | Out-File "$dir\gpt_result.txt" -Force -Append
        "Computer Name: $env:COMPUTERNAME" | Out-File "$dir\gpt_result.txt" -Force -Append
        "$($response.choices.text)" | Out-File "$dir\gpt_result.txt" -Force -Append
    }

    function WriteResultsToHost {
        #Write final results to the shell
        Start-sleep -Milliseconds 250
        Clear-Host
        (Get-Content "$dir\gpt_result.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    }

    function Cleanup {
        Start-sleep -Milliseconds 250
        Get-Process -Name psr | Stop-Process -Force
        Get-ChildItem -path $dir -include "*.mht", "*.zip", "*keylogger.txt" -Recurse -File | Remove-Item -Recurse -Force -Confirm:$false
    }
    
    try {
        $SessionStart = Timestamp
        DisplayLogo
        $issue = Read-Host "Summarize the issue and steps performed by the user."
        DisplayRecordingBanner
        CreateWorkingDirectory
        StartStepsRecorder
        Start-KeyLogger
    }
    finally {
        StopStepsRecorder
        DisplayRecordingCompleteBanner
        $resolution = Read-Host "Session Conclusion"
        DisplayProcessingBanner
        ParseSteps
        CleanupSteps
        RemoveInvalidCharacters
        RemovePasswords
        $SessionEnd = Timestamp
        CalculateSessionTime
        GeneratePrompt
        APICall
        WriteResultsToFile
        WriteResultsToHost
        Cleanup
    }
}