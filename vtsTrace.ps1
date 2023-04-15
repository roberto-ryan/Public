<#
.DESCRIPTION
Trace-vtsSession tracks what you do to assist with ticket notes.
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
        $timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss-ff
        $dir = "$env:LOCALAPPDATA\VTS\PSDOCS\$timestamp"
            
        Read-Host "`nPress Enter to begin"
        
        $rec = @"
██████╗ ███████╗ ██████╗ ██████╗ ██████╗ ██████╗ ██╗███╗   ██╗ ██████╗
██╔══██╗██╔════╝██╔════╝██╔═══██╗██╔══██╗██╔══██╗██║████╗  ██║██╔════╝
██████╔╝█████╗  ██║     ██║   ██║██████╔╝██║  ██║██║██╔██╗ ██║██║  ███╗
██╔══██╗██╔══╝  ██║     ██║   ██║██╔══██╗██║  ██║██║██║╚██╗██║██║   ██║
██║  ██║███████╗╚██████╗╚██████╔╝██║  ██║██████╔╝██║██║ ╚████║╚██████╔╝██╗██╗██╗
╚═╝  ╚═╝╚══════╝ ╚═════╝ ╚═════╝ ╚═╝  ╚═╝╚═════╝ ╚═╝╚═╝  ╚═══╝ ╚═════╝ ╚═╝╚═╝╚═╝

"@
        Write-Host $rec -ForegroundColor Red
        Write-Host "Press Ctrl-C when finished." -ForegroundColor Yellow
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
        Write-Host "`nRecording complete...`n" -ForegroundColor Cyan
        Write-Host "Processing...`n`n" -ForegroundColor Cyan
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

        # Compile Results
        $Result = @"
RecordedSteps:
$joinedSteps
            
Keylogger:
$KeyloggerResult
"@

        $prompt = "Act as a skilled IT Support Tech. Analyze the Keylogger: and RecordedSteps: sections to write a 100 word summary of what steps were taken.
       
$Result"


        #         $prompt = "As an IT Technician, confidently provide responses using complete sentences.
        # Carefully analyze the Keylogger and Recorded Steps sections to accurately determine the technician's intended actions.
        # Be sure to avoid mentioning the use of Problem Steps Recorder, any reference to DesktopWindowXaml, and refrain from using the term 'AI',
        # Include the start and stop times in a [square bracket] at the end.
       
        # $Result"
            
        
        $body = @{
            'prompt'            = $prompt;
            'temperature'       = 0;
            'max_tokens'        = 250;
            'top_p'             = 1.0;
            'frequency_penalty' = 0.0;
            'presence_penalty'  = 0.0;
            'stop'              = @('"""');
        }
         
        $JsonBody = $body | ConvertTo-Json -Compress
        $EncodedJsonBody = [System.Text.Encoding]::UTF8.GetBytes($JsonBody)
             
        $response = Invoke-RestMethod -Uri "https://api.openai.com/v1/engines/text-davinci-003/completions" -Method Post -Body $EncodedJsonBody -Headers @{ Authorization = "Bearer $OpenAIKey" } -ContentType "application/json; charset=utf-8"
        $($response.choices.text) | Out-File "$dir\gpt_result.txt" -Force

        Start-sleep -Milliseconds 250
        #Write final results to the shell
        Write-Host "$(Get-Content "$dir\gpt_result.txt")`n`n" -ForegroundColor Yellow
        Start-sleep -Milliseconds 250
        
        #Cleanup
        Get-Process -Name psr | Stop-Process -Force
        #Remove-Item -Path "$dir\*" -recurse -Force -ErrorAction SilentlyContinue
        
    }
}