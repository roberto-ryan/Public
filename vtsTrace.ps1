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
            $dir = "C:\temp\PSDocs"

            # Define Funtions
            function Start-KeyLogger($Path = "C:\temp\PSDocs\keylogger.txt") {
                # records all key presses until script is aborted

                # Signatures for API Calls
                $signatures = @'
[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
public static extern short GetAsyncKeyState(int virtualKeyCode); 
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int GetKeyboardState(byte[] keystate);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int MapVirtualKey(uint uCode, int uMapType);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpkeystate, System.Text.StringBuilder pwszBuff, int cchBuff, uint wFlags);
'@

                # load signatures and make members available
                $API = Add-Type -MemberDefinition $signatures -Name 'Win32' -Namespace API -PassThru
    
                # create output file
                $null = New-Item -Path $Path -ItemType File -Force

                # # Add Beginning Timestamp
                # Add-Content -Path $Path -Value "Start time: $(Get-Date)"

                # Creates loop that exits when PSR is no longer running.
                while (get-process psr) {
                    Start-Sleep -Milliseconds 10 #20 #40
      
                    # scan all ASCII codes above 8
                    for ($ascii = 9; $ascii -le 254; $ascii++) {
                        # get current key state
                        $state = $API::GetAsyncKeyState($ascii)

                        # is key pressed?
                        if ($state -eq -32767) {
                            $null = [console]::CapsLock

                            # translate scan code to real code
                            $virtualKey = $API::MapVirtualKey($ascii, 3)

                            # get keyboard state for virtual keys
                            $kbstate = New-Object Byte[] 256
                            $checkkbstate = $API::GetKeyboardState($kbstate)

                            # prepare a StringBuilder to receive input key
                            $mychar = New-Object -TypeName System.Text.StringBuilder

                            # translate virtual key
                            $success = $API::ToUnicode($ascii, $virtualKey, $kbstate, $mychar, $mychar.Capacity, 0)

                            if ($success) {
                                # add key to logger file
                                [System.IO.File]::AppendAllText($Path, $mychar, [System.Text.Encoding]::Unicode) 
                            }
                        }
                    }
                }
                # # Add Ending Timestamp
                # Add-Content -Path $Path -Value "End time: $(Get-Date)"
            }

            # Create $path directory if it doesn't exist
            if (-not (Test-Path $dir)) { mkdir $dir | Out-Null }

            $timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss-ff

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

        $dir = "C:\temp\PSDocs"

        # Stop PSR
        psr.exe /stop

        Start-Sleep 3

        # Remove invalid characters from keylogger file
        $File = "$dir\keylogger.txt"
        $Content = Get-Content -Path $File -Encoding UTF8 -Raw
        $CleanedContent = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes($Content))
        Set-Content -Path $File -Value $CleanedContent -Encoding UTF8

        # Get PSR Results
        Expand-Archive (Get-ChildItem C:\temp\PSDocs\*.zip | Sort-Object LastWriteTime | Select-Object -last 1) $dir
        Start-Sleep -Milliseconds 250
        $PSRFile = (Get-ChildItem $dir\*.mht | Sort-Object LastWriteTime | Select-Object -last 1)
        $regex = '^Step.*M\)' #'^Step \d+: \(\u200E\d{1,2}/\u200E\d{1,2}/\u200E\d{4} \d{1,2}:\d{2}:\d{2} (AM|PM)\) '
        ((Get-Content $PSRFile | select-string "^        <p><b>") -replace '^        <p><b>', '' -replace '</b>', '' -replace '</p>', '' -replace '&quot;', "'") -replace $regex | Select-String '^ User' | Select -ExpandProperty Line | Out-File "$dir\steps.txt" 
        $PSRResult = Get-Content "$dir\steps.txt"

        # Get Keylogger Results
        $KeyloggerResult = Get-Content "$dir\keylogger.txt"

        $StepCount = $PSRResult.Count - 2

        # Compile Results
        $Result = @(
            "RecordedSteps:"
            $PSRResult[0..$StepCount]
            ""
            "Keylogger:"
            $KeyloggerResult
        )

        $prompt = "
            Only return complete sentences. `
            Act as IT Technician.`
            Based on the following Keyloger and RecordedSteps sections, intrepret what the tech was trying to do while speaking in first person. `
            Don't include that the Problem Steps Recorder was used. `
            Don't include anything related to DesktopWindowXaml. `
            Don't include the word AI.
            
        
            $Result"
            
        
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
        $($response.choices.text) | Out-File "C:\temp\PSDocs\gpt_result.txt" -Force

        Start-sleep -Milliseconds 250
        Write-Host "$(Get-Content "C:\temp\PSDocs\gpt_result.txt")`n`n" -ForegroundColor Yellow
    }
}