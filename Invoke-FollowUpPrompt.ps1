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
        Write-Host "\\\\\\\\\\\" -ForegroundColor Green
        (Get-Content "$dir\result_header.txt") | ForEach-Object { Write-Host $_ }
        (Get-Content "$dir\gpt_result.txt") | ForEach-Object { Write-Host $_ }
        Write-Host "\\\\\\\\\\\" -ForegroundColor Green
    }

    While ($true) {
        Write-Host "`nType 's' to review recorded actions or 'c' to review copied text.`nOtherwise, you can ask ChatGPT to make alterations to the notes above.`n`n" -ForegroundColor Yellow
        $alterations = Read-Host "GPT3.5>>>"

        switch ($alterations) {
            s {
                Write-Host "\\\\\\\\\\\ STEPS" -ForegroundColor Green
                "`n"
                Get-Content $dir\cleaned_steps.txt 
                "`n"
                Write-Host "\\\\\\\\\\\" -ForegroundColor Green
            }
            c {
                Write-Host "\\\\\\\\\\\ CLIPBOARD" -ForegroundColor Green
                "`n"
                Get-Content $dir\clipboard.txt 
                "`n"
                Write-Host "\\\\\\\\\\\" -ForegroundColor Green
            }
            Default {
                $prompt = @"
Example1 = (
Issue Reported: Screen flickering

Customer Actions Taken: None

Troubleshooting Methods:
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

Always mainting the formatting of the example above.

Ticket Notes:
$(Get-Content $dir\gpt_result.txt -Encoding utf8)


#INPUT = (
Recorded Steps:
$($script:joinedSteps)

MISC:
$(Get-Content "$dir\clipboard.txt")

Issue:
$(Get-Content "$dir\issue.txt")

Resolution:
$(Get-Content "$dir\resolution.txt")
)

Rewrite the Ticket Notes, taking into account the following: 
    
$alterations.

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