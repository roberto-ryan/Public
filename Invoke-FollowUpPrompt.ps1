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
        Write-Host "#################### START ####################" -ForegroundColor Green
        (Get-Content "$dir\result_header.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        (Get-Content "$dir\gpt_result.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        Write-Host "#################### BREAK ####################" -ForegroundColor Green
    }

    While ($true) {
        $alterations = Read-Host "GPT3.5>>>"
        
        switch ($alterations) {
            steps {
                "`n"
                Get-Content $dir\cleaned_steps.txt 
                "`n"
            }
            clipboard {
                "`n"
                Get-Content $dir\clipboard.txt 
                "`n"
            }
            Default {
                $prompt = @"
$(Get-Content $dir\gpt_result.txt -Encoding utf8)


Rewrite the ticket notes taking into account the following: 
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