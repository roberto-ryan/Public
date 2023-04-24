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
        Clear-Host
        (Get-Content "$dir\result_header.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        (Get-Content "$dir\gpt_result2.txt") | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    }
    
    $prompt = @"
$(Get-Content $dir\gpt_result.txt -Raw -Encoding utf8)`n`n`n`n Rewrite the IT ticket above taking into account the following considerations: $alterations.
"@
    
    While ($true) {
        
        $alterations = Read-Host "GPT3.5>>>"
        switch ($alterations) {
            steps { Get-Content $dir\cleaned_steps.txt }
            Default {
                $prompt = @"
        $(Get-Content $dir\gpt_result2.txt -Raw -Encoding utf8)`n`n`n`n Rewrite the IT ticket above taking into account the following considerations: $alterations.
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
                "$($response.choices.text)" | Out-File "$dir\gpt_result2.txt" -Force -Encoding utf8
                WriteResultsToHost
            }
        }
    }
}