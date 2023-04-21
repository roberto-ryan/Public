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
    
    $alterations = Read-Host "GPT3.5>>>"

    $prompt = ("$(Get-Content $dir\gpt_result.txt )`n`n`n`n Rewrite the IT ticket above taking into account the following considerations: $alterations.")

    While ($true) {

    
        $body = @{
            'prompt'            = $prompt;
            'temperature'       = 0;
            'max_tokens'        = 250;
            'top_p'             = 1.0;
            'frequency_penalty' = 0.0;
            'presence_penalty'  = 0.0;
            'stop'              = @('"""');
        }
        
        $response = Invoke-RestMethod -Uri "https://api.openai.com/v1/engines/text-davinci-003/completions" -Method Post -Body ($body | ConvertTo-Json) -Headers @{ Authorization = "Bearer $OpenAIKey" } -ContentType "application/json"
        $($response.choices.text)
        Write-Host "
        
///////" -ForegroundColor Green
        $prompt = ("Prompt:" + $prompt + " " + "Response:" + $response.choices.text)
        $prompt = ($prompt + " " + $(Read-Host "GPT3.5>>>"))
    }
}