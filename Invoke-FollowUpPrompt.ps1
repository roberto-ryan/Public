function GPTFollowUp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $OpenAIKey
    )

    $ErrorActionPreference = 'Continue'

    $script:dir = (Get-ChildItem "C:\Windows\Temp\VTS\PSDOCS\" |
        Sort-Object Name |
        Select-Object -ExpandProperty FullName -last 1)
        
    function WriteResultsToHost {
        Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\" -ForegroundColor Green
        (Get-Content "$script:dir\result_header.txt") | ForEach-Object { Write-Host $_ }
        (Get-Content "$script:dir\gpt_result.txt") | ForEach-Object { Write-Host $_ }
        Write-Host "`nToken Usage: Prompt=$($script:response.usage.prompt_tokens) Completion=$($script:response.usage.completion_tokens) Total=$($script:response.usage.total_tokens) Cost=`$$(($script:response.usage.total_tokens / 1000) * 0.002)" -ForegroundColor Gray
    }
    

    While ($true) {
        Write-Host "`nType 's' to review recorded actions or 'c' to review copied text.`nOtherwise, you can ask ChatGPT to make alterations to the notes above.`n`n" -ForegroundColor Yellow
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
                $ticket = @"
$(Get-Content $script:dir\gpt_result.txt -Encoding utf8 -Raw)
                
Rewrite the ticket notes above taking into account the following new information: 

$alterations.
"@ | ConvertTo-Json

                $Headers = @{
                    "Content-Type"  = "application/json"
                    "Authorization" = "Bearer $OpenAIKey"
                }
                $Body = @{
                    "model"             = "gpt-3.5-turbo"
                    "messages"          = @( @{
                            "role"    = "system"
                            "content" = "You are a helpful assistant that rewrites IT Support ticket notes using updated information."
                        },
                        @{
                            "role"    = "system"
                            "content" = "You always respond in the first person, in the following format:\n\nReported Issue:<text here>\n\nCustomer Actions Taken:<text here>\n\nTroubleshooting Methods:\n- <bulletted troubleshooting steps here>\n\nResolution:<text here>\n\nComments & Misc. info:<text here>\n\nMessage to End User:\n<email to end user here>"
                        },
                        @{
                            "role"    = "system"
                            "content" = "Keep the ticket notes unchanged except for changes that are requested by the user."
                        },
                        @{
                            "role"    = "user"
                            "content" = "$ticket"
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
                WriteResultsToHost
            }
        }
    }
}