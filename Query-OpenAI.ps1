function Query-OpenAI {

$file = Read-Host "Enter filepath (example: C:\temp\errors.csv)"
$question = Read-Host "GPT4>>"


$Headers = @{
    "Content-Type"  = "application/json"
    "Authorization" = "Bearer $OpenAIKey"
}
$Body = @{
    "model"             = "gpt-4"
    "messages"          = @(
        @{
            "role"    = "user"
            "content" = $(Get-Content "$file")
        },
        @{
            "role"    = "user"
            "content" = $question
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
    Write-Host $response.choices.Message.Content
}
catch {
    Write-Error "$($_.Exception.Message)"
}

$messages = @(
    @{
        "role"    = "user"
        "content" = $(Get-Content "$file")
    },
    @{
        "role"    = "user"
        "content" = $question
    },
    @{
        "role"    = "assistant"
        "content" = $response.choices.Message.Content
    })

while ($true) {
    $question2 = Read-Host "GPT4>>"

    $messages += @(@{
            "role"    = "user"
            "content" = $question2
        },
        @{
            "role"    = "assistant"
            "content" = ""
        })

    $Headers = @{
        "Content-Type"  = "application/json"
        "Authorization" = "Bearer $OpenAIKey"
    }
    $Body = @{
        "model"             = "gpt-4"
        "messages"          = $messages
        "temperature"       = 0
        'top_p'             = 1.0
        'frequency_penalty' = 0.0
        'presence_penalty'  = 0.0
        'stop'              = @('"""')
    } | ConvertTo-Json
        
    try {
            
        $script:response = Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Method Post -Body $Body -Headers $Headers
        Write-Host $response.choices.Message.Content
    }
    catch {
        Write-Error "$($_.Exception.Message)"
    }
}
}