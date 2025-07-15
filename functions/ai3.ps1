function ai3 {
  <#
  .SYNOPSIS
  This script uses OpenAI's GPT-4 model to generate IT support ticket notes and follow-up questions.
  
  .DESCRIPTION
  The script consists of three main functions: LineAcrossScreen, Invoke-OpenAIAPI, and Generate-Questions. 
  
  LineAcrossScreen creates a line across the console screen with a specified color. 
  
  Invoke-OpenAIAPI sends a request to OpenAI's API with a given prompt and API key, and optionally a previous response for context. It returns the AI's response.
  
  Generate-Questions uses the OpenAI API to generate a set of follow-up questions based on the provided ticket notes.
  
  The main loop of the script prompts the user to enter an issue description and ticket notes, generates follow-up questions, and finally generates the ticket notes.
  
  .PARAMETER Prompt
  The prompt to be sent to the OpenAI API.
  
  .PARAMETER OpenAIAPIKey
  The API key for OpenAI.
  
  .PARAMETER PreviousResponse
  The previous response from the AI, used for context in the next API call.
  
  .PARAMETER TicketNotes
  The ticket notes to be used as context for generating follow-up questions.
  
  .PARAMETER Color
  The color of the line to be drawn across the console screen.
  
  .EXAMPLE
  PS> ai3
  
  .LINK
  AI
  #>

  Start-Transcript | Out-Null

  if (![string]::IsNullOrEmpty($response)) {
    $continueSession = Read-Host -Prompt "Would you like to continue from where you left off? (y/n)"
    if ($continueSession -eq "n") {
      $context = ""
      $userInput = ""
      $clarifyingQuestions = ""
      $global:response = ""
    }
  }

  $KeyPath = "$env:LOCALAPPDATA\VTS\SecureKeyFile.txt"

  function Encrypt-SecureString {
    param(
      [Parameter(Mandatory = $true)]
      [string]$InputString,
      [Parameter(Mandatory = $true)]
      [string]$FilePath
    )
  
    $secureString = ConvertTo-SecureString -String $InputString -AsPlainText -Force
    $secureString | Export-Clixml -Path $FilePath
  }

  function Decrypt-SecureString {
    param(
      [Parameter(Mandatory = $true)]
      [string]$FilePath
    )

    $secureString = Import-Clixml -Path $FilePath
    $decryptedString = [System.Net.NetworkCredential]::new("", $secureString).Password
    return $decryptedString
  }
    
  if (Test-Path $KeyPath) {
    $OpenAIAPIKey = Decrypt-SecureString -FilePath $KeyPath
  }
    
  if ([string]::IsNullOrEmpty($OpenAIAPIKey)) {
    $OpenAIAPIKey = Read-Host -Prompt "Please enter your OpenAI API Key" -AsSecureString
    $OpenAIAPIKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($OpenAIAPIKey))
    $saveKey = Read-Host -Prompt "Would you like to save the API key for future use? (y/n)"
    if ($saveKey -eq "y") {
      $KeyDirectory = Split-Path -Path $KeyPath -Parent
      # Check if the directory exists, if not, create it
      if (!(Test-Path -Path $KeyDirectory)) {
        New-Item -ItemType Directory -Path $KeyDirectory | Out-Null
      }

      Encrypt-SecureString -InputString $OpenAIAPIKey -FilePath $KeyPath

      Write-Host "Your API key has been saved securely."
    }
  }

  function LineAcrossScreen {
    param (
      [Parameter(Mandatory = $false)]
      [string]$Color = "Green"
    )
    $script:windowWidth = (Get-Host).UI.RawUI.WindowSize.Width
    Write-Host ('-' * $windowWidth) -ForegroundColor $Color
  }

  function Invoke-OpenAIAPI {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Prompt,
      [Parameter(Mandatory = $true)]
      [string]$OpenAIAPIKey,
      [Parameter(Mandatory = $false)]
      [string]$PreviousResponse
    )

    $Headers = @{
      "Content-Type"  = "application/json"
      "Authorization" = "Bearer $OpenAIAPIKey"
    }
    $Body = @{
      "model"             = "gpt-4-0125-preview"
      "messages"          = @(
        @{
          "role"    = "system"
          "content" = "Act as a helpful IT technician to create detailed ticket notes for IT support issues."
        },
        @{
          "role"    = "system"
          "content" = "Always use the plural first person (e.g., 'we') and refer to items impersonally (e.g., 'the computer'). ALWAYS structure responses in the following format:\n\nComputer Name: <detect the computername from notes entered.>\n\nReported Issue:<text describing issue>\n\nTroubleshooting Methods:\n- <bulletted troubleshooting steps here>\n\nResolution:<resolution here>\n\nComments & Misc. info:<miscellaneous info here>\n\nMessage to End User:\n<email to end user here using non-technical, straight-forward, common wording>"
        },
        @{
          "role"    = "system"
          "content" = "Include all troubleshooting steps in the 'Troubleshooting Methods' section. Don't exclude ANY details."
        },
        @{
          "role"    = "system"
          "content" = "Imitate the following writing examples while avoiding using adjectives and uncommon words: Hi Janine,\n\nI understand this has been such a turbulent issue for Dr. Harris. Our ability to support personal non-windows devices is limited as our remote team does not have a way to access this device to provide immediate assistance.\n\nI am working to have an on-site technician deployed to your location to have this resolved as quickly as possible. Please let me know if you have any questions or concerns.\n\nRespectfully,\n\n\nHi Summer,\n\nI have modified the policy to remove gaming, now instead of just running on every computer, it will run for every user. This should catch any stragglers that may have still been out there.\n\nI am going to let the policy deploy across workstations today and through the weekend and check in on Monday to see if any computers still have the Solitaire application.\n\nRespectfully,"
        },
        @{
          "role"    = "user"
          "content" = "Here's an example of the output I want:\n\nComputer Name: SD-PC20\n\nIssue Reported: Screen flickering\n\nTroubleshooting Methods:\n- Checked for Windows Updates.\n- Navigated to the Device Manager, located Display Adapters and right-clicked on the NVIDIA GeForce GTX 1050, selecting Update Driver.\n- Clicked on Search Automatically for Drivers, followed by Search for Updated Drivers on Windows Update.\n- Searched for 'gtx 1050 drivers' and clicked on the first result.\n- Clicked on the Official Drivers link and downloaded the driver.\n- Updated the graphics driver, resolving the issue.\n\nResolution: Updating the graphics driver resolved the issue.\n\nAdditional Comments: None\n\n\nMessage to End User: \n\n[User Name],\n\nWe have successfully resolved the screen flickering issue you were experiencing by updating the graphics driver. At your earliest convenience, please test your system to confirm that the issue with your screen has been rectified. Should you encounter any additional issues or require further assistance, do not hesitate to reach out to us.\n\nRespectfully,"
        },
        @{
          "role"    = "user"
          "content" = "Document only the steps explicitly stated. Ensure accuracy and quality of the ticket notes, and make them sound good."
        },
        @{
          "role"    = "user"
          "content" = "Refer to previous responses for context when updating ticket notes: $PreviousResponse"
        },
        @{
          "role"    = "user"
          "content" = "Update the ticket notes, taking the following into account: $prompt"
        },
        @{
          "role"    = "assistant"
          "content" = ""
        }
      )
      "temperature"       = 0
      'top_p'             = 1.0
      'frequency_penalty' = 0.0
      'presence_penalty'  = 0.0
      'stop'              = @('"""')
    } | ConvertTo-Json
    
    try {
        
      $global:response = (Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Method Post -Body $Body -Headers $Headers).choices.message.content
      Return $global:response
    }
    catch {
      Write-Error "$($_.Exception.Message)"
    }
  }

  function Generate-Questions {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Prompt,
      [Parameter(Mandatory = $true)]
      [string]$OpenAIAPIKey,
      [Parameter(Mandatory = $false)]
      [string]$TicketNotes
    )

    $Headers = @{
      "Content-Type"  = "application/json"
      "Authorization" = "Bearer $OpenAIAPIKey"
    }
    $Body = @{
      "model"             = "gpt-4-0125-preview"
      "messages"          = @( @{
          "role"    = "system"
          "content" = "You are a helpful IT technician assistant that helps generate followup questions, and outputs them one per line."
        },
        @{
          "role"    = "user"
          "content" = "Ticket Notes: $TicketNotes"
        },
        @{
          "role"    = "user"
          "content" = "$prompt"
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
        
      $global:response = (Invoke-RestMethod -Uri "https://api.openai.com/v1/chat/completions" -Method Post -Body $Body -Headers $Headers).choices.message.content
      Return $global:response
    }
    catch {
      Write-Error "$($_.Exception.Message)"
    }
  }

  # Main loop
  $context = "$($global:response)"

  if ( ($null -eq $context) -or ("" -eq $context) ) { $context = Read-Host "`nEnter an issue description`n" ; LineAcrossScreen }
  
  Write-Host "`nEnter ticket note (or 'done' to finish or 'skip' to skip follow up questions)`n`n"
  $userInput = ""
  while ($userInput.ToLower() -ne "done" -and $userInput.ToLower() -ne "skip") {
    $userInput = Read-Host
    LineAcrossScreen
    ""
    if ($userInput.ToLower() -eq "done") {
      break
    }
    if ($userInput.ToLower() -eq "skip") {
      $skip = $true
      break
    }

    $context += "Human: $userInput`n"
  }

  if (-not($skip)) {
    Write-Host "Generating follow up questions..."
  
    LineAcrossScreen -Color Yellow
  
    # Generate clarifying questions
    $clarifyingQuestions = (Generate-Questions -prompt "Please generate a set of clarifying questions for the IT technician to answer. These questions should be based on the ticket notes and aim to address any missing information or unclear details in the original notes: $context" -OpenAIAPIKey $OpenAIAPIKey -TicketNotes $context) -split "`n"
  
    foreach ($question in $clarifyingQuestions) {
      if (![string]::IsNullOrEmpty($question)) {
        $answer = Read-Host -prompt "`n$question`n`n"
        if (![string]::IsNullOrEmpty($answer)) {
          $context += "assistant: $question`nHuman: $answer`n"
        }
        LineAcrossScreen -Color Yellow
      }
    }
  }

  Write-Host "Generating Ticket Notes..."

  LineAcrossScreen

  $response = Invoke-OpenAIAPI -prompt $context -OpenAIAPIKey $OpenAIAPIKey -PreviousResponse $context
  Write-Host ""
  Write-Host "$response"
  Write-Host ""

  LineAcrossScreen

  Stop-Transcript
}

