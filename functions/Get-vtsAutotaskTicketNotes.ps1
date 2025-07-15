function Get-vtsAutotaskTicketNotes {
  <#
  .SYNOPSIS
  This function retrieves all notes of a specific Autotask ticket using the ticket number.
  
  .DESCRIPTION
  The Get-vtsAutotaskTicketNotes function makes a REST API call to Autotask's REST API to retrieve the ticket ID of a specific ticket using the ticket number, and then retrieves all notes of that ticket. The function requires the ticket number, API integration code, username, and secret as parameters.
  
  .PARAMETER TicketNumber
  The number of the ticket for which notes are to be retrieved.
  
  .PARAMETER ApiIntegrationCode
  The API integration code required for the API call.
  
  .PARAMETER UserName
  The username required for the API call.
  
  .PARAMETER Secret
  The secret required for the API call.
  
  .EXAMPLE
  Get-vtsAutotaskTicketNotes -TicketNumber "T20240214.0040" -ApiIntegrationCode 'tOi5y2bp7j8U7=T59YscgQA7!gr4XTFKI5*' -UserName 'avqTtOi5y2bp7j8U7=gQA7!@EXAMPLE.COM' -Secret 'ZWqtUYKztOi5y2bp7j8U7=oJv0!'
  
  .LINK
  AutoTask API
  #>
  param (
    [Parameter(Mandatory = $true)]
    [string]$TicketNumber,
    [Parameter(Mandatory = $true)]
    [string]$ApiIntegrationCode,
    [Parameter(Mandatory = $true)]
    [string]$UserName,
    [Parameter(Mandatory = $true)]
    [string]$Secret
  )

  # Define the base URI for Autotask's REST API
  $baseUri = "https://webservices14.autotask.net/ATServicesRest/V1.0"

  # Set the necessary headers for the API call
  $headers = @{
    "ApiIntegrationCode" = $ApiIntegrationCode
    "UserName"           = $UserName
    "Secret"             = $Secret
  }

  # Define the endpoint for retrieving ticket details
  $endpoint = "/Tickets/query"

  # Define the body for the API call
  $body = @{
    "Filter" = @(
      @{
        "field" = "TicketNumber"
        "op"    = "eq"
        "value" = $TicketNumber
      }
    )
  } | ConvertTo-Json

  # Make the API call to Autotask using the POST method to get the ticket ID
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Post' -Headers $headers -Body $body -ContentType "application/json"

  # Get the ID of the ticket
  $TicketId = $response.items.id

  # Define the endpoint for retrieving ticket notes
  $endpoint = "/Tickets/$TicketId/notes"

  # Make the API call to Autotask using the GET method to get the ticket notes
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Get' -Headers $headers -ContentType "application/json"

  # Output the response
  $response | Select-Object -expand items
}

