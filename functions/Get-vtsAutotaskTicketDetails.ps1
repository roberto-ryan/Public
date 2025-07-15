function Get-vtsAutotaskTicketDetails {
  <#
  .SYNOPSIS
  This function retrieves the details of a specific Autotask ticket.
  
  .DESCRIPTION
  The Get-vtsAutotaskTicketDetails function makes a REST API call to Autotask's REST API to retrieve the details of a specific ticket. The function requires the ticket number, API integration code, username, and secret as parameters.
  
  .PARAMETER TicketNumber
  The number of the ticket for which details are to be retrieved.
  
  .PARAMETER ApiIntegrationCode
  The API integration code required for the API call.
  
  .PARAMETER UserName
  The username required for the API call.
  
  .PARAMETER Secret
  The secret required for the API call.
  
  .EXAMPLE
  Get-vtsAutotaskTicketDetails -TicketNumber "T20240214.0040" -ApiIntegrationCode '3SzZgr4XTFKavqT59YscgQA7!gr4XTFKI5*' -UserName 'avqT59YscgQA7!@EXAMPLE.COM' -Secret 'ZWqtUYKzPoJv0!'
  
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
        "field" = "ticketNumber"
        "op"    = "eq"
        "value" = $TicketNumber
      }
    )
  } | ConvertTo-Json

  # Make the API call to Autotask using the POST method
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Post' -Headers $headers -Body $body -ContentType "application/json"

  # Output the response
  $response | Select-Object -expand items
}

