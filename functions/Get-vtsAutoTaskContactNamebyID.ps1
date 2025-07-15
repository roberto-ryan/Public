function Get-vtsAutoTaskContactNamebyID {
  <#
  .SYNOPSIS
  This function retrieves the first and last name of a contact from Autotask's REST API using the contact's ID.
  
  .DESCRIPTION
  The Get-vtsAutoTaskContactNamebyID function makes a GET request to Autotask's REST API to retrieve the details of a contact. It requires the contact's ID, API integration code, username, and secret as parameters. The function returns the first and last name of the contact.
  
  .PARAMETER ContactID
  The ID of the contact whose details are to be retrieved.
  
  .PARAMETER ApiIntegrationCode
  The API integration code required for the API call.
  
  .PARAMETER UserName
  The username required for the API call.
  
  .PARAMETER Secret
  The secret required for the API call.
  
  .EXAMPLE
  Get-vtsAutoTaskContactNamebyID -ContactID "30683060" -ApiIntegrationCode "YourApiIntegrationCode" -UserName "YourUserName" -Secret "YourSecret"
  This example shows how to call the function with all required parameters.
  
  .LINK
  AutoTask API
  #>
  param(
    [Parameter(Mandatory = $true)]
    [string]$ContactID,
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
        
  # Define the endpoint for contact
  $endpoint = "/Contacts/$ContactID"
        
  # Make the API call to Autotask using the GET method to get the contact
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Get' -Headers $headers -ContentType "application/json"
        
  # Return the first and last name of the contact
  "$($response.item.firstName) $($response.item.lastName)"
}

