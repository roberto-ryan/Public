function Get-vtsAutoTaskResourceNamebyID {
  <#
  .SYNOPSIS
  This function retrieves the first and last name of a resource from Autotask's REST API using the resource's ID.
  
  .DESCRIPTION
  The Get-vtsAutoTaskResourceNamebyID function makes a GET request to Autotask's REST API to retrieve the details of a resource. It requires the resource's ID, API integration code, username, and secret as parameters. The function returns the first and last name of the resource.
  
  .PARAMETER ResourceID
  The ID of the resource whose details are to be retrieved.
  
  .PARAMETER ApiIntegrationCode
  The API integration code required for the API call.
  
  .PARAMETER UserName
  The username required for the API call.
  
  .PARAMETER Secret
  The secret required for the API call.
  
  .EXAMPLE
  Get-vtsAutoTaskResourceNamebyID -ResourceID "30683060" -ApiIntegrationCode "YourApiIntegrationCode" -UserName "YourUserName" -Secret "YourSecret"
  This example shows how to call the function with all required parameters.
  
  .LINK
  AutoTask API
  #>
  param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceID,
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
        
  # Define the endpoint for resource
  $endpoint = "/Resources/$ResourceID"
        
  # Make the API call to Autotask using the GET method to get the resource
  $response = Invoke-RestMethod -Uri "$baseUri$endpoint" -Method 'Get' -Headers $headers -ContentType "application/json"
        
  # Return the first and last name of the resource
  "$($response.item.firstName) $($response.item.lastName)"
}

