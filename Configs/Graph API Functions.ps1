
#region **** Function Get-MyGQuery ****
Function Get-MyGQuery ()
{
  <#
    .SYNOPSIS
      Query Microsoft Graph API with simple paging support.
    .DESCRIPTION
      This function queries the Microsoft Graph API using a provided authentication token and supports basic query options such as API version, resource endpoint, and retrieving all pages of results.
      It is designed for straightforward queries where advanced filtering or selection is not required.
    .PARAMETER AuthToken
      The authentication token (as a hashtable) to use for the request. Typically obtained from an OAuth flow or authentication function.
    .PARAMETER Version
      The Graph API version to use. Accepts "Beta" or "v1.0". Default is "Beta".
    .PARAMETER Resource
      The resource endpoint to query in the Graph API (e.g., "users", "groups", "me/messages").
    .PARAMETER All
      If specified, retrieves all pages of results by following the @odata.nextLink property.
    .PARAMETER Wait
      The number of milliseconds to wait between requests when paging through results. Default is 100.
    .EXAMPLE
      Get-MyGQuery -AuthToken $AuthToken -Resource "users"
    .EXAMPLE
      Get-MyGQuery -AuthToken $AuthToken -Resource "groups" -Version "v1.0" -All
    .EXAMPLE
      Get-MyGQuery -AuthToken $AuthToken -Resource "me/messages" -Wait 200
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  Param (
    [parameter(Mandatory = $True)]
    [Hashtable]$AuthToken = $Script:Authtoken,
    [ValidateSet("Beta", "v1.0")]
    [String]$Version = "Beta",
    [parameter(Mandatory = $True)]
    [String]$Resource,
    [Switch]$All,
    [Int]$Wait = 100
  )
  Write-Verbose -Message "Enter Function Get-MyGQuery"
  
  $Uri = "https://graph.microsoft.com/$($Version)/$($Resource)"
  Do
  {
    Write-Verbose -Message "Query Graph API"
    $ReturnData = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Headers $AuthToken -Method Get -ContentType application/json -ErrorAction SilentlyContinue -Verbose:$False
    If ($ReturnData.StatusCode -eq 200)
    {
      $Content = $ReturnData.Content | ConvertFrom-Json
      If (@($Content.PSObject.Properties.match("value")).Count)
      {
        $Content.Value
      }
      Else
      {
        $Content
      }
      $Uri = ($Content."@odata.nextLink")
      Start-Sleep -Milliseconds $Wait
    }
    Else
    {
      $Uri = $Null
    }
  }
  While ((-not [String]::IsNullOrEmpty($Uri)) -and $All.IsPresent)
  
  Write-Verbose -Message "Exit Function Get-MyGQuery"
}
#endregion **** Function Get-MyGQuery ****

#region **** Function Get-MyOAuthApplicationToken ****
Function Get-MyOAuthApplicationToken ()
{
  <#
    .SYNOPSIS
      Get Application OAuth Token
    .DESCRIPTION
      Retrieves an OAuth 2.0 token for an application using client credentials flow.
      This token can be used to authenticate requests to Microsoft Graph or other Azure AD protected resources.
    .PARAMETER TenantID
      The Azure Active Directory tenant ID where the application is registered.
    .PARAMETER ClientID
      The Application (client) ID of the Azure AD app registration.
    .PARAMETER ClientSecret
      The client secret associated with the Azure AD app registration.
    .PARAMETER Scope
      The resource URI or scope for which the token is requested. Defaults to 'https://graph.microsoft.com/.default'.
    .EXAMPLE
      Get-MyOAuthApplicationToken -TenantID $TenantID -ClientID $ClientID -ClientSecret $ClientSecret
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "New")]
  Param (
    [parameter(Mandatory = $True)]
    [String]$MyTenantID,
    [parameter(Mandatory = $True)]
    [String]$MyClientID,
    [parameter(Mandatory = $True)]
    [String]$MyClientSecret,
    [String]$Scope = "https://graph.microsoft.com/.default"
  )
  Write-Verbose -Message "Enter Function Get-MyOAuthApplicationToken"
  
  $Body = @{
    "grant_type"    = "client_credentials"
    "client_id"     = $MyClientID
    "client_secret" = $MyClientSecret
    "Scope"         = $Scope
  }
  
  $Uri = "https://login.microsoftonline.com/$($MyTenantID)/oauth2/v2.0/token"
  
  Try
  {
    $AuthResult = Invoke-RestMethod -Uri $Uri -Body $Body -Method Post -ContentType "application/x-www-form-urlencoded" -ErrorAction SilentlyContinue
  }
  Catch
  {
    $AuthResult = $Null
  }
  
  If ([String]::IsNullOrEmpty($AuthResult))
  {
    # Failed to Authenticate
    @{
      "Expires_In" = 0
    }
  }
  Else
  {
    # Successful Authentication
    @{
      "Content-Type"  = "application/json"
      "Authorization" = "Bearer " + $AuthResult.Access_Token
      "Expires_In"    = $AuthResult.Expires_In
    }
  }
  
  Write-Verbose -Message "Exit Function Get-MyOAuthApplicationToken"
}
#endregion **** Function Get-MyOAuthApplicationToken ****

