
#region function Get-UserToken
function Get-UserToken ()
{
  <#
    .SYNOPSIS
      Get Users GraphAPI AuthToken
    .DESCRIPTION
      Get Users GraphAPI AuthToken
    .PARAMETER ClientID
    .EXAMPLE
      Get-UserToken
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  Param (
    [String]$ClientID
  )
  Write-Verbose -Message "Enter Function Get-UserToken"

  $MsResponse = Get-MSALToken -Interactive -ClientId $ClientID -RedirectUri "urn:ietf:wg:oauth:2.0:oob" -Authority "https://login.microsoftonline.com/common" -Scopes @("https://graph.microsoft.com/.default") -ExtraQueryParameters @{claims = '{"access_token" : {"amr": { "values": ["mfa"] }}}' }

  @{
    "Content-Type"  = "application/json"
    "Authorization" = "Bearer $($MsResponse.AccessToken)"
    "ExpiresOn"     = ($MsResponse.ExpiresOn.LocalDateTime.ToString())
  }

  Write-Verbose -Message "Exit Function Get-UserToken"
}
#endregion function Get-UserToken

#region function Refresh-UserToken
function Refresh-UserToken ()
{
  <#
    .SYNOPSIS
      Refresh Users GraphAPI AuthToken
    .DESCRIPTION
      Refresh Users GraphAPI AuthToken
    .PARAMETER ClientID
    .EXAMPLE
      Refresh-UserToken
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [String]$ClientID
  )
  Write-Verbose -Message "Enter Function Refresh-UserToken"

  $MsResponse = Get-MSALToken -ForceRefresh -ClientId $ClientID -RedirectUri "urn:ietf:wg:oauth:2.0:oob" -Authority "https://login.microsoftonline.com/common" -Scopes @("https://graph.microsoft.com/.default")

  @{
    "Content-Type"  = "application/json"
    "Authorization" = "Bearer $($MsResponse.AccessToken)"
    "ExpiresOn"     = ($MsResponse.ExpiresOn.LocalDateTime.ToString())
  }

  Write-Verbose -Message "Exit Function Refresh-UserToken"
}
#endregion function Refresh-UserToken

#region function Get-MyOAuthApplicationToken
function Get-MyOAuthApplicationToken
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
  param (
    [parameter(Mandatory = $True)]
    [String]$TenantID,
    [parameter(Mandatory = $True)]
    [String]$ClientID,
    [parameter(Mandatory = $True)]
    [String]$ClientSecret,
    [String]$Scope = "https://graph.microsoft.com/.default"
  )
  Write-Verbose -Message "Enter Function Get-MyOAuthApplicationToken"

  $Body = @{
    "grant_type"    = "client_credentials"
    "client_id"     = $ClientID
    "client_secret" = $ClientSecret
    "Scope"         = $Scope
  }

  $Uri = "https://login.microsoftonline.com/$($TenantID)/oauth2/v2.0/token"

  try
  {
    $AuthResult = Invoke-RestMethod -Uri $Uri -Body $Body -Method Post -ContentType "application/x-www-form-urlencoded" -ErrorAction SilentlyContinue
  }
  catch
  {
    $AuthResult = $Null
  }

  if ([String]::IsNullOrEmpty($AuthResult))
  {
    # Failed to Authenticate
    @{
      "Expires_In" = 0
    }
  }
  else
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
#endregion function Get-MyOAuthApplicationToken

#region function Get-MyGraphQuery
function Get-MyGraphQuery
{
  <#
    .SYNOPSIS
      Query Microsoft Graph API with advanced filtering and selection options.
    .DESCRIPTION
      This function queries the Microsoft Graph API using a provided authentication token and supports advanced query options such as filtering, selecting specific properties,
      ordering, searching, pagination, and retrieving all pages of results.
    .PARAMETER AuthToken
      The authentication token (as a hashtable) to use for the request. Typically obtained from an OAuth flow or authentication function.
    .PARAMETER Version
      The Graph API version to use. Accepts "Beta" or "v1.0". Default is "Beta".
    .PARAMETER Resource
      The resource endpoint to query in the Graph API (e.g., "users", "groups", "me/messages").
    .PARAMETER Count
      If specified, includes a count of the total matching resources in the response.
    .PARAMETER Filter
      An OData filter string to restrict the results (e.g., "startswith(displayName,'A')").
    .PARAMETER Expand
      An OData expand string to include related entities in the response.
    .PARAMETER Select
      An array of property names to select in the response (e.g., "displayName", "mail").
    .PARAMETER Search
      A search string to perform a full-text search on the resource.
    .PARAMETER OrderBy
      An array of property names to order the results by (e.g., "displayName desc").
    .PARAMETER Top
      The maximum number of items to return per page (between 1 and 1000). Default is 500.
    .PARAMETER Skip
      The number of items to skip before returning results (for pagination).
    .PARAMETER All
      If specified, retrieves all pages of results by following the @odata.nextLink property.
    .EXAMPLE
      Get-MyGraphQuery -AuthToken $AuthToken -Resource "users" -Select "displayName","mail" -Top 100 -All
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [Hashtable]$AuthToken = $Script:Authtoken,
    [ValidateSet("Beta", "v1.0")]
    [String]$Version = "Beta",
    [parameter(Mandatory = $True)]
    [String]$Resource,
    [Switch]$Count,
    [String]$Filter,
    [String]$Expand,
    [String[]]$Select,
    [String]$Search,
    [String[]]$OrderBy,
    [ValidateRange(1, 1000)]
    [Int]$Top = 500,
    [Int]$Skip,
    [Switch]$All
  )
  Write-Verbose -Message "Enter Function Get-MyGraphQuery"

  $MyFilters = [System.Collections.ArrayList]::New()

  #region Build Graph Query Search Filter

  if ($Count.IsPresent)
  {
    [Void]$MyFilters.Add("`$count=true")
  }

  if ($PSBoundParameters.ContainsKey("Search"))
  {
    [Void]$MyFilters.Add("`$search=`"$($Search)`"")
  }

  if ($PSBoundParameters.ContainsKey("Select"))
  {
    [Void]$MyFilters.Add("`$select=$(($Select -join ","))")
  }

  if ($PSBoundParameters.ContainsKey("OrderBy"))
  {
    [Void]$MyFilters.Add("`$orderby=$(($OrderBy -join ","))")
  }

  if ($PSBoundParameters.ContainsKey("Top"))
  {
    [Void]$MyFilters.Add("`$top=$($Top)")
  }

  if ($PSBoundParameters.ContainsKey("Skip"))
  {
    [Void]$MyFilters.Add("`$skip=$($Skip)")
  }

  if ($PSBoundParameters.ContainsKey("Filter"))
  {
    [Void]$MyFilters.Add("`$filter=$($Filter)")
  }

  if ($PSBoundParameters.ContainsKey("Expand"))
  {
    [Void]$MyFilters.Add("`$expand=$($Expand)")
  }
  #endregion Build Graph Query Search Filter

  if ($MyFilters.Count)
  {
    $Uri = "https://graph.microsoft.com/$($Version)/$($Resource)?$(($MyFilters -join "&"))"
  }
  else
  {
    $Uri = "https://graph.microsoft.com/$($Version)/$($Resource)"
  }

  do
  {
    Write-Verbose -Message "Query Graph API"
    $ReturnData = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Headers $AuthToken -Method Get -Verbose:$False
    if ($ReturnData.StatusCode -eq 200)
    {
      $Content = $ReturnData.Content | ConvertFrom-Json
      if (@($Content.PSObject.Properties.match("value")).Count)
      {
        $Content.Value
      }
      else
      {
        $Content
      }
      $Uri = ($Content."@odata.nextLink")
    }
    else
    {
      break
    }
  }
  while ((-not [String]::IsNullOrEmpty($Uri)) -and $All.IsPresent)

  Write-Verbose -Message "Exit Function Get-MyGraphQuery"
}
#endregion function Get-MyGraphQuery

#region function Get-MyGQuery
function Get-MyGQuery
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
  param (
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
  do
  {
    Write-Verbose -Message "Query Graph API"
    $ReturnData = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Headers $AuthToken -Method Get -ContentType application/json -ErrorAction SilentlyContinue -Verbose:$False
    if ($ReturnData.StatusCode -eq 200)
    {
      $Content = $ReturnData.Content | ConvertFrom-Json
      if (@($Content.PSObject.Properties.match("value")).Count)
      {
        $Content.Value
      }
      else
      {
        $Content
      }
      $Uri = ($Content."@odata.nextLink")
      Start-Sleep -Milliseconds $Wait
    }
    else
    {
      $Uri = $Null
    }
  }
  while ((-not [String]::IsNullOrEmpty($Uri)) -and $All.IsPresent)

  Write-Verbose -Message "Exit Function Get-MyGQuery"
}
#endregion function Get-MyGQuery

#region function Send-MyGraphMail
function Send-MyGraphMail
{
  <#
    .SYNOPSIS
      Sends an email message using the Microsoft Graph API.
    .DESCRIPTION
      This function sends an email via Microsoft Graph API, supporting advanced options such as specifying recipients, CC, BCC, reply-to, sender, importance, flagging, delivery/read receipts, mentions, attachments, and saving to sent items.
    .PARAMETER Version
      Specifies the Microsoft Graph API version to use. Accepts "v1.0" or "Beta". Default is "Beta".
    .PARAMETER AuthToken
      The authentication token (as a hashtable) to use for the request. Typically obtained from an OAuth flow or authentication function.
    .PARAMETER UsedID
      The User ID of the mailbox to send mail from. Required when sending as a specific user.
    .PARAMETER Subject
      The subject of the email message.
    .PARAMETER Body
      The body content of the email message.
    .PARAMETER AsText
      If specified, sends the body as plain text. Otherwise, sends as HTML.
    .PARAMETER To
      An array of recipient email addresses (System.Net.Mail.MailAddress) to send the message to.
    .PARAMETER Mention
      An array of email addresses to mention in the message. These will be added to the message's mentions collection.
    .PARAMETER CC
      An array of email addresses to include as CC recipients.
    .PARAMETER BCC
      An array of email addresses to include as BCC recipients.
    .PARAMETER ReplyTo
      An array of email addresses to include as reply-to addresses.
    .PARAMETER From
      The sender's email address (System.Net.Mail.MailAddress). If omitted, the authenticated user is used.
    .PARAMETER Importance
      Sets the importance of the message. Accepts "Low", "Normal", or "High". Default is "Normal".
    .PARAMETER Flagged
      If specified, flags the message for follow-up.
    .PARAMETER DeliveryReceipt
      If specified, requests a delivery receipt for the message.
    .PARAMETER ReadReceipt
      If specified, requests a read receipt for the message.
    .PARAMETER Attachments
      An array of file paths to attach to the message. Files are encoded as base64 and sent as file attachments.
    .PARAMETER SaveToSent
      If specified, saves the sent message to the Sent Items folder.
    .EXAMPLE
      Send-MyGraphMail -Version "v1.0" -AuthToken $AuthToken -Subject "Test" -Body "Hello World" -To $To -CC $CC -Attachments @("C:\file.txt") -Importance "High" -SaveToSent
    .NOTES
      Original Function By Ken Sweet
      Provides advanced email sending capabilities via Microsoft Graph API.
  #>
  [CmdletBinding(DefaultParameterSetName = "Me")]
  param (
    [ValidateSet("v1.0", "Beta")]
    [String]$Version = "Beta",
    [parameter(Mandatory = $True)]
    [HashTable]$AuthToken,
    [parameter(Mandatory = $True, ParameterSetName = "User")]
    [String]$UsedID,
    [parameter(Mandatory = $True)]
    [String]$Subject,
    [parameter(Mandatory = $True)]
    [String]$Body,
    [Switch]$AsText,
    [parameter(Mandatory = $True)]
    [System.Net.Mail.MailAddress[]]$To,
    [System.Net.Mail.MailAddress[]]$Mention,
    [System.Net.Mail.MailAddress[]]$CC,
    [System.Net.Mail.MailAddress[]]$BCC,
    [System.Net.Mail.MailAddress[]]$ReplyTo,
    [System.Net.Mail.MailAddress]$From,
    [ValidateSet("Low", "Normal", "High")]
    [String]$Importance = "Normal",
    [Switch]$Flagged,
    [Switch]$DeliveryReceipt,
    [Switch]$ReadReceipt,
    [String[]]$Attachments,
    [Switch]$SaveToSent
  )
  Write-Verbose -Message "Enter Function Send-MyGraphMail"

  $Message = [Ordered]@{ "Message" = [Ordered]@{ "Subject" = $Subject } }

  if ($AsText.IsPresent)
  {
    [Void]$Message.Message.Add("Body", [Ordered]@{ "ContentType" = "TEXT" })
  }
  else
  {
    [Void]$Message.Message.Add("Body", [Ordered]@{ "ContentType" = "HTML" })
  }

  [Void]$Message.Message.Body.Add("Content", $Body)
  [Void]$Message.Message.Add("Importance", $Importance)
  [Void]$Message.Message.Add("isDeliveryReceiptRequested", ($DeliveryReceipt.IsPresent.ToString()))
  [Void]$Message.Message.Add("isReadReceiptRequested", ($ReadReceipt.IsPresent.ToString()))

  if ($Flagged.IsPresent)
  {
    [Void]$Message.Message.Add("flag", [Ordered]@{ "flagStatus" = "flagged" })
    [Void]$Message.Message.Flag.Add("dueDateTime", [Ordered]@{ "dateTime" = ([DateTime]::Now.ToString("yyyy-MM-ddT23:59:59")); "timeZone" = ([TimeZone]::CurrentTimeZone.StandardName) })
    [Void]$Message.Message.Flag.Add("startDateTime", [Ordered]@{ "dateTime" = ([DateTime]::Now.ToString("yyyy-MM-ddT23:59:59")); "timeZone" = ([TimeZone]::CurrentTimeZone.StandardName) })
  }

  if ($PSBoundParameters.ContainsKey("From"))
  {
    [Void]$Message.Message.Add("from", [Ordered]@{ "emailAddress" = @{ "address" = ($From.Address) } })
  }

  [Void]$Message.Message.Add("toRecipients", ([System.Collections.ArrayList]::New()))
  $TO | ForEach-Object -Process { [Void]$Message.Message.toRecipients.Add(@{ "emailAddress" = @{ "address" = ($PSItem.Address) } }) }

  if ($PSBoundParameters.ContainsKey("$Mention"))
  {
    [Void]$Message.Message.Add("Mentions", ([System.Collections.ArrayList]::New()))
    $Mention | ForEach-Object -Process { [Void]$Message.Message.Mentions.Add(@{ "Mentioned" = @{ "name" = ($PSItem.Address); "address" = ($PSItem.Address) } }) }
  }

  if ($PSBoundParameters.ContainsKey("CC"))
  {
    [Void]$Message.Message.Add("ccRecipients", ([System.Collections.ArrayList]::New()))
    $CC | ForEach-Object -Process { [Void]$Message.Message.ccRecipients.Add(@{ "emailAddress" = @{ "address" = ($PSItem.Address) } }) }
  }

  if ($PSBoundParameters.ContainsKey("BCC"))
  {
    [Void]$Message.Message.Add("bccRecipients", ([System.Collections.ArrayList]::New()))
    $BCC | ForEach-Object -Process { [Void]$Message.Message.bccRecipients.Add(@{ "emailAddress" = @{ "address" = ($PSItem.Address) } }) }
  }

  if ($PSBoundParameters.ContainsKey("ReplyTo"))
  {
    [Void]$Message.Message.Add("replyTo", ([System.Collections.ArrayList]::New()))
    $ReplyTo | ForEach-Object -Process { [Void]$Message.Message.replyTo.Add(@{ "emailAddress" = @{ "address" = ($PSItem.Address) } }) }
  }

  if ($PSBoundParameters.ContainsKey("Attachments"))
  {
    [Void]$Message.Message.Add("Attachments", ([System.Collections.ArrayList]::New()))
    foreach ($File in $Attachments)
    {
      if ([System.IO.File]::Exists($File))
      {
        $Base64Encode = [Convert]::ToBase64String(([System.IO.File]::ReadAllBytes($File)))
        [Void]$Message.Message.attachments.Add([Ordered]@{ "@odata.type" = "#microsoft.graph.fileAttachment"; "Name" = ([System.IO.Path]::GetFileName($File)); "contentType" = "MIME types"; "contentBytes" = "$($Base64Encode)" })
      }
    }
  }

  [Void]$Message.Add("saveToSentItems", ($SaveToSent.IsPresent.ToString()))

  if ($PSCmdlet.ParameterSetName -eq "Me")
  {
    $Uri = "https://graph.microsoft.com/$($Version)/me/sendmail"
  }
  else
  {
    $Uri = "https://graph.microsoft.com/$($Version)/users/$($UserID)/sendmail"
  }

  $Result = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Headers $AuthToken -Method Post -Body ($Message | ConvertTo-Json -Depth 99)

  [PSCustomObject]@{ "Success" = ($Result.StatusCode -eq 202) }

  Write-Verbose -Message "Exit Function Send-MyGraphMail"
}
#endregion function Send-MyGraphMail
