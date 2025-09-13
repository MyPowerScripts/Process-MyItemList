
#region function Get-MyADForest
function Get-MyADForest ()
{
  <#
    .SYNOPSIS
      Gets information about an Active Directory Forest.
    .DESCRIPTION
      Retrieves the Active Directory Forest object either for the current forest or for a specified forest name.
    .PARAMETER Name
      The name of the Active Directory forest to retrieve. This parameter is mandatory when using the "Name" parameter set.
    .EXAMPLE
      PS C:\> Get-MyADForest
      Retrieves the current Active Directory forest.
    .EXAMPLE
      PS C:\> Get-MyADForest -Name "contoso.com"
      Retrieves the Active Directory forest with the name "contoso.com".
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Current")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "Name")]
    [String]$Name
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  switch ($PSCmdlet.ParameterSetName)
  {
    "Name"
    {
      $DirectoryContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Forest
      $DirectoryContext = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::New($DirectoryContextType, $Name)
      [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($DirectoryContext)
      $DirectoryContext = $Null
      $DirectoryContextType = $Null
      break
    }
    "Current"
    {
      [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
      break
    }
  }

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Get-MyADForest

#region function Get-MyADDomain
function Get-MyADDomain ()
{
  <#
    .SYNOPSIS
      Gets information about an Active Directory Domain.
    .DESCRIPTION
      Retrieves the Active Directory Domain object either for the current domain, a specified domain name, or the domain associated with the local computer.
    .PARAMETER Name
      The name of the Active Directory domain to retrieve. This parameter is mandatory when using the "Name" parameter set.
    .PARAMETER Computer
      Switch parameter. If specified, retrieves the Active Directory domain associated with the local computer. This parameter is mandatory when using the "Computer" parameter set.
    .EXAMPLE
      PS C:\> Get-MyADDomain
      Retrieves the current Active Directory domain.
    .EXAMPLE
      PS C:\> Get-MyADDomain -Computer
      Retrieves the Active Directory domain associated with the local computer.
    .EXAMPLE
      PS C:\> Get-MyADDomain -Name "contoso.com"
      Retrieves the Active Directory domain with the name "contoso.com".
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Current")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "Name")]
    [String]$Name,
    [parameter(Mandatory = $True, ParameterSetName = "Computer")]
    [Switch]$Computer
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  switch ($PSCmdlet.ParameterSetName)
  {
    "Name"
    {
      $DirectoryContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain
      $DirectoryContext = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::New($DirectoryContextType, $Name)
      [System.DirectoryServices.ActiveDirectory.Domian]::GetDomain($DirectoryContext)
      $DirectoryContext = $Null
      $DirectoryContextType = $Null
      break
    }
    "Computer"
    {
      [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
      break
    }
    "Current"
    {
      [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
      break
    }
  }

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Get-MyADDomain

#region function Get-MyADObject
function Get-MyADObject()
{
  <#
    .SYNOPSIS
      Searches Active Directory and returns an AD SearchResultCollection.
    .DESCRIPTION
      Performs a search in Active Directory using the specified LDAP filter and returns a SearchResultCollection. 
      Supports specifying search root, server, credentials, properties to load, sorting, and paging options.
    .PARAMETER LDAPFilter
      The LDAP filter string to use for the search. Defaults to (objectClass=*).
    .PARAMETER PageSize
      The number of objects to return per page. Default is 1000.
    .PARAMETER SizeLimit
      The maximum number of objects to return. Default is 1000.
    .PARAMETER SearchRoot
      The LDAP path to start the search from. Defaults to the current domain root.
    .PARAMETER ServerName
      The name of the domain controller or server to query. If not specified, uses the default.
    .PARAMETER SearchScope
      The scope of the search. Valid values are Base, OneLevel, or Subtree. Default is Subtree.
    .PARAMETER Sort
      The direction to sort the results. Valid values are Ascending or Descending. Default is Ascending.
    .PARAMETER SortProperty
      The property name to sort the results by.
    .PARAMETER PropertiesToLoad
      An array of property names to load for each result.
    .PARAMETER Credential
      The credentials to use when searching Active Directory.
    .EXAMPLE
      Get-MyADObject -LDAPFilter "(objectClass=user)" -SearchRoot "OU=Users,DC=domain,DC=com"
      Searches for all user objects in the specified OU.
    .EXAMPLE
      Get-MyADObject -ServerName "dc01.domain.com" -PropertiesToLoad "samaccountname","mail"
      Searches using a specific domain controller and returns only the samaccountname and mail properties.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [String]$LDAPFilter = "(objectClass=*)",
    [Long]$PageSize = 1000,
    [Long]$SizeLimit = 1000,
    [String]$SearchRoot = "LDAP://$($([ADSI]'').distinguishedName)",
    [String]$ServerName,
    [ValidateSet("Base", "OneLevel", "Subtree")]
    [System.DirectoryServices.SearchScope]$SearchScope = "SubTree",
    [ValidateSet("Ascending", "Descending")]
    [System.DirectoryServices.SortDirection]$Sort = "Ascending",
    [String]$SortProperty,
    [String[]]$PropertiesToLoad,
    [PSCredential]$Credential
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  $MySearcher = [System.DirectoryServices.DirectorySearcher]::New($LDAPFilter, $PropertiesToLoad, $SearchScope)

  $MySearcher.PageSize = $PageSize
  $MySearcher.SizeLimit = $SizeLimit

  $TempSearchRoot = $SearchRoot.ToUpper()
  switch -regex ($TempSearchRoot)
  {
    "(?:LDAP|GC)://*"
    {
      if ($PSBoundParameters.ContainsKey("ServerName"))
      {
        $MySearchRoot = $TempSearchRoot -replace "(?<LG>(?:LDAP|GC)://)(?:[\w\d\.-]+/)?(?<DN>.+)", "`${LG}$($ServerName)/`${DN}"
      }
      else
      {
        $MySearchRoot = $TempSearchRoot
      }
      break
    }
    default
    {
      if ($PSBoundParameters.ContainsKey("ServerName"))
      {
        $MySearchRoot = "LDAP://$($ServerName)/$($TempSearchRoot)"
      }
      else
      {
        $MySearchRoot = "LDAP://$($TempSearchRoot)"
      }
      break
    }
  }

  if ($PSBoundParameters.ContainsKey("Credential"))
  {
    $MySearcher.SearchRoot = [System.DirectoryServices.DirectoryEntry]::New($MySearchRoot, ($Credential.UserName), (($Credential.GetNetworkCredential()).Password))
  }
  else
  {
    $MySearcher.SearchRoot = [System.DirectoryServices.DirectoryEntry]::New($MySearchRoot)
  }

  if ($PSBoundParameters.ContainsKey("SortProperty"))
  {
    $MySearcher.Sort.PropertyName = $SortProperty
    $MySearcher.Sort.Direction = $Sort
  }

  $MySearcher.FindAll()

  $MySearcher.Dispose()
  $MySearcher = $Null
  $MySearchRoot = $Null
  $TempSearchRoot = $Null

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Get-MyADObject
