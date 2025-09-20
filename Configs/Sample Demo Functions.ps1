
#region **** Function Return-HelloWorld ****
Function Return-HelloWorld ()
{
  <#
    .SYNOPSIS
      Example Funciton
    .DESCRIPTION
      Example Funciton
    .PARAMETER InputValue
      Required Input Value
    .EXAMPLE
      Return-HelloWorld -InputValue $InputValue
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory = $true)]
    [String]$InputValue
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"
  
  Return $InputValue
  
  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion **** Function Return-HelloWorld ****

