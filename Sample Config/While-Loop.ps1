<#
  .SYNOPSIS
    Sample Runspace Pool Thread Script
  .DESCRIPTION
    Sample Runspace Pool Thread Script
  .PARAMETER ListViewItem
    ListViewItem Passed to the Thread Script

    This Paramter is Required in your Thread Script
  .EXAMPLE
    Test-Script.ps1 -ListViewItem $ListViewItem
  .NOTES
    Sample Thread Script
#>
[CmdletBinding()]
Param (
  [parameter(Mandatory = $True)]
  [System.Windows.Forms.ListViewItem]$ListViewItem
)

# Set Preference Variables
$ErrorActionPreference = "Stop"
$VerbosePreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"

$Random = [System.Random]::New()
While ($True)
{
  For ($I = 1; $I -lt 24; $I++)
  {
    $ListViewItem.SubItems[$I].Text = $Random.Next(0, 9)
    [System.Threading.Thread]::Sleep(100)
  }
  [System.Threading.Thread]::Sleep(100)
}

Exit
