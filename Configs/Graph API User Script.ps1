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

   -------------------------
   ListViewItem Status Icons
   -------------------------
   $GoodIcon = Solid Green Circle
   $BadIcon = Solid Red Circle
   $InfoIcon = Solid Blue Circle
   $CheckIcon = Checkmark
   $ErrorIcon = Red X
   $UpIcon = Green up Arrow 
   $DownIcon = Red Down Arrow

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

# -----------------------------------------------------
# Build ListView Column Lookup Table
#
# Reference Columns by Name Incase Column Order Changes
# -----------------------------------------------------
$Columns = @{
}
$ListViewItem.ListView.Columns | ForEach-Object -Process {
  $Columns.Add($PSItem.Text, $PSItem.Index)
}

# ------------------------------------------------
# Check if Thread was Already Completed and Exit
# ------------------------------------------------
If ($ListViewItem.SubItems[$Columns["Status Message"]].Text -eq "Completed")
{
  $ListViewItem.ImageKey = $GoodIcon
  Exit
}

# ----------------------------------------------------
# Check if Threads are Paused and Update Thread Status
# ----------------------------------------------------
If ($SyncedHash.Pause)
{
  # Set Paused Status
  $ListViewItem.SubItems[$Columns["Status Message"]].Text = "Pause"
  $ListViewItem.SubItems[$Columns["Date/Time"]].Text = [DateTime]::Now.ToString("G")
  While ($SyncedHash.Pause)
  {
    [System.Threading.Thread]::Sleep(100)
  }
}

# -----------------------------------------------------
# Check For Termination and Update Thread Status
# -----------------------------------------------------
If ($SyncedHash.Terminate)
{
  # Set Terminated Status and Exit Thread
  $ListViewItem.SubItems[$Columns["Status Message"]].Text = "Terminated"
  $ListViewItem.SubItems[$Columns["Date/Time"]].Text = [DateTime]::Now.ToString("G")
  $ListViewItem.ImageKey = $InfoIcon
  Exit
}

# Sucess Default Exit Status
$WasSuccess = $True
$ListViewItem.SubItems[$Columns["Status Message"]].Text = "Processing"
$UserPrincipalName = $ListViewItem.SubItems[$Columns["UserPrincipalName"]].Text

Try
{
  $AuthToken = Get-MyOAuthApplicationToken -MyTenantID $TenantID -MyClientID $ClientID -MyClientSecret $ClientSecret
  If ($AuthToken.Expires_In -eq 0)
  {
    $ListViewItem.SubItems[$Columns["Error Message"]].Text = "Unable to get AuthToken"
    $WasSuccess = $False
  }
  Else
  {
    $Resource = ("/users/$($UserPrincipalName)?`$top=1&`$select=id,displayname,mail,givenName,surname,accountEnabled" -f $DisplayName)
    $Device = Get-MyGQuery -AuthToken $AuthToken -Resource $Resource -ErrorAction SilentlyContinue
    If ([String]::IsNullOrEmpty($Device.ID))
    {
      $ListViewItem.SubItems[$Columns["Error Message"]].Text = "No Device Found in Azure AD / Entra ID"
      $WasSuccess = $False
    }
    Else
    {
      $ListViewItem.SubItems[$Columns["ID"]].Text = $Device.ID
      $ListViewItem.SubItems[$Columns["E-Mail"]].Text = $Device.Mail
      $ListViewItem.SubItems[$Columns["DisplayName"]].Text = $Device.DisplayName
      $ListViewItem.SubItems[$Columns["FirstName"]].Text = $Device.GivenName
      $ListViewItem.SubItems[$Columns["Surname"]].Text = $Device.Surname
      $ListViewItem.SubItems[$Columns["AccountEnabled"]].Text = $Device.AccountEnabled
    }
  }
}
Catch
{
  # Set Error Message / Thread Failed
  $ListViewItem.SubItems[$Columns["Error Message"]].Text = $PSItem.ToString()
  $WasSuccess = $False
  
  Write-Host -Object ($($Error[0].Exception.Message))
  Write-Host -Object (($Error[0].InvocationInfo.Line).Trim())
  Write-Host -Object ($Error[0].InvocationInfo.ScriptLineNumber)
}

# Set Final Date / Time and Update Status
$ListViewItem.SubItems[$Columns["Date/Time"]].Text = [DateTime]::Now.ToString("G")
If ($WasSuccess)
{
  # Return Success
  $ListViewItem.ImageKey = $GoodIcon
  $ListViewItem.SubItems[$Columns["Status Message"]].Text = "Completed"
  $ListViewItem.SubItems[$Columns["Error Message"]].Text = ""
  $ListViewItem.Tag = "Completed"
}
Else
{
  # Return Success
  $ListViewItem.ImageKey = $BadIcon
  $ListViewItem.SubItems[$Columns["Status Message"]].Text = "Error"
}

Exit


