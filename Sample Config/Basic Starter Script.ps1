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

$ErrorActionPreference = "Stop"
$VerbosePreference = "SilentlyContinue"

# Common Columns
$ItemCol = 0
$DataCol = 1
$StatusCol = 2
$DateTimeCol = 3
$ErrorCol = 4

# ------------------------------------------------
# Check if Thread was Already Completed and Exit
# ------------------------------------------------
If ($ListViewItem.SubItems[$StatusCol].Text -eq "Completed")
{
  $ListViewItem.ImageKey = $GoodIcon
  Exit
}

# ----------------------------------------------------
# Check if Threads are Paused and Update Thread Status
# ----------------------------------------------------
If ($SyncedHash.Paused)
{
  # Set Paused Status
  $ListViewItem.SubItems[1].Text = "Pause"
  While ($SyncedHash.Paused)
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
  $ListViewItem.SubItems[$StatusCol].Text = "Terminated"
  $ListViewItem.SubItems[$DateTimeCol].Text = [DateTime]::Now.ToString("G")
  $ListViewItem.ImageKey = $InfoIcon
  Exit
}

# Sucess Default Exit Status
$WasSuccess = $True
$CurrentItem = $ListViewItem.SubItems[$ItemCol].Text
Try
{
  
  # Get / Update Shared Object / Value
  If ([System.String]::IsNullOrEmpty($SyncedHash.Object))
  {
    $SyncedHash.Object = "First Item"
  }
  $ListViewItem.SubItems[$DataCol].Text = $SyncedHash.Object
  $SyncedHash.Object = $CurrentItem
  
  # ---------------------------------------------------------
  # Open and wait for Mutex - Limit Access to Shared Resource
  # ---------------------------------------------------------
  $MyMutex = [System.Threading.Mutex]::OpenExisting($Mutex)
  [Void]($MyMutex.WaitOne())
  
  # Access / Update Shared Resources
  # $CurrentItem | Out-File -Encoding ascii -FilePath "C:\SharedFile.txt"
  
  # Release Mutex
  $MyMutex.ReleaseMutex()
  
}
Catch
{
  # Set Error Message / Thread Failed
  $ListViewItem.SubItems[$ErrorCol].Text = $PSItem.ToString()
  $WasSuccess = $False
}

# Set Final Date / Time and Update Status
$ListViewItem.SubItems[$DateTimeCol].Text = [DateTime]::Now.ToString("G")
If ($WasSuccess)
{
  # Return Success
  $ListViewItem.ImageKey = $GoodIcon
  $ListViewItem.SubItems[$StatusCol].Text = "Completed"
}
Else
{
  # Return Success
  $ListViewItem.ImageKey = $BadIcon
  $ListViewItem.SubItems[$StatusCol].Text = "Error"
}

Write-Host -Object $ListViewItem.ImageKey

Exit
