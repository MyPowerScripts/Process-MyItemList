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
$Columns = @{}
$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }

# ------------------------------------------------------------
# Check if Thread was Already Completed and Exits Thread
#
# Required to Prevent the Reprocessing of Completed List Items
# ------------------------------------------------------------
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

# -------------------------------------------------------
# Check For Termination and Update Thread Status and Exit
# -------------------------------------------------------
If ($SyncedHash.Terminate)
{
  # Set Terminated Status and Exit Thread
  $ListViewItem.SubItems[$Columns["Status Message"]].Text = "Terminated"
  $ListViewItem.SubItems[$Columns["Date/Time"]].Text = [DateTime]::Now.ToString("G")
  $ListViewItem.ImageKey = $InfoIcon
  Exit
}

# Set Default Exit Status
$WasSuccess = $True

# Set List Item Status to Processing
$ListViewItem.SubItems[$Columns["Status Message"]].Text = "Processing"

# Get List Litem to Process
$CurrentItem = $ListViewItem.SubItems[$Columns["List Item"]].Text

Try
{
  # Get / Update Shared Object / Value
  If ([System.String]::IsNullOrEmpty($SyncedHash.Object))
  {
    $SyncedHash.Object = "First Item"
  }
  $ListViewItem.SubItems[$Columns["Data Column"]].Text = $SyncedHash.Object
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
  $ListViewItem.SubItems[$Columns["Error Message"]].Text = $PSItem.ToString()
  $WasSuccess = $False
}

# File Remaining Columns - Sample Data
For ($I = 4; $I -lt 11; $I++)
{
  $ListViewItem.SubItems[$I].Text = [DateTime]::Now.ToString("G")
  [System.Threading.Thread]::Sleep(100)
}

# --------------------------------------------
# Set Final Date / Time and Update Exit Status
# --------------------------------------------
$ListViewItem.SubItems[$Columns["Date/Time"]].Text = [DateTime]::Now.ToString("G")
If ($WasSuccess)
{
  # ---------------------
  # Return Success Status
  # ---------------------
  $ListViewItem.ImageKey = $GoodIcon
  $ListViewItem.SubItems[$Columns["Status Message"]].Text = "Completed"
  $ListViewItem.SubItems[$Columns["Error Message"]].Text = ""
  
  # ----------------------------------------------------------------------
  # Optional, Set Tag to Any Value to Prevent Reprocessing Completed Items
  # 
  # This Does not Prevent Processing Items from a PIL Data Export that
  #   were Previously Completed
  # ----------------------------------------------------------------------
  $ListViewItem.Tag = "Completed"
}
Else
{
  # -------------------
  # Return Error Status
  # -------------------
  $ListViewItem.ImageKey = $BadIcon
  $ListViewItem.SubItems[$Columns["Status Message"]].Text = "Error"
}

Exit







