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
[CmdletBinding(DefaultParameterSetName = "ByValue")]
Param (
  [parameter(Mandatory = $True)]
  [System.Windows.Forms.ListViewItem]$ListViewItem
)

$ErrorActionPreference = "Stop"
$VerbosePreference = "SilentlyContinue"

# ------------------------------------------------
# Check if Thread was Already Completed and Exit
#
# One Column needs to be the Status the the Thread
#  Status Messages are Customizable
# ------------------------------------------------
If ($ListViewItem.SubItems[1].Text -eq "Completed")
{
  $ListViewItem.ImageKey = $GoodIcon
  Exit
}

# ----------------------------------------------------
# Check if Threads are Paused and Update Thread Status
#
# You can add Multiple Checks for Pasue if Needed
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
#
# You can add Multiple Checks for Termination if Needed
# -----------------------------------------------------
If ($SyncedHash.Terminate)
{
  # Set Terminated Status and Return
  $ListViewItem.SubItems[1].Text = "Terminated"
  $ListViewItem.SubItems[2].Text = [DateTime]::Now.ToString("g")
  $ListViewItem.ImageKey = $InfoIcon
  Exit
}

# Set Proccessing Ststus
$ListViewItem.SubItems[1].Text = "Processing"
$ListViewItem.SubItems[2].Text = [DateTime]::Now.ToString("g")
$WasSuccess = $True

# --------------------------------------------------
# Get Curent List Item
#
# Coulmn 0 Always has the List Item to be Proccessed
# --------------------------------------------------
$CurentItem = $ListViewItem.SubItems[0].Text

# --------------------------------------------------------------
# Open and wait for Mutex
# 
# This is to Pause the Thread Script if Access a Shared Resource
#   and you need toi Limit to 1 Thread at a Time
#
# Using a Mutext is Optional
# --------------------------------------------------------------
$MyMutex = [System.Threading.Mutex]::OpenExisting($Mutex)
[Void]($MyMutex.WaitOne())

# Set Date / Time when Mutext was Opened
$ListViewItem.SubItems[3].Text = [DateTime]::Now.ToString("g")

# --------------------------------------------------------------------------------
# The Synced HashTable has an Object Property to share information between Threads
# --------------------------------------------------------------------------------
If ([String]::IsNullOrEmpty($SyncedHash.Object))
{
  $SyncedHash.Object = "First"
}
$ListViewItem.SubItems[4].Text = $SyncedHash.Object
$SyncedHash.Object = $CurentItem

# Release Mutex
$MyMutex.ReleaseMutex()

# Random Number Generator
$Random = [System.Random]::New()

# ---------------------------------------------------------
# Gernate a Fake Error
#
# Make sure to use Error Catching to make sure thread exits
# ---------------------------------------------------------
Try
{
  Switch ($Random.Next(0, 3))
  {
    "0"
    {
      Throw "This is a Fake Error!"
      Break
    }
    "1"
    {
      Throw "Simulated Error!"
      Break
    }
    "2"
    {
      Throw "Someing Failed!"
      Break
    }
    "3"
    {
      Throw "Unknown Error!"
      Break
    }
  }
}
Catch
{
  # Save Error Mesage
  $ListViewItem.SubItems[5].Text = $Error[0].Exception.Message
}


For ($I = 8; $I -lt 16; $I++)
{
  $ListViewItem.SubItems[$I].Text = [DateTime]::Now.ToString("HH:mm:ss:ffff")
  [System.Threading.Thread]::Sleep(100)
}

$RndValue = $Random.Next(0, 3)
$ListViewItem.SubItems[6].Text = $RndValue
# Random Fail Simlater
If ($RndValue -eq 0)
{
  $WasSuccess = $False
}
$ListViewItem.SubItems[7].Text = $WasSuccess

If ($WasSuccess)
{
  # Return Success
  $ListViewItem.ImageKey = $GoodIcon
  $ListViewItem.SubItems[1].Text = "Completed"
}
Else
{
  # Return Success
  $ListViewItem.ImageKey = $BadIcon
  $ListViewItem.SubItems[1].Text = "Error"
}

Exit
