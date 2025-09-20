<#
  .SYNOPSIS
    Sample Runspace Pool Thread Script
  .DESCRIPTION
    Sample Runspace Pool Thread Script
  .PARAMETER ListViewItem
    ListViewItem Passed to the Thread Script

    This Paramter is Required in your Thread Script
  .EXAMPLE
    Test-Script.ps$Columns["Status"] -ListViewItem $ListViewItem
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
#
# One Column needs to be the Status the the Thread
#  Status Messages are Customizable
# ------------------------------------------------
If ($ListViewItem.SubItems[$Columns["Status"]].Text -eq "Completed")
{
  $ListViewItem.ImageKey = $GoodIcon
  Exit
}

# ----------------------------------------------------
# Check if Threads are Paused and Update Thread Status
#
# You can add Multiple Checks for Pasue if Needed
# ----------------------------------------------------
If ($SyncedHash.Pause)
{
  # Set Paused Status
  $ListViewItem.SubItems[$Columns["Status"]].Text = "Pause"
  While ($SyncedHash.Pause)
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
  $ListViewItem.SubItems[$Columns["Status"]].Text = "Terminated"
  $ListViewItem.SubItems[$Columns["Term/Proc Times"]].Text = [DateTime]::Now.ToString("HH:mm:ss:ffff")
  $ListViewItem.ImageKey = $InfoIcon
  Exit
}

# Set Proccessing Ststus
$ListViewItem.SubItems[$Columns["Status"]].Text = "Processing"
$ListViewItem.SubItems[$Columns["Term/Proc Times"]].Text = [DateTime]::Now.ToString("HH:mm:ss:ffff")
$WasSuccess = $True

# Set Prompt Variable
$ListViewItem.SubItems[$Columns["Prompt Variable"]].Text = $PromptVariable

# --------------------------------------------------
# Get Curent List Item
#
# Coulmn 0 Always has the List Item to be Proccessed
# --------------------------------------------------
$CurentItem = $ListViewItem.SubItems[$Columns["List Item"]].Text
# For Testing you can Write to the Screen
Write-Host -Object "Processing $($CurentItem)"

# --------------------------------------------------------------
# Open and wait for Mutex
# 
# This is to Pause the Thread Script if Access a Shared Resource
#   and you need toi Limit to $Columns["Status"] Thread at a Time
#
# Using a Mutext is Optional
# --------------------------------------------------------------
$MyMutex = [System.Threading.Mutex]::OpenExisting($Mutex)
[Void]($MyMutex.WaitOne())

# Set Date / Time when Mutext was Opened
$ListViewItem.SubItems[$Columns["Open Mutex"]].Text = [DateTime]::Now.ToString("HH:mm:ss:ffff")

# Access / Update Shared Resources
# $CurrentItem | Out-File -Encoding ascii -FilePath "C:\SharedFile.txt"

# Release Mutex
$MyMutex.ReleaseMutex()

# --------------------------------------------------------------------------------
# The Synced HashTable has an Object Property to share information between Threads
# --------------------------------------------------------------------------------
If ([String]::IsNullOrEmpty($SyncedHash.Object))
{
  $SyncedHash.Object = "First"
}
$ListViewItem.SubItems[$Columns["Synced Hash"]].Text = $SyncedHash.Object
$SyncedHash.Object = $CurentItem


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
  $ListViewItem.SubItems[$Columns["Fake Error"]].Text = $Error[0].Exception.Message
}

$ListViewItem.SubItems[$Columns["Function Test"]].Text = Return-HelloWorld -InputValue "Hello World"
$ListViewItem.SubItems[$Columns["Static Variable"]].Text = $StaticVariable

$RndValue = $Random.Next(0, 3)
For ($I = 10; $I -lt 16; $I++)
{
  $ListViewItem.SubItems[$I].Text = [DateTime]::Now.ToString("HH:mm:ss:ffff")
  [System.Threading.Thread]::Sleep(100)
}

# Random Fail Simlater
If ($RndValue -eq 0)
{
  $WasSuccess = $False
}
$ListViewItem.SubItems[$Columns["WasSuccess"]].Text = $WasSuccess

# Set Final Date / Time and Update Status
$ListViewItem.SubItems[$Columns["Term/Proc Times"]].Text = [DateTime]::Now.ToString("HH:mm:ss:ffff")
If ($WasSuccess)
{
  # Return Success
  $ListViewItem.ImageKey = $GoodIcon
  $ListViewItem.SubItems[$Columns["Status"]].Text = "Completed"
  $ListViewItem.Tag = "Completed"
}
Else
{
  # Return Success
  $ListViewItem.ImageKey = $BadIcon
  $ListViewItem.SubItems[$Columns["Status"]].Text = "Error"
}

# Testing Write to Screen
Write-Host -Object "Completed $($CurentItem)"

Exit


