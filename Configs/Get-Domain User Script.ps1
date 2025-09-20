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
$ListViewItem.SubItems[$Columns["Date/Time"]].Text = [DateTime]::Now.ToString("G")
$UserName = $ListViewItem.SubItems[$Columns["UserName"]].Text

Try
{
  # Get Current Domain / Forest
  If ($ADForest -eq "Domain")
  {
    If ($ADDomain -eq "Current")
    {
      $GetADDomain = Get-MyADDomain -ErrorAction SilentlyContinue
    }
    Else
    {
      $GetADDomain = Get-MyADDomain -Name $ADDomain -ErrorAction SilentlyContinue
    }
    If (-not [String]::IsNullOrEmpty($GetADDomain.Name))
    {
      $SearchRoot = "LDAP://$("dc=$(($GetADDomain.Name -split '\.') -join ',dc=')")"
      $GetADDomain.Dispose()
    }
  }
  Else
  {
    If ($ADDomain -eq "Current")
    {
      $GetADForget = Get-MyADForest -ErrorAction SilentlyContinue
    }
    Else
    {
      $GetADForget = Get-MyADForest -Name $ADForest -ErrorAction SilentlyContinue
    }
    If (-not [String]::IsNullOrEmpty($GetADForget.Name))
    {
      $SearchRoot = "GC://$($GetADForget.Name)"
      $GetADForget.Dispose()
    }
  }
  
  # Check Domain / Forest Found
  If ([String]::IsNullOrEmpty($SearchRoot))
  {
    $ListViewItem.SubItems[$Columns["Error Message"]].Text = "Unable to Get Current AD Domain / Forest"
    $WasSuccess = $False
  }
  Else
  {
    $LDAPFilter = "(&(objectClass=user)(objectCategory=person)(sAMAccountType=805306368)(sAMAccountName={0}))" -f $UserName
    $PropertiesToLoad = @("name", "canonicalName", "userPrincipalName", "mail", "lastLogonTimestamp", "pwdLastSet", "userAccountControl", "distinguishedName")
    $ADObject = Get-MyADObject -SearchRoot $SearchRoot -SearchScope Subtree -LDAPFilter $LDAPFilter -PropertiesToLoad $PropertiesToLoad -ErrorAction SilentlyContinue | Select-Object -First 1
    If ([String]::IsNullOrEmpty($ADObject.Path))
    {
      $ListViewItem.SubItems[$Columns["Error Message"]].Text = "Computer Not Found in AD Forest"
      $WasSuccess = $False
    }
    Else
    {
      # User Info
      $ListViewItem.SubItems[$Columns["userPrincipalName"]].Text = $ADObject.Properties["userPrincipalName"][0]
      $ListViewItem.SubItems[$Columns["E-Mail"]].Text = $ADObject.Properties["mail"][0]
      
      # CanonicalName
      $CanonicalName = $ADObject.Properties["canonicalName"][0]
      $ListViewItem.SubItems[$Columns["canonicalName"]].Text = $CanonicalName
      
      # distinguishedName
      $ListViewItem.SubItems[$Columns["distinguishedName"]].Text = $ADObject.Properties["distinguishedName"][0]
      
      # Domain
      $Domain = $CanonicalName -split "/" | Select-Object -First 1
      $ListViewItem.SubItems[$Columns["Domain"]].Text = $Domain
      
      # Zero Hour
      $ZeroHour = [DateTime]::New(1601, 1, 1, 0, 0, 0)
      
      # Last Logon TimeStamp
      $LastLogonTimestamp = $ADObject.Properties["lastLogonTimestamp"][0]
      $LastLogonTimeStampDate = $ZeroHour.AddTicks($LastLogonTimestamp)
      $ListViewItem.SubItems[$Columns["Last Logon"]].Text = $LastLogonTimeStampDate.ToString("G")
      
      # Password Last Set
      $PwdLastSet = $ADObject.Properties["pwdLastSet"][0]
      $PwdLastSetDate = $ZeroHour.AddTicks($PwdLastSet)
      $ListViewItem.SubItems[$Columns["PwdLastSet"]].Text = $PwdLastSetDate.ToString("G")
      
      # User Account Control Flags
      $UserAccountControl = $ADObject.Properties["userAccountControl"][0]
      $ListViewItem.SubItems[$Columns["UserAccountControl"]].Text = "0$($UserAccountControl)"
      $ListViewItem.SubItems[$Columns["PwdNoChange"]].Text = (($UserAccountControl -band 64) -ne 0)
      $ListViewItem.SubItems[$Columns["PwdNoExpire"]].Text = (($UserAccountControl -band 65536) -ne 0)
      $ListViewItem.SubItems[$Columns["PwdExpired"]].Text = (($UserAccountControl -band 8388608) -ne 0)
      $ListViewItem.SubItems[$Columns["Locked Out"]].Text = (($UserAccountControl -band 16) -ne 0)
      $ListViewItem.SubItems[$Columns["Disabled"]].Text = (($UserAccountControl -band 2) -ne 0)
    }
  }
}
Catch
{
  # Set Error Message / Thread Failed
  $ListViewItem.SubItems[$Columns["Error Message"]].Text = $PSItem.ToString()
  $WasSuccess = $False
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



