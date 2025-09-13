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

#region class MyWorkstationInfo
Class MyWorkstationInfo
{
  [String]$ComputerName = [Environment]::MachineName
  [String]$FQDN = [Environment]::MachineName
  [Bool]$Found = $False
  [String]$UserName = ""
  [String]$Domain = ""
  [Bool]$DomainMember = $False
  [int]$ProductType = 0
  [String]$Manufacturer = ""
  [String]$Model = ""
  [Bool]$IsMobile = $False
  [String]$SerialNumber = ""
  [Long]$Memory = 0
  [String]$OperatingSystem = ""
  [String]$BuildNumber = ""
  [String]$Version = ""
  [String]$ServicePack = ""
  [String]$Architecture = ""
  [Bool]$Is64Bit = $False
  [DateTime]$LocalDateTime = [DateTime]::MinValue
  [DateTime]$InstallDate = [DateTime]::MinValue
  [DateTime]$LastBootUpTime = [DateTime]::MinValue
  [String]$IPAddress = ""
  [String]$Status = "Off-Line"
  [DateTime]$StartTime = [DateTime]::Now
  [DateTime]$EndTime = [DateTime]::Now
  
  MyWorkstationInfo ([String]$ComputerName)
  {
    $This.ComputerName = $ComputerName.ToUpper()
    $This.FQDN = $ComputerName.ToUpper()
    $This.Status = "On-Line"
  }
  
  [Void] AddComputerSystem ([String]$TestName, [String]$IPAddress, [String]$ComputerName, [Bool]$DomainMember, [String]$Domain, [String]$Manufacturer, [String]$Model, [String]$UserName, [Long]$Memory)
  {
    $This.IPAddress = $IPAddress
    $This.ComputerName = "$($ComputerName)".ToUpper()
    $This.DomainMember = $DomainMember
    $This.Domain = "$($Domain)".ToUpper()
    If ($DomainMember)
    {
      $This.FQDN = "$($ComputerName).$($Domain)".ToUpper()
    }
    $This.Manufacturer = $Manufacturer
    $This.Model = $Model
    $This.UserName = $UserName
    $This.Memory = $Memory
    $This.Found = ($ComputerName -eq @($TestName.Split("."))[0])
  }
  
  [Void] AddOperatingSystem ([int]$ProductType, [String]$OperatingSystem, [String]$ServicePack, [String]$BuildNumber, [String]$Version, [String]$Architecture, [DateTime]$LocalDateTime, [DateTime]$InstallDate, [DateTime]$LastBootUpTime)
  {
    $This.ProductType = $ProductType
    $This.OperatingSystem = $OperatingSystem
    $This.ServicePack = $ServicePack
    $This.BuildNumber = $BuildNumber
    $This.Version = $Version
    $This.Architecture = $Architecture
    $This.Is64Bit = ($Architecture -eq "64-bit")
    $This.LocalDateTime = $LocalDateTime
    $This.InstallDate = $InstallDate
    $This.LastBootUpTime = $LastBootUpTime
  }
  
  [Void] AddSerialNumber ([String]$SerialNumber)
  {
    $This.SerialNumber = $SerialNumber
  }
  
  [Void] AddIsMobile ([Long[]]$ChassisTypes)
  {
    $This.IsMobile = (@(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32) -contains $ChassisTypes[0])
  }
  
  [Void] UpdateStatus ([String]$Status)
  {
    $This.Status = $Status
  }
  
  [MyWorkstationInfo] SetEndTime ()
  {
    $This.EndTime = [DateTime]::Now
    Return $This
  }
  
  [TimeSpan] GetRunTime ()
  {
    Return ($This.EndTime - $This.StartTime)
  }
}
#endregion class MyWorkstationInfo

# Common Columns
$StatusCol = 17
$DateTimeCol = 18
$ErrorCol = 19

# ------------------------------------------------
# Check if Thread was Already Completed and Exit
#
# One Column needs to be the Status the the Thread
#  Status Messages are Customizable
# ------------------------------------------------
If ($ListViewItem.SubItems[$StatusCol].Text -eq "Completed")
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
  $ListViewItem.SubItems[$StatusCol].Text = "Pause"
  $ListViewItem.SubItems[$DateTimeCol].Text = [DateTime]::Now.ToString("g")
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
  $ListViewItem.SubItems[$StatusCol].Text = "Terminated"
  $ListViewItem.SubItems[$DateTimeCol].Text = [DateTime]::Now.ToString("g")
  $ListViewItem.ImageKey = $InfoIcon
  Exit
}

# --------------------------------------------------
# Get Curent List Item
# --------------------------------------------------
$ComputerName = $ListViewItem.SubItems[0].Text

# Set Proccessing Ststus
$ListViewItem.SubItems[$StatusCol].Text = "Processing"
$ListViewItem.SubItems[$DateTimeCol].Text = [DateTime]::Now.ToString("g")

Try
{
  $WorkstationInfo = Get-MyWorkstationInfo -ComputerName $ComputerName -Serial -Mobile
  $WasSuccess = $WorkstationInfo.Found
  
  $ListViewItem.SubItems[01].Text = $WorkstationInfo.Status
  $ListViewItem.SubItems[02].Text = $WorkstationInfo.IPAddress
  $ListViewItem.SubItems[03].Text = $WorkstationInfo.FQDN
  $ListViewItem.SubItems[04].Text = $WorkstationInfo.Domain
  $ListViewItem.SubItems[05].Text = $WorkstationInfo.ComputerName
  $ListViewItem.SubItems[06].Text = $WorkstationInfo.UserName
  $ListViewItem.SubItems[07].Text = $WorkstationInfo.OperatingSystem
  $ListViewItem.SubItems[08].Text = $WorkstationInfo.BuildNumber
  $ListViewItem.SubItems[09].Text = $WorkstationInfo.Architecture
  $ListViewItem.SubItems[10].Text = $WorkstationInfo.SerialNumber
  $ListViewItem.SubItems[11].Text = $WorkstationInfo.Manufacturer
  $ListViewItem.SubItems[12].Text = $WorkstationInfo.Model
  $ListViewItem.SubItems[13].Text = $WorkstationInfo.IsMobile
  $ListViewItem.SubItems[14].Text = $WorkstationInfo.Memory
  $ListViewItem.SubItems[15].Text = $WorkstationInfo.InstallDate
  $ListViewItem.SubItems[16].Text = $WorkstationInfo.LastBootUpTime
  
}
Catch [System.Management.Automation.RuntimeException]
{
  $WasSuccess = $False
  $ListViewItem.SubItems[$ErrorCol].Text = $PSItem.Message
}
Catch [System.Management.Automation.ErrorRecord]
{
  $WasSuccess = $False
  $ListViewItem.SubItems[$ErrorCol].Text = $PSItem.Exception.Message
}
Catch
{
  $WasSuccess = $False
  $ListViewItem.SubItems[$ErrorCol].Text = $PSItem.ToString()
}


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
