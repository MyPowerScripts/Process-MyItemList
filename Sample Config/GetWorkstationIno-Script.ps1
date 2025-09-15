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

# Build ListView Column Lookup Table
#
# Reference Columns by Name Incase Column Order Changes
# -----------------------------------------------------
$Columns = @{}
$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }

# ------------------------------------------------
# Check if Thread was Already Completed and Exit
#
# One Column needs to be the Status the the Thread
#  Status Messages are Customizable
# ------------------------------------------------
If ($ListViewItem.SubItems[$Columns["Job Status"]].Text -eq "Completed")
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
  $ListViewItem.SubItems[$Columns["Job Status"]].Text = "Pause"
  $ListViewItem.SubItems[$Columns["Date / Time"]].Text = [DateTime]::Now.ToString("g")
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
  $ListViewItem.SubItems[$Columns["Job Status"]].Text = "Terminated"
  $ListViewItem.SubItems[$Columns["Date / Time"]].Text = [DateTime]::Now.ToString("g")
  $ListViewItem.ImageKey = $InfoIcon
  Exit
}

# --------------------------------------------------
# Get Curent List Item
# --------------------------------------------------
$ComputerName = $ListViewItem.SubItems[0].Text

# Set Proccessing Ststus
$ListViewItem.SubItems[$Columns["Job Status"]].Text = "Processing"
$ListViewItem.SubItems[$Columns["Date / Time"]].Text = [DateTime]::Now.ToString("g")

Try
{
  $WorkstationInfo = Get-MyWorkstationInfo -ComputerName $ComputerName -Serial -Mobile
  $WasSuccess = $WorkstationInfo.Found
  
  $ListViewItem.SubItems[$Columns["On-Line"]].Text = $WorkstationInfo.Status
  $ListViewItem.SubItems[$Columns["IP Address"]].Text = $WorkstationInfo.IPAddress
  $ListViewItem.SubItems[$Columns["FQDN"]].Text = $WorkstationInfo.FQDN
  $ListViewItem.SubItems[$Columns["Domain"]].Text = $WorkstationInfo.Domain
  $ListViewItem.SubItems[$Columns["Computer Name"]].Text = $WorkstationInfo.ComputerName
  $ListViewItem.SubItems[$Columns["User Name"]].Text = $WorkstationInfo.UserName
  $ListViewItem.SubItems[$Columns["Operating System"]].Text = $WorkstationInfo.OperatingSystem
  $ListViewItem.SubItems[$Columns["Build Number"]].Text = $WorkstationInfo.BuildNumber
  $ListViewItem.SubItems[$Columns["Architecture"]].Text = $WorkstationInfo.Architecture
  $ListViewItem.SubItems[$Columns["Serial Number"]].Text = $WorkstationInfo.SerialNumber
  $ListViewItem.SubItems[$Columns["Manufacturer"]].Text = $WorkstationInfo.Manufacturer
  $ListViewItem.SubItems[$Columns["Model"]].Text = $WorkstationInfo.Model
  $ListViewItem.SubItems[$Columns["IsMobile"]].Text = $WorkstationInfo.IsMobile
  $ListViewItem.SubItems[$Columns["Memory"]].Text = $WorkstationInfo.Memory
  $ListViewItem.SubItems[$Columns["Install Date"]].Text = $WorkstationInfo.InstallDate
  $ListViewItem.SubItems[$Columns["Last Reboot"]].Text = $WorkstationInfo.LastBootUpTime
  
}
Catch [System.Management.Automation.RuntimeException]
{
  $WasSuccess = $False
  $ListViewItem.SubItems[$Columns[$Columns["Error Message"]]].Text = $PSItem.Message
}
Catch [System.Management.Automation.ErrorRecord]
{
  $WasSuccess = $False
  $ListViewItem.SubItems[$Columns[$Columns["Error Message"]]].Text = $PSItem.Exception.Message
}
Catch
{
  $WasSuccess = $False
  $ListViewItem.SubItems[$Columns["Error Message"]].Text = $PSItem.ToString()
}

# Set Final Date / Time and Update Status
$ListViewItem.SubItems[$Columns["Date / Time"]].Text = [DateTime]::Now.ToString("g")
If ($WasSuccess)
{
  # Return Success
  $ListViewItem.ImageKey = $GoodIcon
  $ListViewItem.SubItems[$Columns["Job Status"]].Text = "Completed"
}
Else
{
  # Return Success
  $ListViewItem.ImageKey = $BadIcon
  $ListViewItem.SubItems[$Columns["Job Status"]].Text = "Error"
}

Exit
