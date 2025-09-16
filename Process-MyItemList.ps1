# ----------------------------------------------------------------------------------------------------------------------
#  Script: Process-MyItemList.ps1
# ----------------------------------------------------------------------------------------------------------------------
<#
Change Log for PIL
------------------------------------------------------------------------------------------------
2.0.0.2 - Add Copy All Functions in Thread Config Dialog
          Add Double Click Variable to Edit in Thread Config Dialog
------------------------------------------------------------------------------------------------
2.0.0.1 - Initial Version
------------------------------------------------------------------------------------------------
#>

#requires -version 5.0

Using namespace System.Windows.Forms
Using namespace System.Drawing
Using namespace System.Collections
Using namespace System.Collections.Specialized

<#
  .SYNOPSIS
    Proccess My Item List
  .DESCRIPTION
    Proccess My Item List
  .PARAMETER StartColumns
    Number of Columns to Start with in the ListView (Default 12, Min 5, Max 24)
  .PARAMETER ConfigFile
    Path to Configuration File
  .PARAMETER ImportFile
    Path to Export File for Importind a PIL Data Export
  .EXAMPLE
    Process-MyItemList.ps1 -StartColumns 10 -ConfigFile "C:\Temp\MyConfig.json" -ImportFile "C:\Temp\MyImport.csv"
  .NOTES
    My Script PIL Version 1.0 by kensw on 08/27/2025
    Created with "Form Code Generator" Version 7.0.0.2
#>
[CmdletBinding(DefaultParameterSetName = "StartColumns")]
Param (
  [Parameter(Mandatory = $False, ParameterSetName = "StartColumns")]
  [ValidateRange(5, 24)]
  [uint16]$StartColumns = 12,
  [Parameter(Mandatory = $True, ParameterSetName = "ConfigFile")]
  [String]$ConfigFile,
  [Parameter(Mandatory = $False, ParameterSetName = "ConfigFile")]
  [String]$ImportFile
)

$ErrorActionPreference = "Stop"
#$ErrorActionPreference = "Continue"

# Set $VerbosePreference to 'SilentlyContinue' for Production Deployment
$VerbosePreference = "SilentlyContinue"
#$VerbosePreference = "Continue"

# Set $DebugPreference for Production Deployment
$DebugPreference = "SilentlyContinue"

# Hide Console Window Progress Bar
$ProgressPreference = "SilentlyContinue"

# Clear Previous Error Messages
$Error.Clear()

# Pre-Load Required Assemblies
[Void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[Void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# Enable Visual Styles
[System.Windows.Forms.Application]::EnableVisualStyles()

#region ******** PIL Configuration ********

#region ******** PIL Default Colors ********

Class Colors
{
  [Object]$Back
  [Object]$Fore
  [Object]$LabelFore
  [Object]$ErrorFore
  [Object]$TitleBack
  [Object]$TitleFore
  [Object]$GroupFore
  [Object]$TextBack
  [Object]$TextROBack
  [Object]$TextFore
  [Object]$TextTitle
  [Object]$TextHint
  [Object]$TextBad
  [Object]$TextWarn
  [Object]$TextGood
  [Object]$TextInfo
  [Object]$ButtonBack
  [Object]$ButtonFore
  
  Colors ([Object]$Back, [Object]$Fore, [Object]$LabelFore, [Object]$ErrorFore, [Object]$TitleBack, [Object]$TitleFore, [Object]$GroupFore, [Object]$TextBack, [Object]$TextROBack, [Object]$TextFore, [Object]$TextTitle, [Object]$TextHint, [Object]$TextBad, [Object]$TextWarn, [Object]$TextGood, [Object]$TextInfo, [Object]$ButtonBack, [Object]$ButtonFore)
  {
    $This.Back = $Back
    $This.Fore = $Fore
    $This.LabelFore = $LabelFore
    $This.ErrorFore = $ErrorFore
    $This.TitleBack = $TitleBack
    $This.TitleFore = $TitleFore
    $This.GroupFore = $GroupFore
    $This.TextBack = $TextBack
    $This.TextROBack = $TextROBack
    $This.TextFore = $TextFore
    $This.TextTitle = $TextTitle
    $This.TextHint = $TextHint
    $This.TextBad = $TextBad
    $This.TextWarn = $TextWarn
    $This.TextGood = $TextGood
    $This.TextInfo = $TextInfo
    $This.ButtonBack = $ButtonBack
    $This.ButtonFore = $ButtonFore
  }
}

#endregion ******** PIL Default Colors ********

#region ******** PIL Default Font ********

Class Font
{
  [Object]$Regular
  [Object]$Hint
  [Object]$Bold
  [Object]$Title
  [Single]$Ratio
  [Single]$Width
  [Single]$Height
  
  Font ([Object]$Regular, [Object]$Hint, [Object]$Bold, [Object]$Title, [Single]$Ratio, [Single]$Width, [Single]$Height)
  {
    $This.Regular = $Regular
    $This.Hint = $Hint
    $This.Bold = $Bold
    $This.Title = $Title
    $This.Ratio = $Ratio
    $This.Width = $Width
    $This.Height = $Height
  }
}

#endregion ******** PIL Default Font ********

#region ******** PIL MyConfig ********

Class MyConfig
{
  # Default Form Run Mode
  static [bool]$Production = $True

  static [String]$ScriptName = "Process-MyItemList"
  static [Version]$ScriptVersion = [Version]::New("2.0.0.2")
  static [String]$ScriptAuthor = "Ken Sweet"

  # Script Configuration
  static [String]$ScriptRoot = ""

  # Script Runtime Values
  static [Bool]$Is64Bit = ([IntPtr]::Size -eq 8)

  # Default Form Settings
  static [Int]$FormSpacer = 4
  static [int]$FormMinWidth = 80
  static [int]$FormMinHeight = 35

  # Default Font
  static [String]$FontFamily = "Verdana"
  static [Single]$FontSize = 10
  static [Single]$FontTitle = 1.5

  # Default Form Color Mode
  static [Bool]$DarkMode = ((Get-Itemproperty -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -ErrorAction "SilentlyContinue").AppsUseLightTheme -eq "0")

  # Form Auto Exit
  static [Int]$AutoExit = 0
  static [Int]$AutoExitMax = 0
  static [Int]$AutoExitTic = 60000

  # Administrative Rights
  static [Bool]$IsLocalAdmin = ([Security.Principal.WindowsPrincipal]::New([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  static [Bool]$IsPowerUser = ([Security.Principal.WindowsPrincipal]::New([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::PowerUser)

  # Network / Internet
  static [__ComObject]$IsConnected = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]"{DCB00C01-570F-4A9B-8D69-199FDBA5723B}"))

  # Help / Issues Uri's
  static [String]$HelpURL = "https://github.com/MyPowerScripts/Process-MyItemList"

  # Web Browser File Path's
  static [String]$EdgePath = (Get-Itemproperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe" -ErrorAction "SilentlyContinue")."(default)"
  static [String]$ChromePath = (Get-Itemproperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe" -ErrorAction "SilentlyContinue")."(default)"

  # Current DateTime Offset
  static [DateTimeOffset]$DateTimeOffset = [System.DateTimeOffset]::Now

  static [Colors]$Colors

  static [Font]$Font
}

#endregion ******** PIL MyConfig ********

# Get Script Path
if ([String]::IsNullOrEmpty($HostInvocation))
{
  [MyConfig]::ScriptRoot = [System.IO.Path]::GetDirectoryName($Script:MyInvocation.MyCommand.Path)
}
else
{
  [MyConfig]::ScriptRoot = [System.IO.Path]::GetDirectoryName($HostInvocation.MyCommand.Path)
}

#region ******** PIL Default Colors ********

If ([MyConfig]::DarkMode)
{
  [MyConfig]::Colors = [Colors]::New(
    [System.Drawing.Color]::FromArgb(48, 48, 48), # Back
    [System.Drawing.Color]::DodgerBlue, # Fore [System.Drawing.Color]::LightCoral
    [System.Drawing.Color]::WhiteSmoke, # LabelForr
    [System.Drawing.Color]::Red, # ErrorFoer
    [System.Drawing.Color]::DarkGray, # TitleFore
    [System.Drawing.Color]::Black, # TitleBack
    [System.Drawing.Color]::WhiteSmoke, # GroupFore
    [System.Drawing.Color]::Gainsboro, # TextBack
    [System.Drawing.Color]::DarkGray, # TextROBack
    [System.Drawing.Color]::Black, #TextFore
    [System.Drawing.Color]::Navy, # TextTitle
    [System.Drawing.Color]::Gray, # TextHint
    [System.Drawing.Color]::FireBrick, # TextBad
    [System.Drawing.Color]::Sienna, # TextWarn
    [System.Drawing.Color]::ForestGreen, # TextGood
    [System.Drawing.Color]::CornflowerBlue, # TextInfo
    [System.Drawing.Color]::DarkGray, # ButtonBack
    [System.Drawing.Color]::Black # ButtonFore
  )
}
Else
{
  [MyConfig]::Colors = [Colors]::New(
    [System.Drawing.Color]::WhiteSmoke, # Back
    [System.Drawing.Color]::Navy, # Fore
    [System.Drawing.Color]::Black, # LabelFor
    [System.Drawing.Color]::Red, # ErrorFoer
    [System.Drawing.Color]::LightBlue, # TitleFore
    [System.Drawing.Color]::Navy, # TitleBack
    [System.Drawing.Color]::Navy, # GroupFore
    [System.Drawing.Color]::White, # TextBack
    [System.Drawing.Color]::Gainsboro, # TextROBack
    [System.Drawing.Color]::Black, # TextFore
    [System.Drawing.Color]::Navy, # TextTitle
    [System.Drawing.Color]::Gray, # TextHint
    [System.Drawing.Color]::FireBrick, #TextBad
    [System.Drawing.Color]::Sienna, # TextWarn
    [System.Drawing.Color]::ForestGreen, # TextGood
    [System.Drawing.Color]::CornflowerBlue, # TextInfo
    [System.Drawing.Color]::Gainsboro, # ButtonBack
    [System.Drawing.Color]::Navy) # ButtonFore
}

#region Default Colors
<#
[MyConfig]::Colors = [Colors]::New(
  [System.Drawing.SystemColors]::Control, # Back
  [System.Drawing.SystemColors]::ControlText, # Fore
  [System.Drawing.SystemColors]::ControlText, # LabelFore
  [System.Drawing.SystemColors]::ControlText, # ErrorFore
  [System.Drawing.SystemColors]::ControlText, # TitleFore
  [System.Drawing.SystemColors]::Control, # TitleBack
  [System.Drawing.SystemColors]::ControlText, # GroupFore
  [System.Drawing.SystemColors]::Window, # #TextBack
  [System.Drawing.SystemColors]::Window, # TextROBack
  [System.Drawing.SystemColors]::WindowText, # TextFore
  [System.Drawing.SystemColors]::WindowText, # TextTitle
  [System.Drawing.SystemColors]::GrayText, # TextHint
  [System.Drawing.SystemColors]::WindowText, # TextBad
  [System.Drawing.SystemColors]::WindowText, # TextWarn
  [System.Drawing.SystemColors]::WindowText, # TextGood
  [System.Drawing.SystemColors]::WindowText, # TextInfo
  [System.Drawing.SystemColors]::Control, # ButtonBack
  [System.Drawing.SystemColors]::ControlText # ButtonFore
)
#>
#endregion Default Colors

#endregion ******** PIL Default Colors ********

#region ******** PIL Default Font ********

$MonitorSize = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize
:FontCheck Do
{
  $Bold = [System.Drawing.Font]::New([MyConfig]::FontFamily, [MyConfig]::FontSize, [System.Drawing.FontStyle]::Bold)
  $Graphics = [System.Drawing.Graphics]::FromHwnd([System.IntPtr]::Zero)
  $MeasureString = $Graphics.MeasureString("X", $Bold)
  If (($MonitorSize.Width -le ([MyConfig]::FormMinWidth * [Math]::Floor($MeasureString.Width))) -or ($MonitorSize.Height -le ([MyConfig]::FormMinHeight * [Math]::Floor($MeasureString.Height))))
  {
    [MyConfig]::FontSize = [MyConfig]::FontSize - .1
  }
  Else
  {
    break FontCheck
  }
}
While ($True)
$Regular = [System.Drawing.Font]::New([MyConfig]::FontFamily, [MyConfig]::FontSize, [System.Drawing.FontStyle]::Regular)
$Hint = [System.Drawing.Font]::New([MyConfig]::FontFamily, [MyConfig]::FontSize, [System.Drawing.FontStyle]::Italic)
$Title = [System.Drawing.Font]::New([MyConfig]::FontFamily, ([MyConfig]::FontSize * [MyConfig]::FontTitle), [System.Drawing.FontStyle]::Bold)
[MyConfig]::Font = [Font]::New($Regular, $Hint, $Bold, $Title, ($Graphics.DpiX / 96), ([Math]::Floor($MeasureString.Width)), ([Math]::Ceiling($MeasureString.Height)))
$MonitorSize = $Null
$Regular = $Null
$Hint = $Null
$Bold = $Null
$Title = $Null
$MeasureString = $Null
$Graphics.Dispose()
$Graphics = $Null

#endregion ******** PIL Default Font ********

#endregion ******** PIL Configuration  ********

#region ******** PIL Custom Config Classes ********

#region Class PILModule
Class PILModule
{
  [String]$Location
  [String]$Name
  [Version]$Version
  
  PILModule ([String]$Location, [String]$Name, [String]$Version)
  {
    $This.Location = $Location
    $This.Name = $Name
    $This.Version = [System.Version]::New($Version)
  }
}
#endregion Class PILModule

#region Class PILFunction
Class PILFunction
{
  [String]$Name
  [String]$ScriptBlock
  
  PILFunction ([String]$Name, [String]$ScriptBlock)
  {
    $This.Name = $Name
    $This.ScriptBlock = $ScriptBlock
  }
}
#endregion Class PILFunction

#region Class PILVariable
Class PILVariable
{
  [String]$Name
  [String]$Value
  PILVariable ([String]$Name, [String]$Value)
  {
    $This.Name = $Name
    $This.Value = $Value
  }
}
#endregion Class PILVariable

#region Class PILThreadConfig
Class PILThreadConfig
{
  [ArrayList]$ColumnNames = [ArrayList]::New()
  [OrderedDictionary]$Modules = [OrderedDictionary]::New()
  [HashTable]$Functions = [HashTable]::New()
  [HashTable]$Variables = [HashTable]::New()
  [uint16]$ThreadCount = 8
  [String]$ThreadScript = $Null
  
  PILThreadConfig ([uint16]$MaxColumns)
  {
    $This.ColumnNames.Clear()
    For ($I = 0; $I -lt $MaxColumns; $I++)
    {
      [Void]$This.ColumnNames.Add(("Column Name {0:00}" -f $I))
    }
  }
  
  [Void] SetColumnNames ([String[]]$ColumnNames)
  {
    $Max = $ColumnNames.Count
    $This.ColumnNames.Clear()
    For ($I = 0; $I -lt $Max; $I++)
    {
      [Void]$This.ColumnNames.Add($ColumnNames[$I])
    }
  }
  
  [OrderedDictionary] GetColumnNames ()
  {
    $TmpValue = [Ordered]@{ }
    
    $I = 0
    ForEach ($ColumnName In $This.ColumnNames)
    {
      [Void]$TmpValue.Add(("Column Name {0:00}" -f $I), $ColumnName)
      $I++
    }
    Return $TmpValue
  }
  
  [Void] UpdateThreadInfo ([uint16]$ThreadCount, [String]$ThreadScript)
  {
    $This.ThreadCount = $ThreadCount
    $This.ThreadScript = $ThreadScript
  }
}
#endregion Class PILThreadConfig

#endregion ******** PIL Custom Config Classes ********

#region ******** PIL Runtime Values ********

Class MyRuntime
{
  # Min/Max Number of Columns
  Static [Uint16]$MinColumns = 5
  Static [Uint16]$MaxColumns = 24
  Static [UInt16]$StartColumns = $StartColumns
  Static [UInt16]$CurrentColumns = $StartColumns
  
  Static [String[]]$ConfigProperties = @("ColumnNames", "Modules", "Variables", "Functions", "ThreadCount", "ThreadScript")
  
  # Thread Configuration
  Static [PILThreadConfig]$ThreadConfig = [PILThreadConfig]::New([MyRuntime]::CurrentColumns)
  
  # Path to Module Install Locatiosn
  Static [String]$AUModules = "$($ENV:ProgramFiles)\WindowsPowerShell\Modules"
  Static [String]$CUModules = "$([Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell\Modules"
  
  # List of Installed Modules
  Static [HashTable]$Modules = [HashTable]::New()
  
  # Loaded Functions
  Static [HashTable]$Functions = [HashTable]::New()
  
  Static [Void] UpdateTotalColumn ([Uint16]$MaxColumns)
  {
    [MyRuntime]::CurrentColumns = $MaxColumns
    [MyRuntime]::ThreadConfig = [PILThreadConfig]::New($MaxColumns)
  }
  
  Static [Void] AddPILColumn ([UInt16]$Index, [String]$AddName)
  {
    [MyRuntime]::CurrentColumns += 1
    [MyRuntime]::ThreadConfig.ColumnNames.Insert($Index, $AddName)
  }
  
  Static [Void] RemovePILComun ([String]$Index)
  {
    [MyRuntime]::CurrentColumns -= 1
    [MyRuntime]::ThreadConfig.ColumnNames.RemoveAt($Index)
  }
  
  Static [String]$ConfigName = "Unknown Configuration"
}

#endregion ******** PIL Runtime  Values ********

#region ******** PIL Demos ********

#region $SampleDemo
$SampleDemo = @'
{
  "ColumnNames": [
    "List Item",
    "Status",
    "Term/Proc Times",
    "Prompt Variable",
    "Open Mutex",
    "Synced Hash",
    "Fake Error",
    "Function Test",
    "Static Variable",
    "WasSuccess",
    "Update Time 01",
    "Update Time 02",
    "Update Time 03",
    "Update Time 04",
    "Update Time 05",
    "Update Time 06"
  ],
  "Modules": {},
  "Functions": {
    "Example-Function": {
      "Name": "Example-Function",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Example Funciton\r\n    .DESCRIPTION\r\n      Example Funciton\r\n    .PARAMETER InputValue\r\n      Required Input Value\r\n    .EXAMPLE\r\n      Example-Function -InputValue $InputValue\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding()]\r\n  Param (\r\n    [Parameter(Mandatory = $true)]\r\n    [String]$InputValue\r\n  )\r\n  Write-Verbose -Message \"Enter Function $($MyInvocation.MyCommand)\"\r\n  \r\n  Return $InputValue\r\n  \r\n  Write-Verbose -Message \"Exit Function $($MyInvocation.MyCommand)\"\r\n"
    }
  },
  "Variables": {
    "PromptVariable": {
      "Name": "PromptVariable",
      "Value": "*"
    },
    "StaticVariable": {
      "Name": "StaticVariable",
      "Value": "Static"
    }
  },
  "ThreadCount": 4,
  "ThreadScript": "\u003c#\r\n  .SYNOPSIS\r\n    Sample Runspace Pool Thread Script\r\n  .DESCRIPTION\r\n    Sample Runspace Pool Thread Script\r\n  .PARAMETER ListViewItem\r\n    ListViewItem Passed to the Thread Script\r\n\r\n    This Paramter is Required in your Thread Script\r\n  .EXAMPLE\r\n    Test-Script.ps$Columns[\"Status\"] -ListViewItem $ListViewItem\r\n  .NOTES\r\n    Sample Thread Script\r\n\r\n   -------------------------\r\n   ListViewItem Status Icons\r\n   -------------------------\r\n   $GoodIcon = Solid Green Circle\r\n   $BadIcon = Solid Red Circle\r\n   $InfoIcon = Solid Blue Circle\r\n   $CheckIcon = Checkmark\r\n   $ErrorIcon = Red X\r\n   $UpIcon = Green up Arrow \r\n   $DownIcon = Red Down Arrow\r\n\r\n#\u003e\r\n[CmdletBinding(DefaultParameterSetName = \"ByValue\")]\r\nParam (\r\n  [parameter(Mandatory = $True)]\r\n  [System.Windows.Forms.ListViewItem]$ListViewItem\r\n)\r\n\r\n# Set Preference Variables\r\n$ErrorActionPreference = \"Stop\"\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$ProgressPreference = \"SilentlyContinue\"\r\n\r\n# -----------------------------------------------------\r\n# Build ListView Column Lookup Table\r\n#\r\n# Reference Columns by Name Incase Column Order Changes\r\n# -----------------------------------------------------\r\n$Columns = @{}\r\n$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }\r\n\r\n# ------------------------------------------------\r\n# Check if Thread was Already Completed and Exit\r\n#\r\n# One Column needs to be the Status the the Thread\r\n#  Status Messages are Customizable\r\n# ------------------------------------------------\r\nIf ($ListViewItem.SubItems[$Columns[\"Status\"]].Text -eq \"Completed\")\r\n{\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  Exit\r\n}\r\n\r\n# ----------------------------------------------------\r\n# Check if Threads are Paused and Update Thread Status\r\n#\r\n# You can add Multiple Checks for Pasue if Needed\r\n# ----------------------------------------------------\r\nIf ($SyncedHash.Pause)\r\n{\r\n  # Set Paused Status\r\n  $ListViewItem.SubItems[$Columns[\"Status\"]].Text = \"Pause\"\r\n  While ($SyncedHash.Pause)\r\n  {\r\n    [System.Threading.Thread]::Sleep(100)\r\n  }\r\n}\r\n\r\n# -----------------------------------------------------\r\n# Check For Termination and Update Thread Status\r\n#\r\n# You can add Multiple Checks for Termination if Needed\r\n# -----------------------------------------------------\r\nIf ($SyncedHash.Terminate)\r\n{\r\n  # Set Terminated Status and Return\r\n  $ListViewItem.SubItems[$Columns[\"Status\"]].Text = \"Terminated\"\r\n  $ListViewItem.SubItems[$Columns[\"Term/Proc Times\"]].Text = [DateTime]::Now.ToString(\"HH:mm:ss:ffff\")\r\n  $ListViewItem.ImageKey = $InfoIcon\r\n  Exit\r\n}\r\n\r\n# Set Proccessing Ststus\r\n$ListViewItem.SubItems[$Columns[\"Status\"]].Text = \"Processing\"\r\n$ListViewItem.SubItems[$Columns[\"Term/Proc Times\"]].Text = [DateTime]::Now.ToString(\"HH:mm:ss:ffff\")\r\n$WasSuccess = $True\r\n\r\n# Set Prompt Variable\r\n$ListViewItem.SubItems[$Columns[\"Prompt Variable\"]].Text = $PromptVariable\r\n\r\n# --------------------------------------------------\r\n# Get Curent List Item\r\n#\r\n# Coulmn 0 Always has the List Item to be Proccessed\r\n# --------------------------------------------------\r\n$CurentItem = $ListViewItem.SubItems[$Columns[\"List Item\"]].Text\r\n# For Testing you can Write to the Screen\r\nWrite-Host -Object \"Processing $($CurentItem)\"\r\n\r\n# --------------------------------------------------------------\r\n# Open and wait for Mutex\r\n# \r\n# This is to Pause the Thread Script if Access a Shared Resource\r\n#   and you need toi Limit to $Columns[\"Status\"] Thread at a Time\r\n#\r\n# Using a Mutext is Optional\r\n# --------------------------------------------------------------\r\n$MyMutex = [System.Threading.Mutex]::OpenExisting($Mutex)\r\n[Void]($MyMutex.WaitOne())\r\n\r\n# Set Date / Time when Mutext was Opened\r\n$ListViewItem.SubItems[$Columns[\"Open Mutex\"]].Text = [DateTime]::Now.ToString(\"HH:mm:ss:ffff\")\r\n\r\n# Access / Update Shared Resources\r\n# $CurrentItem | Out-File -Encoding ascii -FilePath \"C:\\SharedFile.txt\"\r\n\r\n# Release Mutex\r\n$MyMutex.ReleaseMutex()\r\n\r\n# --------------------------------------------------------------------------------\r\n# The Synced HashTable has an Object Property to share information between Threads\r\n# --------------------------------------------------------------------------------\r\nIf ([String]::IsNullOrEmpty($SyncedHash.Object))\r\n{\r\n  $SyncedHash.Object = \"First\"\r\n}\r\n$ListViewItem.SubItems[$Columns[\"Synced Hash\"]].Text = $SyncedHash.Object\r\n$SyncedHash.Object = $CurentItem\r\n\r\n\r\n# Random Number Generator\r\n$Random = [System.Random]::New()\r\n\r\n# ---------------------------------------------------------\r\n# Gernate a Fake Error\r\n#\r\n# Make sure to use Error Catching to make sure thread exits\r\n# ---------------------------------------------------------\r\nTry\r\n{\r\n  Switch ($Random.Next(0, 3))\r\n  {\r\n    \"0\"\r\n    {\r\n      Throw \"This is a Fake Error!\"\r\n      Break\r\n    }\r\n    \"1\"\r\n    {\r\n      Throw \"Simulated Error!\"\r\n      Break\r\n    }\r\n    \"2\"\r\n    {\r\n      Throw \"Someing Failed!\"\r\n      Break\r\n    }\r\n    \"3\"\r\n    {\r\n      Throw \"Unknown Error!\"\r\n      Break\r\n    }\r\n  }\r\n}\r\nCatch\r\n{\r\n  # Save Error Mesage\r\n  $ListViewItem.SubItems[$Columns[\"Fake Error\"]].Text = $Error[0].Exception.Message\r\n}\r\n\r\n$ListViewItem.SubItems[$Columns[\"Function Test\"]].Text = Example-Function -InputValue \"Hello World\"\r\n$ListViewItem.SubItems[$Columns[\"Static Variable\"]].Text = $StaticVariable\r\n\r\n$RndValue = $Random.Next(0, 3)\r\nFor ($I = 10; $I -lt 16; $I++)\r\n{\r\n  $ListViewItem.SubItems[$I].Text = [DateTime]::Now.ToString(\"HH:mm:ss:ffff\")\r\n  [System.Threading.Thread]::Sleep(100)\r\n}\r\n\r\n# Random Fail Simlater\r\nIf ($RndValue -eq 0)\r\n{\r\n  $WasSuccess = $False\r\n}\r\n$ListViewItem.SubItems[$Columns[\"WasSuccess\"]].Text = $WasSuccess\r\n\r\n# Set Final Date / Time and Update Status\r\n$ListViewItem.SubItems[$Columns[\"Term/Proc Times\"]].Text = [DateTime]::Now.ToString(\"HH:mm:ss:ffff\")\r\nIf ($WasSuccess)\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status\"]].Text = \"Completed\"\r\n}\r\nElse\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $BadIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status\"]].Text = \"Error\"\r\n}\r\n\r\n# Testing Write to Screen\r\nWrite-Host -Object \"Completed $($CurentItem)\"\r\n\r\nExit\r\n\r\n\r\n"
}
'@
#endregion $SampleDemo

#region $GetWorkstationInfo
$GetWorkstationInfo = @'
{
  "ColumnNames": [
    "Workstation",
    "On-Line",
    "IP Address",
    "FQDN",
    "Domain",
    "Computer Name",
    "User Name",
    "Operating System",
    "Build Number",
    "Architecture",
    "Serial Number",
    "Manufacturer",
    "Model",
    "IsMobile",
    "Memory",
    "Install Date",
    "Last Reboot",
    "Job Status",
    "Date / Time",
    "Error Message"
  ],
  "Modules": {},
  "Functions": {
    "Get-MyWorkstationInfo": {
      "Name": "Get-MyWorkstationInfo",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Verify Remote Workstation is the Correct One\r\n    .DESCRIPTION\r\n      Verify Remote Workstation is the Correct One\r\n    .PARAMETER ComputerName\r\n      Name of the Computer to Verify\r\n    .PARAMETER Credential\r\n      Credentials to use when connecting to the Remote Computer\r\n    .PARAMETER Serial\r\n      Return Serial Number\r\n    .PARAMETER Mobile\r\n      Check if System is Desktop / Laptop\r\n    .INPUTS\r\n    .OUTPUTS\r\n    .EXAMPLE\r\n      Get-MyWorkstationInfo -ComputerName \"MyWorkstation\"\r\n    .NOTES\r\n      Original Script By Ken Sweet\r\n    .LINK\r\n  #\u003e\r\n  [CmdletBinding()]\r\n  param (\r\n    [parameter(Mandatory = $False, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]\r\n    [String[]]$ComputerName = [System.Environment]::MachineName,\r\n    [PSCredential]$Credential,\r\n    [Switch]$Serial,\r\n    [Switch]$Mobile\r\n  )\r\n  begin\r\n  {\r\n    Write-Verbose -Message \"Enter Function Get-MyWorkstationInfo\"\r\n\r\n    # Default Common Get-WmiObject Options\r\n    if ($PSBoundParameters.ContainsKey(\"Credential\"))\r\n    {\r\n      $Params = @{\r\n        \"ComputerName\" = $Null\r\n        \"Credential\"   = $Credential\r\n      }\r\n    }\r\n    else\r\n    {\r\n      $Params = @{\r\n        \"ComputerName\" = $Null\r\n      }\r\n    }\r\n  }\r\n  process\r\n  {\r\n    Write-Verbose -Message \"Enter Function Get-MyWorkstationInfo - Process\"\r\n\r\n    foreach ($Computer in $ComputerName)\r\n    {\r\n      # Start Setting Return Values as they are Found\r\n      $VerifyObject = [MyWorkstationInfo]::New($Computer)\r\n\r\n      # Validate ComputerName\r\n      if (($Computer -match \"^(([a-zA-Z]|[a-zA-Z][a-zA-Z0-9\\-]*[a-zA-Z0-9])\\.)*([A-Za-z]|[A-Za-z][A-Za-z0-9\\-]*[A-Za-z0-9])$\") -or ($Computer -match \"(?:25[0-5]|2[0-4][0-9]|1\\d{2}|[1-9]?\\d)(?:\\.(?:25[0-5]|2[0-4][0-9]|1\\d{2}|[1-9]?\\d)){3}\"))\r\n      {\r\n        try\r\n        {\r\n          # Get IP Address from DNS, you want to do all remote checks using IP rather than ComputerName.  If you connect to a computer using the wrong name Get-WmiObject will fail and using the IP Address will not\r\n          $IPAddresses = @([System.Net.Dns]::GetHostAddresses($Computer) | Where-Object -FilterScript { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } | Select-Object -ExpandProperty IPAddressToString)\r\n          :FoundMyWork foreach ($IPAddress in $IPAddresses)\r\n          {\r\n            if ([System.Net.NetworkInformation.Ping]::New().Send($IPAddress).Status -eq [System.Net.NetworkInformation.IPStatus]::Success)\r\n            {\r\n              # Set Default Parms\r\n              $Params.ComputerName = $IPAddress\r\n\r\n              # Get ComputerSystem\r\n              [Void]($MyCompData = Get-WmiObject @Params -Class Win32_ComputerSystem)\r\n              $VerifyObject.AddComputerSystem($Computer, $IPAddress, ($MyCompData.Name), ($MyCompData.PartOfDomain), ($MyCompData.Domain), ($MyCompData.Manufacturer), ($MyCompData.Model), ($MyCompData.UserName), ($MyCompData.TotalPhysicalMemory))\r\n              $MyCompData.Dispose()\r\n\r\n              # Verify Remote Computer is the Connect Computer, No need to get any more information\r\n              if ($VerifyObject.Found)\r\n              {\r\n                # Start Secondary Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer\r\n                [Void]($MyOSData = Get-WmiObject @Params -Class Win32_OperatingSystem)\r\n                $VerifyObject.AddOperatingSystem(($MyOSData.ProductType), ($MyOSData.Caption), ($MyOSData.CSDVersion), ($MyOSData.BuildNumber), ($MyOSData.Version), ($MyOSData.OSArchitecture), ([System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.LocalDateTime)), ([System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.InstallDate)), ([System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.LastBootUpTime)))\r\n                $MyOSData.Dispose()\r\n\r\n                # Optional SerialNumber Job\r\n                if ($Serial.IsPresent)\r\n                {\r\n                  # Start Optional Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer\r\n                  [Void]($MyBIOSData = Get-WmiObject @Params -Class Win32_Bios)\r\n                  $VerifyObject.AddSerialNumber($MyBIOSData.SerialNumber)\r\n                  $MyBIOSData.Dispose()\r\n                }\r\n\r\n                # Optional Mobile / ChassisType Job\r\n                if ($Mobile.IsPresent)\r\n                {\r\n                  # Start Optional Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer\r\n                  [Void]($MyChassisData = Get-WmiObject @Params -Class Win32_SystemEnclosure)\r\n                  $VerifyObject.AddIsMobile($MyChassisData.ChassisTypes)\r\n                  $MyChassisData.Dispose()\r\n                }\r\n              }\r\n              else\r\n              {\r\n                $VerifyObject.UpdateStatus(\"Wrong Workstation Name\")\r\n              }\r\n              # Beak out of Loop, Verify was a Success no need to try other IP Address if any\r\n              break FoundMyWork\r\n            }\r\n          }\r\n        }\r\n        catch\r\n        {\r\n          # Workstation Not in DNS\r\n          $VerifyObject.UpdateStatus(\"Workstation Not in DNS\")\r\n        }\r\n      }\r\n      else\r\n      {\r\n        $VerifyObject.UpdateStatus(\"Invalid Computer Name\")\r\n      }\r\n\r\n      # Set End Time and Return Results\r\n      $VerifyObject.SetEndTime()\r\n    }\r\n    Write-Verbose -Message \"Exit Function Get-MyWorkstationInfo - Process\"\r\n  }\r\n  end\r\n  {\r\n    [System.GC]::Collect()\r\n    [System.GC]::WaitForPendingFinalizers()\r\n    Write-Verbose -Message \"Exit Function Get-MyWorkstationInfo\"\r\n  }\r\n"
    }
  },
  "Variables": {},
  "ThreadCount": 4,
  "ThreadScript": "\u003c#\r\n  .SYNOPSIS\r\n    Sample Runspace Pool Thread Script\r\n  .DESCRIPTION\r\n    Sample Runspace Pool Thread Script\r\n  .PARAMETER ListViewItem\r\n    ListViewItem Passed to the Thread Script\r\n\r\n    This Paramter is Required in your Thread Script\r\n  .EXAMPLE\r\n    Test-Script.ps1 -ListViewItem $ListViewItem\r\n  .NOTES\r\n    Sample Thread Script\r\n\r\n   -------------------------\r\n   ListViewItem Status Icons\r\n   -------------------------\r\n   $GoodIcon = Solid Green Circle\r\n   $BadIcon = Solid Red Circle\r\n   $InfoIcon = Solid Blue Circle\r\n   $CheckIcon = Checkmark\r\n   $ErrorIcon = Red X\r\n   $UpIcon = Green up Arrow \r\n   $DownIcon = Red Down Arrow\r\n\r\n#\u003e\r\n[CmdletBinding()]\r\nParam (\r\n  [parameter(Mandatory = $True)]\r\n  [System.Windows.Forms.ListViewItem]$ListViewItem\r\n)\r\n\r\n# Set Preference Variables\r\n$ErrorActionPreference = \"Stop\"\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$ProgressPreference = \"SilentlyContinue\"\r\n\r\n# -----------------------------------------------------\r\n# Build ListView Column Lookup Table\r\n#\r\n# Reference Columns by Name Incase Column Order Changes\r\n# -----------------------------------------------------\r\n$Columns = @{}\r\n$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }\r\n\r\n#region class MyWorkstationInfo\r\nClass MyWorkstationInfo\r\n{\r\n  [String]$ComputerName = [Environment]::MachineName\r\n  [String]$FQDN = [Environment]::MachineName\r\n  [Bool]$Found = $False\r\n  [String]$UserName = \"\"\r\n  [String]$Domain = \"\"\r\n  [Bool]$DomainMember = $False\r\n  [int]$ProductType = 0\r\n  [String]$Manufacturer = \"\"\r\n  [String]$Model = \"\"\r\n  [Bool]$IsMobile = $False\r\n  [String]$SerialNumber = \"\"\r\n  [Long]$Memory = 0\r\n  [String]$OperatingSystem = \"\"\r\n  [String]$BuildNumber = \"\"\r\n  [String]$Version = \"\"\r\n  [String]$ServicePack = \"\"\r\n  [String]$Architecture = \"\"\r\n  [Bool]$Is64Bit = $False\r\n  [DateTime]$LocalDateTime = [DateTime]::MinValue\r\n  [DateTime]$InstallDate = [DateTime]::MinValue\r\n  [DateTime]$LastBootUpTime = [DateTime]::MinValue\r\n  [String]$IPAddress = \"\"\r\n  [String]$Status = \"Off-Line\"\r\n  [DateTime]$StartTime = [DateTime]::Now\r\n  [DateTime]$EndTime = [DateTime]::Now\r\n  \r\n  MyWorkstationInfo ([String]$ComputerName)\r\n  {\r\n    $This.ComputerName = $ComputerName.ToUpper()\r\n    $This.FQDN = $ComputerName.ToUpper()\r\n    $This.Status = \"On-Line\"\r\n  }\r\n  \r\n  [Void] AddComputerSystem ([String]$TestName, [String]$IPAddress, [String]$ComputerName, [Bool]$DomainMember, [String]$Domain, [String]$Manufacturer, [String]$Model, [String]$UserName, [Long]$Memory)\r\n  {\r\n    $This.IPAddress = $IPAddress\r\n    $This.ComputerName = \"$($ComputerName)\".ToUpper()\r\n    $This.DomainMember = $DomainMember\r\n    $This.Domain = \"$($Domain)\".ToUpper()\r\n    If ($DomainMember)\r\n    {\r\n      $This.FQDN = \"$($ComputerName).$($Domain)\".ToUpper()\r\n    }\r\n    $This.Manufacturer = $Manufacturer\r\n    $This.Model = $Model\r\n    $This.UserName = $UserName\r\n    $This.Memory = $Memory\r\n    $This.Found = ($ComputerName -eq @($TestName.Split(\".\"))[0])\r\n  }\r\n  \r\n  [Void] AddOperatingSystem ([int]$ProductType, [String]$OperatingSystem, [String]$ServicePack, [String]$BuildNumber, [String]$Version, [String]$Architecture, [DateTime]$LocalDateTime, [DateTime]$InstallDate, [DateTime]$LastBootUpTime)\r\n  {\r\n    $This.ProductType = $ProductType\r\n    $This.OperatingSystem = $OperatingSystem\r\n    $This.ServicePack = $ServicePack\r\n    $This.BuildNumber = $BuildNumber\r\n    $This.Version = $Version\r\n    $This.Architecture = $Architecture\r\n    $This.Is64Bit = ($Architecture -eq \"64-bit\")\r\n    $This.LocalDateTime = $LocalDateTime\r\n    $This.InstallDate = $InstallDate\r\n    $This.LastBootUpTime = $LastBootUpTime\r\n  }\r\n  \r\n  [Void] AddSerialNumber ([String]$SerialNumber)\r\n  {\r\n    $This.SerialNumber = $SerialNumber\r\n  }\r\n  \r\n  [Void] AddIsMobile ([Long[]]$ChassisTypes)\r\n  {\r\n    $This.IsMobile = (@(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32) -contains $ChassisTypes[0])\r\n  }\r\n  \r\n  [Void] UpdateStatus ([String]$Status)\r\n  {\r\n    $This.Status = $Status\r\n  }\r\n  \r\n  [MyWorkstationInfo] SetEndTime ()\r\n  {\r\n    $This.EndTime = [DateTime]::Now\r\n    Return $This\r\n  }\r\n  \r\n  [TimeSpan] GetRunTime ()\r\n  {\r\n    Return ($This.EndTime - $This.StartTime)\r\n  }\r\n}\r\n#endregion class MyWorkstationInfo\r\n\r\n# Build ListView Column Lookup Table\r\n#\r\n# Reference Columns by Name Incase Column Order Changes\r\n# -----------------------------------------------------\r\n$Columns = @{}\r\n$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }\r\n\r\n# ------------------------------------------------\r\n# Check if Thread was Already Completed and Exit\r\n#\r\n# One Column needs to be the Status the the Thread\r\n#  Status Messages are Customizable\r\n# ------------------------------------------------\r\nIf ($ListViewItem.SubItems[$Columns[\"Job Status\"]].Text -eq \"Completed\")\r\n{\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  Exit\r\n}\r\n\r\n# ----------------------------------------------------\r\n# Check if Threads are Paused and Update Thread Status\r\n#\r\n# You can add Multiple Checks for Pasue if Needed\r\n# ----------------------------------------------------\r\nIf ($SyncedHash.Pause)\r\n{\r\n  # Set Paused Status\r\n  $ListViewItem.SubItems[$Columns[\"Job Status\"]].Text = \"Pause\"\r\n  $ListViewItem.SubItems[$Columns[\"Date / Time\"]].Text = [DateTime]::Now.ToString(\"g\")\r\n  While ($SyncedHash.Pause)\r\n  {\r\n    [System.Threading.Thread]::Sleep(100)\r\n  }\r\n}\r\n\r\n# -----------------------------------------------------\r\n# Check For Termination and Update Thread Status\r\n#\r\n# You can add Multiple Checks for Termination if Needed\r\n# -----------------------------------------------------\r\nIf ($SyncedHash.Terminate)\r\n{\r\n  # Set Terminated Status and Return\r\n  $ListViewItem.SubItems[$Columns[\"Job Status\"]].Text = \"Terminated\"\r\n  $ListViewItem.SubItems[$Columns[\"Date / Time\"]].Text = [DateTime]::Now.ToString(\"g\")\r\n  $ListViewItem.ImageKey = $InfoIcon\r\n  Exit\r\n}\r\n\r\n# --------------------------------------------------\r\n# Get Curent List Item\r\n# --------------------------------------------------\r\n$ComputerName = $ListViewItem.SubItems[0].Text\r\n\r\n# Set Proccessing Ststus\r\n$ListViewItem.SubItems[$Columns[\"Job Status\"]].Text = \"Processing\"\r\n$ListViewItem.SubItems[$Columns[\"Date / Time\"]].Text = [DateTime]::Now.ToString(\"g\")\r\n\r\nTry\r\n{\r\n  $WorkstationInfo = Get-MyWorkstationInfo -ComputerName $ComputerName -Serial -Mobile\r\n  $WasSuccess = $WorkstationInfo.Found\r\n  \r\n  $ListViewItem.SubItems[$Columns[\"On-Line\"]].Text = $WorkstationInfo.Status\r\n  $ListViewItem.SubItems[$Columns[\"IP Address\"]].Text = $WorkstationInfo.IPAddress\r\n  $ListViewItem.SubItems[$Columns[\"FQDN\"]].Text = $WorkstationInfo.FQDN\r\n  $ListViewItem.SubItems[$Columns[\"Domain\"]].Text = $WorkstationInfo.Domain\r\n  $ListViewItem.SubItems[$Columns[\"Computer Name\"]].Text = $WorkstationInfo.ComputerName\r\n  $ListViewItem.SubItems[$Columns[\"User Name\"]].Text = $WorkstationInfo.UserName\r\n  $ListViewItem.SubItems[$Columns[\"Operating System\"]].Text = $WorkstationInfo.OperatingSystem\r\n  $ListViewItem.SubItems[$Columns[\"Build Number\"]].Text = $WorkstationInfo.BuildNumber\r\n  $ListViewItem.SubItems[$Columns[\"Architecture\"]].Text = $WorkstationInfo.Architecture\r\n  $ListViewItem.SubItems[$Columns[\"Serial Number\"]].Text = $WorkstationInfo.SerialNumber\r\n  $ListViewItem.SubItems[$Columns[\"Manufacturer\"]].Text = $WorkstationInfo.Manufacturer\r\n  $ListViewItem.SubItems[$Columns[\"Model\"]].Text = $WorkstationInfo.Model\r\n  $ListViewItem.SubItems[$Columns[\"IsMobile\"]].Text = $WorkstationInfo.IsMobile\r\n  $ListViewItem.SubItems[$Columns[\"Memory\"]].Text = $WorkstationInfo.Memory\r\n  $ListViewItem.SubItems[$Columns[\"Install Date\"]].Text = $WorkstationInfo.InstallDate\r\n  $ListViewItem.SubItems[$Columns[\"Last Reboot\"]].Text = $WorkstationInfo.LastBootUpTime\r\n  \r\n}\r\nCatch [System.Management.Automation.RuntimeException]\r\n{\r\n  $WasSuccess = $False\r\n  $ListViewItem.SubItems[$Columns[$Columns[\"Error Message\"]]].Text = $PSItem.Message\r\n}\r\nCatch [System.Management.Automation.ErrorRecord]\r\n{\r\n  $WasSuccess = $False\r\n  $ListViewItem.SubItems[$Columns[$Columns[\"Error Message\"]]].Text = $PSItem.Exception.Message\r\n}\r\nCatch\r\n{\r\n  $WasSuccess = $False\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = $PSItem.ToString()\r\n}\r\n\r\n# Set Final Date / Time and Update Status\r\n$ListViewItem.SubItems[$Columns[\"Date / Time\"]].Text = [DateTime]::Now.ToString(\"g\")\r\nIf ($WasSuccess)\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  $ListViewItem.SubItems[$Columns[\"Job Status\"]].Text = \"Completed\"\r\n}\r\nElse\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $BadIcon\r\n  $ListViewItem.SubItems[$Columns[\"Job Status\"]].Text = \"Error\"\r\n}\r\n\r\nExit\r\n\r\n\r\n"
}
'@
#endregion $GetWorkstationInfo

#region $GetDomainComputer
$GetDomainComputer = @'
{
  "ColumnNames": [
    "ComputerName",
    "Domain",
    "Last Logon",
    "PwdLastSet",
    "UserAccountControl",
    "Locked Out",
    "Disabled",
    "OperatingSystem",
    "OperatingSystemVersion",
    "DistinguishedName",
    "CanonicalName",
    "Date/Time",
    "Status Message",
    "Error Message"
  ],
  "Modules": {},
  "Functions": {
    "Get-MyADObject": {
      "Name": "Get-MyADObject",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Searches Active Directory and returns an AD SearchResultCollection.\r\n    .DESCRIPTION\r\n      Performs a search in Active Directory using the specified LDAP filter and returns a SearchResultCollection. \r\n      Supports specifying search root, server, credentials, properties to load, sorting, and paging options.\r\n    .PARAMETER LDAPFilter\r\n      The LDAP filter string to use for the search. Defaults to (objectClass=*).\r\n    .PARAMETER PageSize\r\n      The number of objects to return per page. Default is 1000.\r\n    .PARAMETER SizeLimit\r\n      The maximum number of objects to return. Default is 1000.\r\n    .PARAMETER SearchRoot\r\n      The LDAP path to start the search from. Defaults to the current domain root.\r\n    .PARAMETER ServerName\r\n      The name of the domain controller or server to query. If not specified, uses the default.\r\n    .PARAMETER SearchScope\r\n      The scope of the search. Valid values are Base, OneLevel, or Subtree. Default is Subtree.\r\n    .PARAMETER Sort\r\n      The direction to sort the results. Valid values are Ascending or Descending. Default is Ascending.\r\n    .PARAMETER SortProperty\r\n      The property name to sort the results by.\r\n    .PARAMETER PropertiesToLoad\r\n      An array of property names to load for each result.\r\n    .PARAMETER Credential\r\n      The credentials to use when searching Active Directory.\r\n    .EXAMPLE\r\n      Get-MyADObject -LDAPFilter \"(objectClass=user)\" -SearchRoot \"OU=Users,DC=domain,DC=com\"\r\n      Searches for all user objects in the specified OU.\r\n    .EXAMPLE\r\n      Get-MyADObject -ServerName \"dc01.domain.com\" -PropertiesToLoad \"samaccountname\",\"mail\"\r\n      Searches using a specific domain controller and returns only the samaccountname and mail properties.\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding(DefaultParameterSetName = \"Default\")]\r\n  param (\r\n    [String]$LDAPFilter = \"(objectClass=*)\",\r\n    [Long]$PageSize = 1000,\r\n    [Long]$SizeLimit = 1000,\r\n    [String]$SearchRoot = \"LDAP://$($([ADSI]\u0027\u0027).distinguishedName)\",\r\n    [String]$ServerName,\r\n    [ValidateSet(\"Base\", \"OneLevel\", \"Subtree\")]\r\n    [System.DirectoryServices.SearchScope]$SearchScope = \"SubTree\",\r\n    [ValidateSet(\"Ascending\", \"Descending\")]\r\n    [System.DirectoryServices.SortDirection]$Sort = \"Ascending\",\r\n    [String]$SortProperty,\r\n    [String[]]$PropertiesToLoad,\r\n    [PSCredential]$Credential\r\n  )\r\n  Write-Verbose -Message \"Enter Function $($MyInvocation.MyCommand)\"\r\n\r\n  $MySearcher = [System.DirectoryServices.DirectorySearcher]::New($LDAPFilter, $PropertiesToLoad, $SearchScope)\r\n\r\n  $MySearcher.PageSize = $PageSize\r\n  $MySearcher.SizeLimit = $SizeLimit\r\n\r\n  $TempSearchRoot = $SearchRoot.ToUpper()\r\n  switch -regex ($TempSearchRoot)\r\n  {\r\n    \"(?:LDAP|GC)://*\"\r\n    {\r\n      if ($PSBoundParameters.ContainsKey(\"ServerName\"))\r\n      {\r\n        $MySearchRoot = $TempSearchRoot -replace \"(?\u003cLG\u003e(?:LDAP|GC)://)(?:[\\w\\d\\.-]+/)?(?\u003cDN\u003e.+)\", \"`${LG}$($ServerName)/`${DN}\"\r\n      }\r\n      else\r\n      {\r\n        $MySearchRoot = $TempSearchRoot\r\n      }\r\n      break\r\n    }\r\n    default\r\n    {\r\n      if ($PSBoundParameters.ContainsKey(\"ServerName\"))\r\n      {\r\n        $MySearchRoot = \"LDAP://$($ServerName)/$($TempSearchRoot)\"\r\n      }\r\n      else\r\n      {\r\n        $MySearchRoot = \"LDAP://$($TempSearchRoot)\"\r\n      }\r\n      break\r\n    }\r\n  }\r\n\r\n  if ($PSBoundParameters.ContainsKey(\"Credential\"))\r\n  {\r\n    $MySearcher.SearchRoot = [System.DirectoryServices.DirectoryEntry]::New($MySearchRoot, ($Credential.UserName), (($Credential.GetNetworkCredential()).Password))\r\n  }\r\n  else\r\n  {\r\n    $MySearcher.SearchRoot = [System.DirectoryServices.DirectoryEntry]::New($MySearchRoot)\r\n  }\r\n\r\n  if ($PSBoundParameters.ContainsKey(\"SortProperty\"))\r\n  {\r\n    $MySearcher.Sort.PropertyName = $SortProperty\r\n    $MySearcher.Sort.Direction = $Sort\r\n  }\r\n\r\n  $MySearcher.FindAll()\r\n\r\n  $MySearcher.Dispose()\r\n  $MySearcher = $Null\r\n  $MySearchRoot = $Null\r\n  $TempSearchRoot = $Null\r\n\r\n  Write-Verbose -Message \"Exit Function $($MyInvocation.MyCommand)\"\r\n"
    },
    "Get-MyADForest": {
      "Name": "Get-MyADForest",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Gets information about an Active Directory Forest.\r\n    .DESCRIPTION\r\n      Retrieves the Active Directory Forest object either for the current forest or for a specified forest name.\r\n    .PARAMETER Name\r\n      The name of the Active Directory forest to retrieve. This parameter is mandatory when using the \"Name\" parameter set.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADForest\r\n      Retrieves the current Active Directory forest.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADForest -Name \"contoso.com\"\r\n      Retrieves the Active Directory forest with the name \"contoso.com\".\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding(DefaultParameterSetName = \"Current\")]\r\n  param (\r\n    [parameter(Mandatory = $True, ParameterSetName = \"Name\")]\r\n    [String]$Name\r\n  )\r\n  Write-Verbose -Message \"Enter Function $($MyInvocation.MyCommand)\"\r\n\r\n  switch ($PSCmdlet.ParameterSetName)\r\n  {\r\n    \"Name\"\r\n    {\r\n      $DirectoryContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Forest\r\n      $DirectoryContext = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::New($DirectoryContextType, $Name)\r\n      [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($DirectoryContext)\r\n      $DirectoryContext = $Null\r\n      $DirectoryContextType = $Null\r\n      break\r\n    }\r\n    \"Current\"\r\n    {\r\n      [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()\r\n      break\r\n    }\r\n  }\r\n\r\n  Write-Verbose -Message \"Exit Function $($MyInvocation.MyCommand)\"\r\n"
    },
    "Get-MyADDomain": {
      "Name": "Get-MyADDomain",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Gets information about an Active Directory Domain.\r\n    .DESCRIPTION\r\n      Retrieves the Active Directory Domain object either for the current domain, a specified domain name, or the domain associated with the local computer.\r\n    .PARAMETER Name\r\n      The name of the Active Directory domain to retrieve. This parameter is mandatory when using the \"Name\" parameter set.\r\n    .PARAMETER Computer\r\n      Switch parameter. If specified, retrieves the Active Directory domain associated with the local computer. This parameter is mandatory when using the \"Computer\" parameter set.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADDomain\r\n      Retrieves the current Active Directory domain.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADDomain -Computer\r\n      Retrieves the Active Directory domain associated with the local computer.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADDomain -Name \"contoso.com\"\r\n      Retrieves the Active Directory domain with the name \"contoso.com\".\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding(DefaultParameterSetName = \"Current\")]\r\n  param (\r\n    [parameter(Mandatory = $True, ParameterSetName = \"Name\")]\r\n    [String]$Name,\r\n    [parameter(Mandatory = $True, ParameterSetName = \"Computer\")]\r\n    [Switch]$Computer\r\n  )\r\n  Write-Verbose -Message \"Enter Function $($MyInvocation.MyCommand)\"\r\n\r\n  switch ($PSCmdlet.ParameterSetName)\r\n  {\r\n    \"Name\"\r\n    {\r\n      $DirectoryContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain\r\n      $DirectoryContext = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::New($DirectoryContextType, $Name)\r\n      [System.DirectoryServices.ActiveDirectory.Domian]::GetDomain($DirectoryContext)\r\n      $DirectoryContext = $Null\r\n      $DirectoryContextType = $Null\r\n      break\r\n    }\r\n    \"Computer\"\r\n    {\r\n      [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()\r\n      break\r\n    }\r\n    \"Current\"\r\n    {\r\n      [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()\r\n      break\r\n    }\r\n  }\r\n\r\n  Write-Verbose -Message \"Exit Function $($MyInvocation.MyCommand)\"\r\n"
    }
  },
  "Variables": {
    "ADDomain": { "Name": "ADDomain", "Value": "Current" },
    "ADForest": { "Name": "ADForest", "Value": "Current" }
  },
  "ThreadCount": 4,
  "ThreadScript": "\u003c#\r\n  .SYNOPSIS\r\n    Sample Runspace Pool Thread Script\r\n  .DESCRIPTION\r\n    Sample Runspace Pool Thread Script\r\n  .PARAMETER ListViewItem\r\n    ListViewItem Passed to the Thread Script\r\n\r\n    This Paramter is Required in your Thread Script\r\n  .EXAMPLE\r\n    Test-Script.ps1 -ListViewItem $ListViewItem\r\n  .NOTES\r\n    Sample Thread Script\r\n\r\n   -------------------------\r\n   ListViewItem Status Icons\r\n   -------------------------\r\n   $GoodIcon = Solid Green Circle\r\n   $BadIcon = Solid Red Circle\r\n   $InfoIcon = Solid Blue Circle\r\n   $CheckIcon = Checkmark\r\n   $ErrorIcon = Red X\r\n   $UpIcon = Green up Arrow \r\n   $DownIcon = Red Down Arrow\r\n\r\n#\u003e\r\n[CmdletBinding()]\r\nParam (\r\n  [parameter(Mandatory = $True)]\r\n  [System.Windows.Forms.ListViewItem]$ListViewItem\r\n)\r\n\r\n# Set Preference Variables\r\n$ErrorActionPreference = \"Stop\"\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$ProgressPreference = \"SilentlyContinue\"\r\n\r\n# -----------------------------------------------------\r\n# Build ListView Column Lookup Table\r\n#\r\n# Reference Columns by Name Incase Column Order Changes\r\n# -----------------------------------------------------\r\n$Columns = @{}\r\n$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }\r\n\r\n# ------------------------------------------------\r\n# Check if Thread was Already Completed and Exit\r\n# ------------------------------------------------\r\nIf ($ListViewItem.SubItems[$Columns[\"Status Message\"]].Text -eq \"Completed\")\r\n{\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  Exit\r\n}\r\n\r\n# ----------------------------------------------------\r\n# Check if Threads are Paused and Update Thread Status\r\n# ----------------------------------------------------\r\nIf ($SyncedHash.Pause)\r\n{\r\n  # Set Paused Status\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Pause\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  While ($SyncedHash.Pause)\r\n  {\r\n    [System.Threading.Thread]::Sleep(100)\r\n  }\r\n}\r\n\r\n# -----------------------------------------------------\r\n# Check For Termination and Update Thread Status\r\n# -----------------------------------------------------\r\nIf ($SyncedHash.Terminate)\r\n{\r\n  # Set Terminated Status and Exit Thread\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Terminated\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  $ListViewItem.ImageKey = $InfoIcon\r\n  Exit\r\n}\r\n\r\n# Sucess Default Exit Status\r\n$WasSuccess = $True\r\n$ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Processing\"\r\n$ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n$ComputerName = $ListViewItem.SubItems[$Columns[\"ComputerName\"]].Text\r\n\r\nTry\r\n{\r\n  # Get Current Domain / Forest\r\n  if ($ADForest -eq \"Domain\")\r\n  {\r\n    if ($ADDomain -eq \"Current\")\r\n    {\r\n      $GetADDomain = Get-MyADDomain -ErrorAction SilentlyContinue\r\n    }\r\n    else\r\n    {\r\n      $GetADDomain = Get-MyADDomain -Name $ADDomain -ErrorAction SilentlyContinue\r\n    }\r\n    if (-not [String]::IsNullOrEmpty($GetADDomain.Name))\r\n    {\r\n      $SearchRoot = \"LDAP://$(\"dc=$(($GetADDomain.Name -split \u0027\\.\u0027) -join \u0027,dc=\u0027)\")\"\r\n      $GetADDomain.Dispose()\r\n    }\r\n  }\r\n  else\r\n  {\r\n    if ($ADDomain -eq \"Current\")\r\n    {\r\n      $GetADForget = Get-MyADForest -ErrorAction SilentlyContinue\r\n    }\r\n    else\r\n    {\r\n      $GetADForget = Get-MyADForest -Name $ADForest -ErrorAction SilentlyContinue\r\n    }\r\n    if (-not [String]::IsNullOrEmpty($GetADForget.Name))\r\n    {\r\n      $SearchRoot = \"GC://$($GetADForget.Name)\"\r\n      $GetADForget.Dispose()\r\n    }\r\n  }\r\n  \r\n  # Check Domain / Forest Found\r\n  if ([String]::IsNullOrEmpty($SearchRoot))\r\n  {\r\n    $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"Unable to Get Current AD Domain / Forest\"\r\n    $WasSuccess = $False\r\n  }\r\n  else\r\n  {\r\n    $LDAPFilter = \"(\u0026(objectClass=user)(objectCategory=computer)(sAMAccountType=805306369)(cn={0}))\" -f $ComputerName\r\n    $PropertiesToLoad = @(\"name\", \"canonicalName\", \"lastLogonTimestamp\", \"pwdLastSet\", \"userAccountControl\", \"OperatingSystem\", \"OperatingSystemVersion\", \"distinguishedName\")\r\n    $ADObject = Get-MyADObject -SearchRoot $SearchRoot -SearchScope Subtree -LDAPFilter $LDAPFilter -PropertiesToLoad $PropertiesToLoad -ErrorAction SilentlyContinue | Select-Object -First 1\r\n    if ([String]::IsNullOrEmpty($ADObject.Path))\r\n    {\r\n      $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"Computer Not Found in AD Forest\"\r\n      $WasSuccess = $False\r\n    }\r\n    else\r\n    {\r\n      # CanonicalName\r\n      $CanonicalName = $ADObject.Properties[\"canonicalName\"][0]\r\n      $ListViewItem.SubItems[$Columns[\"canonicalName\"]].Text = $CanonicalName\r\n      \r\n      # distinguishedName\r\n      $ListViewItem.SubItems[$Columns[\"distinguishedName\"]].Text = $ADObject.Properties[\"distinguishedName\"][0]\r\n      \r\n      # Domain\r\n      $Domain = $CanonicalName -split \"/\" | Select-Object -First 1\r\n      $ListViewItem.SubItems[$Columns[\"Domain\"]].Text = $Domain\r\n      \r\n      # Zero Hour\r\n      $ZeroHour = [DateTime]::New(1601, 1, 1, 0, 0, 0)\r\n      \r\n      # Last Logon TimeStamp\r\n      $LastLogonTimestamp = $ADObject.Properties[\"LastLogonTimestamp\"][0]\r\n      $LastLogonTimeStampDate = $ZeroHour.AddTicks($LastLogonTimestamp)\r\n      $ListViewItem.SubItems[$Columns[\"Last Logon\"]].Text = $LastLogonTimeStampDate.ToString(\"G\")\r\n      \r\n      # Password Last Set\r\n      $PwdLastSet = $ADObject.Properties[\"pwdLastSet\"][0]\r\n      $PwdLastSetDate = $ZeroHour.AddTicks($PwdLastSet)\r\n      $ListViewItem.SubItems[$Columns[\"PwdLastSet\"]].Text = $PwdLastSetDate\r\n      \r\n      # User Account Control Flags\r\n      $UserAccountControl = $ADObject.Properties[\"userAccountControl\"][0]\r\n      $ListViewItem.SubItems[$Columns[\"UserAccountControl\"]].Text = $UserAccountControl\r\n      $ListViewItem.SubItems[$Columns[\"Locked Out\"]].Text = (($UserAccountControl -band 16) -ne 0)\r\n      $ListViewItem.SubItems[$Columns[\"Disabled\"]].Text = (($UserAccountControl -band 2) -ne 0)\r\n      \r\n      # Operating System\r\n      if ($ADForest -eq \"Domain\")\r\n      {\r\n        $ListViewItem.SubItems[$Columns[\"operatingSystem\"]].Text = $ADObject.Properties[\"operatingSystem\"][0]\r\n        $ListViewItem.SubItems[$Columns[\"operatingSystemVersion\"]].Text = $ADObject.Properties[\"operatingSystemVersion\"][0]\r\n      }\r\n      else\r\n      {\r\n        $ListViewItem.SubItems[$Columns[\"operatingSystem\"]].Text = \"Domain Only\"\r\n        $ListViewItem.SubItems[$Columns[\"operatingSystemVersion\"]].Text = \"Domain Only\"\r\n      }\r\n    }\r\n  }\r\n}\r\nCatch\r\n{\r\n  # Set Error Message / Thread Failed\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = $PSItem.ToString()\r\n  $WasSuccess = $False\r\n}\r\n\r\n\r\n# Set Final Date / Time and Update Status\r\n$ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\nIf ($WasSuccess)\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Completed\"\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"\"\r\n}\r\nElse\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $BadIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Error\"\r\n}\r\n\r\nExit\r\n\r\n"
}
'@
#endregion $GetDomainComputer

#region $GetDomainUser
$GetDomainUser = @'
{
  "ColumnNames": [
    "UserName",
    "Domain",
    "UserPrincipalName",
    "E-Mail",
    "Last Logon",
    "PwdLastSet",
    "UserAccountControl",
    "PwdNoChange",
    "PwdNoExpire",
    "PwdExpired",
    "Locked out",
    "Disabled",
    "DistinguishedName",
    "CanonicalName",
    "Date/Time",
    "Status Message",
    "Error Message"
  ],
  "Modules": {},
  "Functions": {
    "Get-MyADObject": {
      "Name": "Get-MyADObject",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Searches Active Directory and returns an AD SearchResultCollection.\r\n    .DESCRIPTION\r\n      Performs a search in Active Directory using the specified LDAP filter and returns a SearchResultCollection. \r\n      Supports specifying search root, server, credentials, properties to load, sorting, and paging options.\r\n    .PARAMETER LDAPFilter\r\n      The LDAP filter string to use for the search. Defaults to (objectClass=*).\r\n    .PARAMETER PageSize\r\n      The number of objects to return per page. Default is 1000.\r\n    .PARAMETER SizeLimit\r\n      The maximum number of objects to return. Default is 1000.\r\n    .PARAMETER SearchRoot\r\n      The LDAP path to start the search from. Defaults to the current domain root.\r\n    .PARAMETER ServerName\r\n      The name of the domain controller or server to query. If not specified, uses the default.\r\n    .PARAMETER SearchScope\r\n      The scope of the search. Valid values are Base, OneLevel, or Subtree. Default is Subtree.\r\n    .PARAMETER Sort\r\n      The direction to sort the results. Valid values are Ascending or Descending. Default is Ascending.\r\n    .PARAMETER SortProperty\r\n      The property name to sort the results by.\r\n    .PARAMETER PropertiesToLoad\r\n      An array of property names to load for each result.\r\n    .PARAMETER Credential\r\n      The credentials to use when searching Active Directory.\r\n    .EXAMPLE\r\n      Get-MyADObject -LDAPFilter \"(objectClass=user)\" -SearchRoot \"OU=Users,DC=domain,DC=com\"\r\n      Searches for all user objects in the specified OU.\r\n    .EXAMPLE\r\n      Get-MyADObject -ServerName \"dc01.domain.com\" -PropertiesToLoad \"samaccountname\",\"mail\"\r\n      Searches using a specific domain controller and returns only the samaccountname and mail properties.\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding(DefaultParameterSetName = \"Default\")]\r\n  param (\r\n    [String]$LDAPFilter = \"(objectClass=*)\",\r\n    [Long]$PageSize = 1000,\r\n    [Long]$SizeLimit = 1000,\r\n    [String]$SearchRoot = \"LDAP://$($([ADSI]\u0027\u0027).distinguishedName)\",\r\n    [String]$ServerName,\r\n    [ValidateSet(\"Base\", \"OneLevel\", \"Subtree\")]\r\n    [System.DirectoryServices.SearchScope]$SearchScope = \"SubTree\",\r\n    [ValidateSet(\"Ascending\", \"Descending\")]\r\n    [System.DirectoryServices.SortDirection]$Sort = \"Ascending\",\r\n    [String]$SortProperty,\r\n    [String[]]$PropertiesToLoad,\r\n    [PSCredential]$Credential\r\n  )\r\n  Write-Verbose -Message \"Enter Function $($MyInvocation.MyCommand)\"\r\n\r\n  $MySearcher = [System.DirectoryServices.DirectorySearcher]::New($LDAPFilter, $PropertiesToLoad, $SearchScope)\r\n\r\n  $MySearcher.PageSize = $PageSize\r\n  $MySearcher.SizeLimit = $SizeLimit\r\n\r\n  $TempSearchRoot = $SearchRoot.ToUpper()\r\n  switch -regex ($TempSearchRoot)\r\n  {\r\n    \"(?:LDAP|GC)://*\"\r\n    {\r\n      if ($PSBoundParameters.ContainsKey(\"ServerName\"))\r\n      {\r\n        $MySearchRoot = $TempSearchRoot -replace \"(?\u003cLG\u003e(?:LDAP|GC)://)(?:[\\w\\d\\.-]+/)?(?\u003cDN\u003e.+)\", \"`${LG}$($ServerName)/`${DN}\"\r\n      }\r\n      else\r\n      {\r\n        $MySearchRoot = $TempSearchRoot\r\n      }\r\n      break\r\n    }\r\n    default\r\n    {\r\n      if ($PSBoundParameters.ContainsKey(\"ServerName\"))\r\n      {\r\n        $MySearchRoot = \"LDAP://$($ServerName)/$($TempSearchRoot)\"\r\n      }\r\n      else\r\n      {\r\n        $MySearchRoot = \"LDAP://$($TempSearchRoot)\"\r\n      }\r\n      break\r\n    }\r\n  }\r\n\r\n  if ($PSBoundParameters.ContainsKey(\"Credential\"))\r\n  {\r\n    $MySearcher.SearchRoot = [System.DirectoryServices.DirectoryEntry]::New($MySearchRoot, ($Credential.UserName), (($Credential.GetNetworkCredential()).Password))\r\n  }\r\n  else\r\n  {\r\n    $MySearcher.SearchRoot = [System.DirectoryServices.DirectoryEntry]::New($MySearchRoot)\r\n  }\r\n\r\n  if ($PSBoundParameters.ContainsKey(\"SortProperty\"))\r\n  {\r\n    $MySearcher.Sort.PropertyName = $SortProperty\r\n    $MySearcher.Sort.Direction = $Sort\r\n  }\r\n\r\n  $MySearcher.FindAll()\r\n\r\n  $MySearcher.Dispose()\r\n  $MySearcher = $Null\r\n  $MySearchRoot = $Null\r\n  $TempSearchRoot = $Null\r\n\r\n  Write-Verbose -Message \"Exit Function $($MyInvocation.MyCommand)\"\r\n"
    },
    "Get-MyADForest": {
      "Name": "Get-MyADForest",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Gets information about an Active Directory Forest.\r\n    .DESCRIPTION\r\n      Retrieves the Active Directory Forest object either for the current forest or for a specified forest name.\r\n    .PARAMETER Name\r\n      The name of the Active Directory forest to retrieve. This parameter is mandatory when using the \"Name\" parameter set.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADForest\r\n      Retrieves the current Active Directory forest.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADForest -Name \"contoso.com\"\r\n      Retrieves the Active Directory forest with the name \"contoso.com\".\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding(DefaultParameterSetName = \"Current\")]\r\n  param (\r\n    [parameter(Mandatory = $True, ParameterSetName = \"Name\")]\r\n    [String]$Name\r\n  )\r\n  Write-Verbose -Message \"Enter Function $($MyInvocation.MyCommand)\"\r\n\r\n  switch ($PSCmdlet.ParameterSetName)\r\n  {\r\n    \"Name\"\r\n    {\r\n      $DirectoryContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Forest\r\n      $DirectoryContext = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::New($DirectoryContextType, $Name)\r\n      [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($DirectoryContext)\r\n      $DirectoryContext = $Null\r\n      $DirectoryContextType = $Null\r\n      break\r\n    }\r\n    \"Current\"\r\n    {\r\n      [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()\r\n      break\r\n    }\r\n  }\r\n\r\n  Write-Verbose -Message \"Exit Function $($MyInvocation.MyCommand)\"\r\n"
    },
    "Get-MyADDomain": {
      "Name": "Get-MyADDomain",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Gets information about an Active Directory Domain.\r\n    .DESCRIPTION\r\n      Retrieves the Active Directory Domain object either for the current domain, a specified domain name, or the domain associated with the local computer.\r\n    .PARAMETER Name\r\n      The name of the Active Directory domain to retrieve. This parameter is mandatory when using the \"Name\" parameter set.\r\n    .PARAMETER Computer\r\n      Switch parameter. If specified, retrieves the Active Directory domain associated with the local computer. This parameter is mandatory when using the \"Computer\" parameter set.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADDomain\r\n      Retrieves the current Active Directory domain.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADDomain -Computer\r\n      Retrieves the Active Directory domain associated with the local computer.\r\n    .EXAMPLE\r\n      PS C:\\\u003e Get-MyADDomain -Name \"contoso.com\"\r\n      Retrieves the Active Directory domain with the name \"contoso.com\".\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding(DefaultParameterSetName = \"Current\")]\r\n  param (\r\n    [parameter(Mandatory = $True, ParameterSetName = \"Name\")]\r\n    [String]$Name,\r\n    [parameter(Mandatory = $True, ParameterSetName = \"Computer\")]\r\n    [Switch]$Computer\r\n  )\r\n  Write-Verbose -Message \"Enter Function $($MyInvocation.MyCommand)\"\r\n\r\n  switch ($PSCmdlet.ParameterSetName)\r\n  {\r\n    \"Name\"\r\n    {\r\n      $DirectoryContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain\r\n      $DirectoryContext = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::New($DirectoryContextType, $Name)\r\n      [System.DirectoryServices.ActiveDirectory.Domian]::GetDomain($DirectoryContext)\r\n      $DirectoryContext = $Null\r\n      $DirectoryContextType = $Null\r\n      break\r\n    }\r\n    \"Computer\"\r\n    {\r\n      [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()\r\n      break\r\n    }\r\n    \"Current\"\r\n    {\r\n      [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()\r\n      break\r\n    }\r\n  }\r\n\r\n  Write-Verbose -Message \"Exit Function $($MyInvocation.MyCommand)\"\r\n"
    }
  },
  "Variables": {
    "ADDomain": { "Name": "ADDomain", "Value": "Current" },
    "ADForest": { "Name": "ADForest", "Value": "Current" }
  },
  "ThreadCount": 4,
  "ThreadScript": "\u003c#\r\n  .SYNOPSIS\r\n    Sample Runspace Pool Thread Script\r\n  .DESCRIPTION\r\n    Sample Runspace Pool Thread Script\r\n  .PARAMETER ListViewItem\r\n    ListViewItem Passed to the Thread Script\r\n\r\n    This Paramter is Required in your Thread Script\r\n  .EXAMPLE\r\n    Test-Script.ps1 -ListViewItem $ListViewItem\r\n  .NOTES\r\n    Sample Thread Script\r\n\r\n   -------------------------\r\n   ListViewItem Status Icons\r\n   -------------------------\r\n   $GoodIcon = Solid Green Circle\r\n   $BadIcon = Solid Red Circle\r\n   $InfoIcon = Solid Blue Circle\r\n   $CheckIcon = Checkmark\r\n   $ErrorIcon = Red X\r\n   $UpIcon = Green up Arrow \r\n   $DownIcon = Red Down Arrow\r\n\r\n#\u003e\r\n[CmdletBinding()]\r\nParam (\r\n  [parameter(Mandatory = $True)]\r\n  [System.Windows.Forms.ListViewItem]$ListViewItem\r\n)\r\n\r\n# Set Preference Variables\r\n$ErrorActionPreference = \"Stop\"\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$ProgressPreference = \"SilentlyContinue\"\r\n\r\n# -----------------------------------------------------\r\n# Build ListView Column Lookup Table\r\n#\r\n# Reference Columns by Name Incase Column Order Changes\r\n# -----------------------------------------------------\r\n$Columns = @{}\r\n$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }\r\n\r\n# ------------------------------------------------\r\n# Check if Thread was Already Completed and Exit\r\n# ------------------------------------------------\r\nIf ($ListViewItem.SubItems[$Columns[\"Status Message\"]].Text -eq \"Completed\")\r\n{\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  Exit\r\n}\r\n\r\n# ----------------------------------------------------\r\n# Check if Threads are Paused and Update Thread Status\r\n# ----------------------------------------------------\r\nIf ($SyncedHash.Pause)\r\n{\r\n  # Set Paused Status\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Pause\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  While ($SyncedHash.Pause)\r\n  {\r\n    [System.Threading.Thread]::Sleep(100)\r\n  }\r\n}\r\n\r\n# -----------------------------------------------------\r\n# Check For Termination and Update Thread Status\r\n# -----------------------------------------------------\r\nIf ($SyncedHash.Terminate)\r\n{\r\n  # Set Terminated Status and Exit Thread\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Terminated\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  $ListViewItem.ImageKey = $InfoIcon\r\n  Exit\r\n}\r\n\r\n# Sucess Default Exit Status\r\n$WasSuccess = $True\r\n$ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Processing\"\r\n$ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n$UserName = $ListViewItem.SubItems[$Columns[\"UserName\"]].Text\r\n\r\nTry\r\n{\r\n  # Get Current Domain / Forest\r\n  if ($ADForest -eq \"Domain\")\r\n  {\r\n    if ($ADDomain -eq \"Current\")\r\n    {\r\n      $GetADDomain = Get-MyADDomain -ErrorAction SilentlyContinue\r\n    }\r\n    else\r\n    {\r\n      $GetADDomain = Get-MyADDomain -Name $ADDomain -ErrorAction SilentlyContinue\r\n    }\r\n    if (-not [String]::IsNullOrEmpty($GetADDomain.Name))\r\n    {\r\n      $SearchRoot = \"LDAP://$(\"dc=$(($GetADDomain.Name -split \u0027\\.\u0027) -join \u0027,dc=\u0027)\")\"\r\n      $GetADDomain.Dispose()\r\n    }\r\n  }\r\n  else\r\n  {\r\n    if ($ADDomain -eq \"Current\")\r\n    {\r\n      $GetADForget = Get-MyADForest -ErrorAction SilentlyContinue\r\n    }\r\n    else\r\n    {\r\n      $GetADForget = Get-MyADForest -Name $ADForest -ErrorAction SilentlyContinue\r\n    }\r\n    if (-not [String]::IsNullOrEmpty($GetADForget.Name))\r\n    {\r\n      $SearchRoot = \"GC://$($GetADForget.Name)\"\r\n      $GetADForget.Dispose()\r\n    }\r\n  }\r\n  \r\n  # Check Domain / Forest Found\r\n  if ([String]::IsNullOrEmpty($SearchRoot))\r\n  {\r\n    $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"Unable to Get Current AD Domain / Forest\"\r\n    $WasSuccess = $False\r\n  }\r\n  else\r\n  {\r\n    $LDAPFilter = \"(\u0026(objectClass=user)(objectCategory=person)(sAMAccountType=805306368)(sAMAccountName={0}))\" -f $UserName\r\n    $PropertiesToLoad = @(\"name\", \"canonicalName\", \"userPrincipalName\", \"mail\", \"lastLogonTimestamp\", \"pwdLastSet\", \"userAccountControl\", \"distinguishedName\")\r\n    $ADObject = Get-MyADObject -SearchRoot $SearchRoot -SearchScope Subtree -LDAPFilter $LDAPFilter -PropertiesToLoad $PropertiesToLoad -ErrorAction SilentlyContinue | Select-Object -First 1\r\n    if ([String]::IsNullOrEmpty($ADObject.Path))\r\n    {\r\n      $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"Computer Not Found in AD Forest\"\r\n      $WasSuccess = $False\r\n    }\r\n    else\r\n    {\r\n      # User Info\r\n      $ListViewItem.SubItems[$Columns[\"userPrincipalName\"]].Text = $ADObject.Properties[\"userPrincipalName\"][0]\r\n      $ListViewItem.SubItems[$Columns[\"E-Mail\"]].Text = $ADObject.Properties[\"mail\"][0]\r\n      \r\n      # CanonicalName\r\n      $CanonicalName = $ADObject.Properties[\"canonicalName\"][0]\r\n      $ListViewItem.SubItems[$Columns[\"canonicalName\"]].Text = $CanonicalName\r\n      \r\n      # distinguishedName\r\n      $ListViewItem.SubItems[$Columns[\"distinguishedName\"]].Text = $ADObject.Properties[\"distinguishedName\"][0]\r\n      \r\n      # Domain\r\n      $Domain = $CanonicalName -split \"/\" | Select-Object -First 1\r\n      $ListViewItem.SubItems[$Columns[\"Domain\"]].Text = $Domain\r\n      \r\n      # Zero Hour\r\n      $ZeroHour = [DateTime]::New(1601, 1, 1, 0, 0, 0)\r\n      \r\n      # Last Logon TimeStamp\r\n      $LastLogonTimestamp = $ADObject.Properties[\"lastLogonTimestamp\"][0]\r\n      $LastLogonTimeStampDate = $ZeroHour.AddTicks($LastLogonTimestamp)\r\n      $ListViewItem.SubItems[$Columns[\"Last Logon\"]].Text = $LastLogonTimeStampDate.ToString(\"G\")\r\n      \r\n      # Password Last Set\r\n      $PwdLastSet = $ADObject.Properties[\"pwdLastSet\"][0]\r\n      $PwdLastSetDate = $ZeroHour.AddTicks($PwdLastSet)\r\n      $ListViewItem.SubItems[$Columns[\"PwdLastSet\"]].Text = $PwdLastSetDate.ToString(\"G\")\r\n      \r\n      # User Account Control Flags\r\n      $UserAccountControl = $ADObject.Properties[\"userAccountControl\"][0]\r\n      $ListViewItem.SubItems[$Columns[\"UserAccountControl\"]].Text = \"0$($UserAccountControl)\"\r\n      $ListViewItem.SubItems[$Columns[\"PwdNoChange\"]].Text = (($UserAccountControl -band 64) -ne 0)\r\n      $ListViewItem.SubItems[$Columns[\"PwdNoExpire\"]].Text = (($UserAccountControl -band 65536) -ne 0)\r\n      $ListViewItem.SubItems[$Columns[\"PwdExpired\"]].Text = (($UserAccountControl -band 8388608) -ne 0)\r\n      $ListViewItem.SubItems[$Columns[\"Locked Out\"]].Text = (($UserAccountControl -band 16) -ne 0)\r\n      $ListViewItem.SubItems[$Columns[\"Disabled\"]].Text = (($UserAccountControl -band 2) -ne 0)\r\n    }\r\n  }\r\n}\r\nCatch\r\n{\r\n  # Set Error Message / Thread Failed\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = $PSItem.ToString()\r\n  $WasSuccess = $False\r\n}\r\n\r\n\r\n# Set Final Date / Time and Update Status\r\n$ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\nIf ($WasSuccess)\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Completed\"\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"\"\r\n}\r\nElse\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $BadIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Error\"\r\n}\r\n\r\nExit\r\n\r\n\r\n\r\n"
}
'@
#endregion $GetDomainUser

#region $GraphAPIDevice
$GraphAPIDevice = @'
{
  "ColumnNames": [
    "DisplayName",
    "ID",
    "DeviceID",
    "DeviceOwnership",
    "TrustType",
    "Manufacturer",
    "Model",
    "operatingSystem",
    "OperatingSystemVersion",
    "AccountEnabled",
    "Date/Time",
    "Status Message",
    "Error Message"
  ],
  "Modules": {},
  "Functions": {
    "Get-MyGQuery": {
      "Name": "Get-MyGQuery",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Query Microsoft Graph API with simple paging support.\r\n    .DESCRIPTION\r\n      This function queries the Microsoft Graph API using a provided authentication token and supports basic query options such as API version, resource endpoint, and retrieving all pages of results.\r\n      It is designed for straightforward queries where advanced filtering or selection is not required.\r\n    .PARAMETER AuthToken\r\n      The authentication token (as a hashtable) to use for the request. Typically obtained from an OAuth flow or authentication function.\r\n    .PARAMETER Version\r\n      The Graph API version to use. Accepts \"Beta\" or \"v1.0\". Default is \"Beta\".\r\n    .PARAMETER Resource\r\n      The resource endpoint to query in the Graph API (e.g., \"users\", \"groups\", \"me/messages\").\r\n    .PARAMETER All\r\n      If specified, retrieves all pages of results by following the @odata.nextLink property.\r\n    .PARAMETER Wait\r\n      The number of milliseconds to wait between requests when paging through results. Default is 100.\r\n    .EXAMPLE\r\n      Get-MyGQuery -AuthToken $AuthToken -Resource \"users\"\r\n    .EXAMPLE\r\n      Get-MyGQuery -AuthToken $AuthToken -Resource \"groups\" -Version \"v1.0\" -All\r\n    .EXAMPLE\r\n      Get-MyGQuery -AuthToken $AuthToken -Resource \"me/messages\" -Wait 200\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding()]\r\n  param (\r\n    [parameter(Mandatory = $True)]\r\n    [Hashtable]$AuthToken = $Script:Authtoken,\r\n    [ValidateSet(\"Beta\", \"v1.0\")]\r\n    [String]$Version = \"Beta\",\r\n    [parameter(Mandatory = $True)]\r\n    [String]$Resource,\r\n    [Switch]$All,\r\n    [Int]$Wait = 100\r\n  )\r\n  Write-Verbose -Message \"Enter Function Get-MyGQuery\"\r\n  \r\n  $Uri = \"https://graph.microsoft.com/$($Version)/$($Resource)\"\r\n  do\r\n  {\r\n    Write-Verbose -Message \"Query Graph API\"\r\n    $ReturnData = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Headers $AuthToken -Method Get -ContentType application/json -ErrorAction SilentlyContinue -Verbose:$False\r\n    if ($ReturnData.StatusCode -eq 200)\r\n    {\r\n      $Content = $ReturnData.Content | ConvertFrom-Json\r\n      if (@($Content.PSObject.Properties.match(\"value\")).Count)\r\n      {\r\n        $Content.Value\r\n      }\r\n      else\r\n      {\r\n        $Content\r\n      }\r\n      $Uri = ($Content.\"@odata.nextLink\")\r\n      Start-Sleep -Milliseconds $Wait\r\n    }\r\n    else\r\n    {\r\n      $Uri = $Null\r\n    }\r\n  }\r\n  while ((-not [String]::IsNullOrEmpty($Uri)) -and $All.IsPresent)\r\n  \r\n  Write-Verbose -Message \"Exit Function Get-MyGQuery\"\r\n"
    },
    "Get-MyOAuthApplicationToken": {
      "Name": "Get-MyOAuthApplicationToken",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Get Application OAuth Token\r\n    .DESCRIPTION\r\n      Retrieves an OAuth 2.0 token for an application using client credentials flow.\r\n      This token can be used to authenticate requests to Microsoft Graph or other Azure AD protected resources.\r\n    .PARAMETER TenantID\r\n      The Azure Active Directory tenant ID where the application is registered.\r\n    .PARAMETER ClientID\r\n      The Application (client) ID of the Azure AD app registration.\r\n    .PARAMETER ClientSecret\r\n      The client secret associated with the Azure AD app registration.\r\n    .PARAMETER Scope\r\n      The resource URI or scope for which the token is requested. Defaults to \u0027https://graph.microsoft.com/.default\u0027.\r\n    .EXAMPLE\r\n      Get-MyOAuthApplicationToken -TenantID $TenantID -ClientID $ClientID -ClientSecret $ClientSecret\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding(DefaultParameterSetName = \"New\")]\r\n  param (\r\n    [parameter(Mandatory = $True)]\r\n    [String]$MyTenantID,\r\n    [parameter(Mandatory = $True)]\r\n    [String]$MyClientID,\r\n    [parameter(Mandatory = $True)]\r\n    [String]$MyClientSecret,\r\n    [String]$Scope = \"https://graph.microsoft.com/.default\"\r\n  )\r\n  Write-Verbose -Message \"Enter Function Get-MyOAuthApplicationToken\"\r\n  \r\n  $Body = @{\r\n    \"grant_type\"    = \"client_credentials\"\r\n    \"client_id\"     = $MyClientID\r\n    \"client_secret\" = $MyClientSecret\r\n    \"Scope\"         = $Scope\r\n  }\r\n  \r\n  $Uri = \"https://login.microsoftonline.com/$($MyTenantID)/oauth2/v2.0/token\"\r\n  \r\n  try\r\n  {\r\n    $AuthResult = Invoke-RestMethod -Uri $Uri -Body $Body -Method Post -ContentType \"application/x-www-form-urlencoded\" -ErrorAction SilentlyContinue\r\n  }\r\n  catch\r\n  {\r\n    $AuthResult = $Null\r\n  }\r\n  \r\n  if ([String]::IsNullOrEmpty($AuthResult))\r\n  {\r\n    # Failed to Authenticate\r\n    @{\r\n      \"Expires_In\" = 0\r\n    }\r\n  }\r\n  else\r\n  {\r\n    # Successful Authentication\r\n    @{\r\n      \"Content-Type\"  = \"application/json\"\r\n      \"Authorization\" = \"Bearer \" + $AuthResult.Access_Token\r\n      \"Expires_In\"    = $AuthResult.Expires_In\r\n    }\r\n  }\r\n  \r\n  Write-Verbose -Message \"Exit Function Get-MyOAuthApplicationToken\"\r\n"
    }
  },
  "Variables": {
    "ClientID": { "Name": "ClientID", "Value": "*" },
    "ClientSecret": { "Name": "ClientSecret", "Value": "*" },
    "TenantID": { "Name": "TenantID", "Value": "*" }
  },
  "ThreadCount": 4,
  "ThreadScript": "\u003c#\r\n  .SYNOPSIS\r\n    Sample Runspace Pool Thread Script\r\n  .DESCRIPTION\r\n    Sample Runspace Pool Thread Script\r\n  .PARAMETER ListViewItem\r\n    ListViewItem Passed to the Thread Script\r\n\r\n    This Paramter is Required in your Thread Script\r\n  .EXAMPLE\r\n    Test-Script.ps1 -ListViewItem $ListViewItem\r\n  .NOTES\r\n    Sample Thread Script\r\n\r\n   -------------------------\r\n   ListViewItem Status Icons\r\n   -------------------------\r\n   $GoodIcon = Solid Green Circle\r\n   $BadIcon = Solid Red Circle\r\n   $InfoIcon = Solid Blue Circle\r\n   $CheckIcon = Checkmark\r\n   $ErrorIcon = Red X\r\n   $UpIcon = Green up Arrow \r\n   $DownIcon = Red Down Arrow\r\n\r\n#\u003e\r\n[CmdletBinding()]\r\nParam (\r\n  [parameter(Mandatory = $True)]\r\n  [System.Windows.Forms.ListViewItem]$ListViewItem\r\n)\r\n\r\n# Set Preference Variables\r\n$ErrorActionPreference = \"Stop\"\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$ProgressPreference = \"SilentlyContinue\"\r\n\r\n# -----------------------------------------------------\r\n# Build ListView Column Lookup Table\r\n#\r\n# Reference Columns by Name Incase Column Order Changes\r\n# -----------------------------------------------------\r\n$Columns = @{}\r\n$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }\r\n\r\n# ------------------------------------------------\r\n# Check if Thread was Already Completed and Exit\r\n# ------------------------------------------------\r\nIf ($ListViewItem.SubItems[$Columns[\"Status Message\"]].Text -eq \"Completed\")\r\n{\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  Exit\r\n}\r\n\r\n# ----------------------------------------------------\r\n# Check if Threads are Paused and Update Thread Status\r\n# ----------------------------------------------------\r\nIf ($SyncedHash.Pause)\r\n{\r\n  # Set Paused Status\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Pause\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  While ($SyncedHash.Pause)\r\n  {[System.Threading.Thread]::Sleep(100)\r\n  }\r\n}\r\n\r\n# -----------------------------------------------------\r\n# Check For Termination and Update Thread Status\r\n# -----------------------------------------------------\r\nIf ($SyncedHash.Terminate)\r\n{\r\n  # Set Terminated Status and Exit Thread\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Terminated\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  $ListViewItem.ImageKey = $InfoIcon\r\n  Exit\r\n}\r\n\r\n# Sucess Default Exit Status\r\n$WasSuccess = $True\r\n$ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Processing\"\r\n$DisplayName = $ListViewItem.SubItems[$Columns[\"DisplayName\"]].Text\r\n\r\nTry\r\n{\r\n  $AuthToken = Get-MyOAuthApplicationToken -MyTenantID $TenantID -MyClientID $ClientID -MyClientSecret $ClientSecret\r\n  if ($AuthToken.Expires_In -eq 0)\r\n  {\r\n    $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"Unable to get AuthToken\"\r\n    $WasSuccess = $False\r\n  }\r\n  else\r\n  {\r\n    $Resource = (\"/devices?`$filter=displayName eq \u0027{0}\u0027\u0026`$top=1\u0026`$select=id,displayname,deviceId,deviceOwnership,trustType,manufacturer,model,operatingSystem,operatingSystemVersion,accountEnabled\" -f $DisplayName)\r\n    $Device = Get-MyGQuery -AuthToken $AuthToken -Resource $Resource -ErrorAction SilentlyContinue\r\n    if ([String]::IsNullOrEmpty($Device.ID))\r\n    {\r\n      $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"No Device Found in Azure AD / Entra ID\"\r\n      $WasSuccess = $False\r\n    }\r\n    else\r\n    {\r\n      $ListViewItem.SubItems[$Columns[\"DisplayName\"]].Text = $Device.DisplayName\r\n      $ListViewItem.SubItems[$Columns[\"ID\"]].Text = $Device.ID\r\n      $ListViewItem.SubItems[$Columns[\"DeviceID\"]].Text = $Device.DeviceID\r\n      $ListViewItem.SubItems[$Columns[\"DeviceOwnership\"]].Text = $Device.DeviceOwnership\r\n      $ListViewItem.SubItems[$Columns[\"TrustType\"]].Text = $Device.TrustType\r\n      $ListViewItem.SubItems[$Columns[\"Manufacturer\"]].Text = $Device.Manufacturer\r\n      $ListViewItem.SubItems[$Columns[\"Model\"]].Text = $Device.Model\r\n      $ListViewItem.SubItems[$Columns[\"OperatingSystem\"]].Text = $Device.OperatingSystem\r\n      $ListViewItem.SubItems[$Columns[\"OperatingSystemVersion\"]].Text = $Device.OperatingSystemVersion\r\n      $ListViewItem.SubItems[$Columns[\"AccountEnabled\"]].Text = $Device.AccountEnabled\r\n    }\r\n  }\r\n}\r\nCatch\r\n{\r\n  # Set Error Message / Thread Failed\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = $PSItem.ToString()\r\n  $WasSuccess = $False\r\n}\r\n\r\n# Set Final Date / Time and Update Status\r\n$ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\nIf ($WasSuccess)\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Completed\"\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"\"\r\n}\r\nElse\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $BadIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Error\"\r\n}\r\n\r\nExit\r\n\r\n\r\n"
}
'@
#endregion $GraphAPIDevice

#region $GraphAPIUser
$GraphAPIUser = @'
{
  "ColumnNames": [
    "UserPrincipalName",
    "ID",
    "E-Mail",
    "DisplayName",
    "FirstName",
    "Surname",
    "AccountEnabled",
    "Date/Time",
    "Status Message",
    "Error Message"
  ],
  "Modules": {},
  "Functions": {
    "Get-MyOAuthApplicationToken": {
      "Name": "Get-MyOAuthApplicationToken",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Get Application OAuth Token\r\n    .DESCRIPTION\r\n      Retrieves an OAuth 2.0 token for an application using client credentials flow.\r\n      This token can be used to authenticate requests to Microsoft Graph or other Azure AD protected resources.\r\n    .PARAMETER TenantID\r\n      The Azure Active Directory tenant ID where the application is registered.\r\n    .PARAMETER ClientID\r\n      The Application (client) ID of the Azure AD app registration.\r\n    .PARAMETER ClientSecret\r\n      The client secret associated with the Azure AD app registration.\r\n    .PARAMETER Scope\r\n      The resource URI or scope for which the token is requested. Defaults to \u0027https://graph.microsoft.com/.default\u0027.\r\n    .EXAMPLE\r\n      Get-MyOAuthApplicationToken -TenantID $TenantID -ClientID $ClientID -ClientSecret $ClientSecret\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding(DefaultParameterSetName = \"New\")]\r\n  param (\r\n    [parameter(Mandatory = $True)]\r\n    [String]$MyTenantID,\r\n    [parameter(Mandatory = $True)]\r\n    [String]$MyClientID,\r\n    [parameter(Mandatory = $True)]\r\n    [String]$MyClientSecret,\r\n    [String]$Scope = \"https://graph.microsoft.com/.default\"\r\n  )\r\n  Write-Verbose -Message \"Enter Function Get-MyOAuthApplicationToken\"\r\n  \r\n  $Body = @{\r\n    \"grant_type\"    = \"client_credentials\"\r\n    \"client_id\"     = $MyClientID\r\n    \"client_secret\" = $MyClientSecret\r\n    \"Scope\"         = $Scope\r\n  }\r\n  \r\n  $Uri = \"https://login.microsoftonline.com/$($MyTenantID)/oauth2/v2.0/token\"\r\n  \r\n  try\r\n  {\r\n    $AuthResult = Invoke-RestMethod -Uri $Uri -Body $Body -Method Post -ContentType \"application/x-www-form-urlencoded\" -ErrorAction SilentlyContinue\r\n  }\r\n  catch\r\n  {\r\n    $AuthResult = $Null\r\n  }\r\n  \r\n  if ([String]::IsNullOrEmpty($AuthResult))\r\n  {\r\n    # Failed to Authenticate\r\n    @{\r\n      \"Expires_In\" = 0\r\n    }\r\n  }\r\n  else\r\n  {\r\n    # Successful Authentication\r\n    @{\r\n      \"Content-Type\"  = \"application/json\"\r\n      \"Authorization\" = \"Bearer \" + $AuthResult.Access_Token\r\n      \"Expires_In\"    = $AuthResult.Expires_In\r\n    }\r\n  }\r\n  \r\n  Write-Verbose -Message \"Exit Function Get-MyOAuthApplicationToken\"\r\n"
    },
    "Get-MyGQuery": {
      "Name": "Get-MyGQuery",
      "ScriptBlock": "\r\n  \u003c#\r\n    .SYNOPSIS\r\n      Query Microsoft Graph API with simple paging support.\r\n    .DESCRIPTION\r\n      This function queries the Microsoft Graph API using a provided authentication token and supports basic query options such as API version, resource endpoint, and retrieving all pages of results.\r\n      It is designed for straightforward queries where advanced filtering or selection is not required.\r\n    .PARAMETER AuthToken\r\n      The authentication token (as a hashtable) to use for the request. Typically obtained from an OAuth flow or authentication function.\r\n    .PARAMETER Version\r\n      The Graph API version to use. Accepts \"Beta\" or \"v1.0\". Default is \"Beta\".\r\n    .PARAMETER Resource\r\n      The resource endpoint to query in the Graph API (e.g., \"users\", \"groups\", \"me/messages\").\r\n    .PARAMETER All\r\n      If specified, retrieves all pages of results by following the @odata.nextLink property.\r\n    .PARAMETER Wait\r\n      The number of milliseconds to wait between requests when paging through results. Default is 100.\r\n    .EXAMPLE\r\n      Get-MyGQuery -AuthToken $AuthToken -Resource \"users\"\r\n    .EXAMPLE\r\n      Get-MyGQuery -AuthToken $AuthToken -Resource \"groups\" -Version \"v1.0\" -All\r\n    .EXAMPLE\r\n      Get-MyGQuery -AuthToken $AuthToken -Resource \"me/messages\" -Wait 200\r\n    .NOTES\r\n      Original Function By Ken Sweet\r\n  #\u003e\r\n  [CmdletBinding()]\r\n  param (\r\n    [parameter(Mandatory = $True)]\r\n    [Hashtable]$AuthToken = $Script:Authtoken,\r\n    [ValidateSet(\"Beta\", \"v1.0\")]\r\n    [String]$Version = \"Beta\",\r\n    [parameter(Mandatory = $True)]\r\n    [String]$Resource,\r\n    [Switch]$All,\r\n    [Int]$Wait = 100\r\n  )\r\n  Write-Verbose -Message \"Enter Function Get-MyGQuery\"\r\n  \r\n  $Uri = \"https://graph.microsoft.com/$($Version)/$($Resource)\"\r\n  do\r\n  {\r\n    Write-Verbose -Message \"Query Graph API\"\r\n    $ReturnData = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Headers $AuthToken -Method Get -ContentType application/json -ErrorAction SilentlyContinue -Verbose:$False\r\n    if ($ReturnData.StatusCode -eq 200)\r\n    {\r\n      $Content = $ReturnData.Content | ConvertFrom-Json\r\n      if (@($Content.PSObject.Properties.match(\"value\")).Count)\r\n      {\r\n        $Content.Value\r\n      }\r\n      else\r\n      {\r\n        $Content\r\n      }\r\n      $Uri = ($Content.\"@odata.nextLink\")\r\n      Start-Sleep -Milliseconds $Wait\r\n    }\r\n    else\r\n    {\r\n      $Uri = $Null\r\n    }\r\n  }\r\n  while ((-not [String]::IsNullOrEmpty($Uri)) -and $All.IsPresent)\r\n  \r\n  Write-Verbose -Message \"Exit Function Get-MyGQuery\"\r\n"
    }
  },
  "Variables": {
    "ClientID": {
      "Name": "ClientID",
      "Value": "*"
    },
    "ClientSecret": {
      "Name": "ClientSecret",
      "Value": "*"
    },
    "TenantID": {
      "Name": "TenantID",
      "Value": "*"
    }
  },
  "ThreadCount": 4,
  "ThreadScript": "\u003c#\r\n  .SYNOPSIS\r\n    Sample Runspace Pool Thread Script\r\n  .DESCRIPTION\r\n    Sample Runspace Pool Thread Script\r\n  .PARAMETER ListViewItem\r\n    ListViewItem Passed to the Thread Script\r\n\r\n    This Paramter is Required in your Thread Script\r\n  .EXAMPLE\r\n    Test-Script.ps1 -ListViewItem $ListViewItem\r\n  .NOTES\r\n    Sample Thread Script\r\n\r\n   -------------------------\r\n   ListViewItem Status Icons\r\n   -------------------------\r\n   $GoodIcon = Solid Green Circle\r\n   $BadIcon = Solid Red Circle\r\n   $InfoIcon = Solid Blue Circle\r\n   $CheckIcon = Checkmark\r\n   $ErrorIcon = Red X\r\n   $UpIcon = Green up Arrow \r\n   $DownIcon = Red Down Arrow\r\n\r\n#\u003e\r\n[CmdletBinding()]\r\nParam (\r\n  [parameter(Mandatory = $True)]\r\n  [System.Windows.Forms.ListViewItem]$ListViewItem\r\n)\r\n\r\n# Set Preference Variables\r\n$ErrorActionPreference = \"Stop\"\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$ProgressPreference = \"SilentlyContinue\"\r\n\r\n# -----------------------------------------------------\r\n# Build ListView Column Lookup Table\r\n#\r\n# Reference Columns by Name Incase Column Order Changes\r\n# -----------------------------------------------------\r\n$Columns = @{}\r\n$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }\r\n\r\n# ------------------------------------------------\r\n# Check if Thread was Already Completed and Exit\r\n# ------------------------------------------------\r\nIf ($ListViewItem.SubItems[$Columns[\"Status Message\"]].Text -eq \"Completed\")\r\n{\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  Exit\r\n}\r\n\r\n# ----------------------------------------------------\r\n# Check if Threads are Paused and Update Thread Status\r\n# ----------------------------------------------------\r\nIf ($SyncedHash.Pause)\r\n{\r\n  # Set Paused Status\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Pause\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  While ($SyncedHash.Pause)\r\n  {[System.Threading.Thread]::Sleep(100)\r\n  }\r\n}\r\n\r\n# -----------------------------------------------------\r\n# Check For Termination and Update Thread Status\r\n# -----------------------------------------------------\r\nIf ($SyncedHash.Terminate)\r\n{\r\n  # Set Terminated Status and Exit Thread\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Terminated\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  $ListViewItem.ImageKey = $InfoIcon\r\n  Exit\r\n}\r\n\r\n# Sucess Default Exit Status\r\n$WasSuccess = $True\r\n$ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Processing\"\r\n$UserPrincipalName = $ListViewItem.SubItems[$Columns[\"UserPrincipalName\"]].Text\r\n\r\nTry\r\n{\r\n  $AuthToken = Get-MyOAuthApplicationToken -MyTenantID $TenantID -MyClientID $ClientID -MyClientSecret $ClientSecret\r\n  if ($AuthToken.Expires_In -eq 0)\r\n  {\r\n    $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"Unable to get AuthToken\"\r\n    $WasSuccess = $False\r\n  }\r\n  else\r\n  {\r\n    $Resource = (\"/users/$($UserPrincipalName)?`$top=1\u0026`$select=id,displayname,mail,givenName,surname,accountEnabled\" -f $DisplayName)\r\n    $Device = Get-MyGQuery -AuthToken $AuthToken -Resource $Resource -ErrorAction SilentlyContinue\r\n    if ([String]::IsNullOrEmpty($Device.ID))\r\n    {\r\n      $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"No Device Found in Azure AD / Entra ID\"\r\n      $WasSuccess = $False\r\n    }\r\n    else\r\n    {\r\n      $ListViewItem.SubItems[$Columns[\"ID\"]].Text = $Device.ID\r\n      $ListViewItem.SubItems[$Columns[\"E-Mail\"]].Text = $Device.Mail\r\n      $ListViewItem.SubItems[$Columns[\"DisplayName\"]].Text = $Device.DisplayName\r\n      $ListViewItem.SubItems[$Columns[\"FirstName\"]].Text = $Device.GivenName\r\n      $ListViewItem.SubItems[$Columns[\"Surname\"]].Text = $Device.Surname\r\n      $ListViewItem.SubItems[$Columns[\"AccountEnabled\"]].Text = $Device.AccountEnabled\r\n    }\r\n  }\r\n}\r\nCatch\r\n{\r\n  # Set Error Message / Thread Failed\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = $PSItem.ToString()\r\n  $WasSuccess = $False\r\n  \r\n  Write-Host -Object ($($Error[0].Exception.Message))\r\n  Write-Host -Object (($Error[0].InvocationInfo.Line).Trim())\r\n  Write-Host -Object ($Error[0].InvocationInfo.ScriptLineNumber)\r\n}\r\n\r\n# Set Final Date / Time and Update Status\r\n$ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\nIf ($WasSuccess)\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Completed\"\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"\"\r\n}\r\nElse\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $BadIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Error\"\r\n}\r\n\r\nExit\r\n\r\n\r\n"
}
'@
#endregion $GraphAPIUser

#region Basic Starter Script
$StarterConfig = @'
{
  "ColumnNames": [
    "List Item",
    "Data Column",
    "Column Name 02",
    "Column Name 03",
    "Column Name 04",
    "Column Name 05",
    "Column Name 06",
    "Column Name 07",
    "Column Name 08",
    "Date/Time",
    "Status Message",
    "Error Message"
  ],
  "Modules": {},
  "Functions": {},
  "Variables": {},
  "ThreadCount": 4,
  "ThreadScript": "\u003c#\r\n  .SYNOPSIS\r\n    Sample Runspace Pool Thread Script\r\n  .DESCRIPTION\r\n    Sample Runspace Pool Thread Script\r\n  .PARAMETER ListViewItem\r\n    ListViewItem Passed to the Thread Script\r\n\r\n    This Paramter is Required in your Thread Script\r\n  .EXAMPLE\r\n    Test-Script.ps1 -ListViewItem $ListViewItem\r\n  .NOTES\r\n    Sample Thread Script\r\n\r\n   -------------------------\r\n   ListViewItem Status Icons\r\n   -------------------------\r\n   $GoodIcon = Solid Green Circle\r\n   $BadIcon = Solid Red Circle\r\n   $InfoIcon = Solid Blue Circle\r\n   $CheckIcon = Checkmark\r\n   $ErrorIcon = Red X\r\n   $UpIcon = Green up Arrow \r\n   $DownIcon = Red Down Arrow\r\n\r\n#\u003e\r\n[CmdletBinding()]\r\nParam (\r\n  [parameter(Mandatory = $True)]\r\n  [System.Windows.Forms.ListViewItem]$ListViewItem\r\n)\r\n\r\n# Set Preference Variables\r\n$ErrorActionPreference = \"Stop\"\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$ProgressPreference = \"SilentlyContinue\"\r\n\r\n# -----------------------------------------------------\r\n# Build ListView Column Lookup Table\r\n#\r\n# Reference Columns by Name Incase Column Order Changes\r\n# -----------------------------------------------------\r\n$Columns = @{ }\r\n$ListViewItem.ListView.Columns | ForEach-Object -Process { $Columns.Add($PSItem.Text, $PSItem.Index) }\r\n\r\n# ------------------------------------------------\r\n# Check if Thread was Already Completed and Exit\r\n# ------------------------------------------------\r\nIf ($ListViewItem.SubItems[$Columns[\"Status Message\"]].Text -eq \"Completed\")\r\n{\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  Exit\r\n}\r\n\r\n# ----------------------------------------------------\r\n# Check if Threads are Paused and Update Thread Status\r\n# ----------------------------------------------------\r\nIf ($SyncedHash.Pause)\r\n{\r\n  # Set Paused Status\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Pause\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  While ($SyncedHash.Pause)\r\n  {\r\n    [System.Threading.Thread]::Sleep(100)\r\n  }\r\n}\r\n\r\n# -----------------------------------------------------\r\n# Check For Termination and Update Thread Status\r\n# -----------------------------------------------------\r\nIf ($SyncedHash.Terminate)\r\n{\r\n  # Set Terminated Status and Exit Thread\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Terminated\"\r\n  $ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\n  $ListViewItem.ImageKey = $InfoIcon\r\n  Exit\r\n}\r\n\r\n# Sucess Default Exit Status\r\n$WasSuccess = $True\r\n$ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Processing\"\r\n$CurrentItem = $ListViewItem.SubItems[$Columns[\"List Item\"]].Text\r\n\r\nTry\r\n{\r\n  # Get / Update Shared Object / Value\r\n  If ([System.String]::IsNullOrEmpty($SyncedHash.Object))\r\n  {\r\n    $SyncedHash.Object = \"First Item\"\r\n  }\r\n  $ListViewItem.SubItems[$Columns[\"Data Column\"]].Text = $SyncedHash.Object\r\n  $SyncedHash.Object = $CurrentItem\r\n  \r\n  # ---------------------------------------------------------\r\n  # Open and wait for Mutex - Limit Access to Shared Resource\r\n  # ---------------------------------------------------------\r\n  $MyMutex = [System.Threading.Mutex]::OpenExisting($Mutex)\r\n  [Void]($MyMutex.WaitOne())\r\n  \r\n  # Access / Update Shared Resources\r\n  # $CurrentItem | Out-File -Encoding ascii -FilePath \"C:\\SharedFile.txt\"\r\n  \r\n  # Release Mutex\r\n  $MyMutex.ReleaseMutex()\r\n}\r\nCatch\r\n{\r\n  # Set Error Message / Thread Failed\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = $PSItem.ToString()\r\n  $WasSuccess = $False\r\n}\r\n\r\n# File Remaining Columns\r\nFor ($I = 4; $I -lt 11; $I++)\r\n{\r\n  $ListViewItem.SubItems[$I].Text = [DateTime]::Now.ToString(\"G\")\r\n  [System.Threading.Thread]::Sleep(100)\r\n}\r\n\r\n# Set Final Date / Time and Update Status\r\n$ListViewItem.SubItems[$Columns[\"Date/Time\"]].Text = [DateTime]::Now.ToString(\"G\")\r\nIf ($WasSuccess)\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $GoodIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Completed\"\r\n  $ListViewItem.SubItems[$Columns[\"Error Message\"]].Text = \"\"\r\n}\r\nElse\r\n{\r\n  # Return Success\r\n  $ListViewItem.ImageKey = $BadIcon\r\n  $ListViewItem.SubItems[$Columns[\"Status Message\"]].Text = \"Error\"\r\n}\r\n\r\nExit\r\n\r\n\r\n\r\n\r\n\r\n\r\n\r\n"
}
'@
#endregion Basic Starter Script

#region $UnknownConfig
$UnknownConfig = @'
{
  "ColumnNames": [
    "Column Name 00",
    "Column Name 01",
    "Column Name 02",
    "Column Name 03",
    "Column Name 04",
    "Column Name 05",
    "Column Name 06",
    "Column Name 07",
    "Column Name 08",
    "Column Name 09",
    "Column Name 10",
    "Column Name 11"
  ],
  "Modules": {},
  "Functions": {},
  "Variables": {},
  "ThreadCount": 8,
  "ThreadScript": ""
}
'@
#endregion $UnknownConfig

#endregion ******** PIL Demos ********

#region ******** My Default Enumerations ********

#region ******** enum MyAnswer ********
[Flags()]
Enum MyAnswer
{
  Unknown = 0
  No = 1
  Yes = 2
  Maybe = 3
}
#endregion ******** enum MyAnswer ********

#region ******** enum MyDigit ********
Enum MyDigit
{
  Zero
  One
  Two
  Three
  Four
  Five
  Six
  Seven
  Eight
  Nine
}
#endregion ******** enum MyDigit ********

#region ******** enum MyBits ********
[Flags()]
Enum MyBits
{
  Bit01 = 0x00000001
  Bit02 = 0x00000002
  Bit03 = 0x00000004
  Bit04 = 0x00000008
  Bit05 = 0x00000010
  Bit06 = 0x00000020
  Bit07 = 0x00000040
  Bit08 = 0x00000080
  Bit09 = 0x00000100
  Bit10 = 0x00000200
  Bit11 = 0x00000400
  Bit12 = 0x00000800
  Bit13 = 0x00001000
  Bit14 = 0x00002000
  Bit15 = 0x00004000
  Bit16 = 0x00008000
}
#endregion ******** enum MyBits ********

#endregion ******** My Default Enumerations ********

#region ******** My Custom Class ********

#region ******** MyListItem Class ********
Class MyListItem
{
  [String]$Text
  [Object]$Value
  [Object]$Tag
  [MyBits]$Flags
  
  MyListItem ([String]$Text, [Object]$Value)
  {
    $This.Text = $Text
    $This.Value = $Value
  }
  
  MyListItem ([String]$Text, [Object]$Value, [MyBits]$Flags)
  {
    $This.Text = $Text
    $This.Value = $Value
    $This.Flags = $Flags
  }
  
  MyListItem ([String]$Text, [Object]$Value, [Object]$Tag)
  {
    $This.Text = $Text
    $This.Value = $Value
    $This.Tag = $Tag
  }
  
  MyListItem ([String]$Text, [Object]$Value, [Object]$Tag, [MyBits]$Flags)
  {
    $This.Text = $Text
    $This.Value = $Value
    $This.Tag = $Tag
    $This.Flags = $Flags
  }
}
#endregion ******** MyListItem Class ********

#endregion ******** My Custom Class ********

#region ******** Windows APIs ********

#region ******** [Console.Window] ********

#[Void][Console.Window]::Hide()
#[Void][Console.Window]::Show()

$MyCode = @"
using System;
using System.Runtime.InteropServices;

namespace Console
{
  public class Window
  {
    [DllImport("Kernel32.dll")]
    private static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

    public static bool Hide()
    {
      return ShowWindowAsync(GetConsoleWindow(), 0);
    }

    public static bool Show()
    {
      return ShowWindowAsync(GetConsoleWindow(), 5);
    }
  }
}
"@
Add-Type -TypeDefinition $MyCode -Debug:$False
#endregion ******** [Console.Window] ********

[System.Console]::Title = "RUNNING: $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
if ([MyConfig]::Production)
{
  [Void][Console.Window]::Hide()
}

#endregion ******** Windows APIs ********

#region ******** Functions Library ********

#region Function Prompt
Function Prompt
{
  [Console]::Title = $PWD
  "PS$($PSVersionTable.PSVersion.Major)$(">" * ($NestedPromptLevel + 1)) "
}
#endregion Function Prompt

#region function New-MenuItem
function New-MenuItem()
{
  <#
    .SYNOPSIS
      Creates and adds a new MenuItem to a Menu or ToolStrip control.
    .DESCRIPTION
      This function creates a new System.Windows.Forms.ToolStripMenuItem with the specified properties and adds it to the provided Menu or ToolStrip control.
      It supports customization of text, name, tooltip, icon, image index/key, text-image relation, display style, alignment, tag, enabled/disabled state, checked state, shortcut keys, font, and colors.
      The new MenuItem can optionally be returned via the PassThru switch.
    .PARAMETER Menu
      The Menu or ToolStrip control to which the new MenuItem will be added. This parameter is mandatory.
    .PARAMETER Text
      The display text for the new MenuItem. This parameter is mandatory.
    .PARAMETER Name
      The name of the new MenuItem. If not specified, the Text value is used.
    .PARAMETER ToolTip
      The tooltip text to display when the mouse hovers over the MenuItem.
    .PARAMETER Icon
      The icon to display for the MenuItem. Used when specifying images by icon. Mandatory for the 'Icon' parameter set.
    .PARAMETER ImageIndex
      The index of the image to display for the MenuItem. Used when specifying images by index. Mandatory for the 'ImageIndex' parameter set.
    .PARAMETER ImageKey
      The key of the image to display for the MenuItem. Used when specifying images by key. Mandatory for the 'ImageKey' parameter set.
    .PARAMETER TextImageRelation
      Specifies the position of the text and image relative to each other. Defaults to 'ImageBeforeText'.
    .PARAMETER DisplayStyle
      Specifies how the MenuItem displays its image and text. Defaults to 'Text'.
    .PARAMETER Alignment
      Specifies the alignment of the MenuItem's text and image. Defaults to 'MiddleCenter'.
    .PARAMETER Tag
      An object to associate additional data with the new MenuItem.
    .PARAMETER Disable
      If specified, disables the MenuItem (sets Enabled to $false).
    .PARAMETER Check
      If specified, sets the MenuItem's Checked property to $true.
    .PARAMETER ClickOnCheck
      If specified, enables the CheckOnClick property for the MenuItem.
    .PARAMETER ShortcutKeys
      Specifies the shortcut keys for the MenuItem. Defaults to 'None'.
    .PARAMETER Font
      The font to use for the MenuItem text. Defaults to [MyConfig]::Font.Regular.
    .PARAMETER BackColor
      The background color of the MenuItem. Defaults to [MyConfig]::Colors.Back.
    .PARAMETER ForeColor
      The foreground (text) color of the MenuItem. Defaults to [MyConfig]::Colors.Fore.
    .PARAMETER PassThru
      If specified, returns the newly created MenuItem object.
    .EXAMPLE
      $NewItem = New-MenuItem -Menu $menuStrip -Text "Open" -Tag "OpenFile"
      Adds a new MenuItem with text "Open" and tag "OpenFile" to $menuStrip.
    .EXAMPLE
      $item = New-MenuItem -Menu $contextMenu -Text "Save" -ImageIndex 2 -PassThru
      Adds a new MenuItem with an image at index 2 and returns the created MenuItem object.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [parameter(Mandatory = $True)]
    [Object]$Menu,
    [parameter(Mandatory = $True)]
    [String]$Text,
    [String]$Name,
    [String]$ToolTip,
    [parameter(Mandatory = $True, ParameterSetName = "Icon")]
    [System.Drawing.Icon]$Icon,
    [parameter(Mandatory = $True, ParameterSetName = "ImageIndex")]
    [Int]$ImageIndex,
    [parameter(Mandatory = $True, ParameterSetName = "ImageKey")]
    [String]$ImageKey,
    [System.Windows.Forms.TextImageRelation]$TextImageRelation = "ImageBeforeText",
    [System.Windows.Forms.ToolStripItemDisplayStyle]$DisplayStyle = "Text",
    [System.Drawing.ContentAlignment]$Alignment = "MiddleCenter",
    [Object]$Tag,
    [Switch]$Disable,
    [Switch]$Check,
    [Switch]$ClickOnCheck,
    [System.Windows.Forms.Keys]$ShortcutKeys = "None",
    [System.Drawing.Font]$Font = [MyConfig]::Font.Regular,
    [System.Drawing.Color]$BackColor = [MyConfig]::Colors.Back,
    [System.Drawing.Color]$ForeColor = [MyConfig]::Colors.Fore,
    [switch]$PassThru
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  #region $TempMenuItem = [System.Windows.Forms.ToolStripMenuItem]
  $TempMenuItem = [System.Windows.Forms.ToolStripMenuItem]::New($Text)

  if ($Menu.GetType().Name -eq "ToolStripMenuItem")
  {
    [Void]$Menu.DropDownItems.Add($TempMenuItem)
    if ($Menu.DropDown.Items.Count -eq 1)
    {
      $Menu.DropDown.BackColor = $Menu.BackColor
      $Menu.DropDown.ForeColor = $Menu.ForeColor
      $Menu.DropDown.ImageList = $Menu.Owner.ImageList
    }
  }
  else
  {
    [Void]$Menu.Items.Add($TempMenuItem)
  }

  if ($PSBoundParameters.ContainsKey("Name"))
  {
    $TempMenuItem.Name = $Name
  }
  else
  {
    $TempMenuItem.Name = $Text
  }

  $TempMenuItem.ShortcutKeys = $ShortcutKeys
  $TempMenuItem.Tag = $Tag
  $TempMenuItem.ToolTipText = $ToolTip
  $TempMenuItem.TextAlign = $Alignment
  $TempMenuItem.Checked = $Check.IsPresent
  $TempMenuItem.CheckOnClick = $ClickOnCheck.IsPresent
  $TempMenuItem.DisplayStyle = $DisplayStyle
  $TempMenuItem.Enabled = (-not $Disable.IsPresent)

  $TempMenuItem.BackColor = $BackColor
  $TempMenuItem.ForeColor = $ForeColor
  $TempMenuItem.Font = $Font

  if ($PSCmdlet.ParameterSetName -eq "Default")
  {
    $TempMenuItem.TextImageRelation = [System.Windows.Forms.TextImageRelation]::TextBeforeImage
  }
  else
  {
    switch ($PSCmdlet.ParameterSetName)
    {
      "Icon"
      {
        $TempMenuItem.Image = $Icon
        break
      }
      "ImageIndex"
      {
        $TempMenuItem.ImageIndex = $ImageIndex
        break
      }
      "ImageKey"
      {
        $TempMenuItem.ImageKey = $ImageKey
        break
      }
    }
    $TempMenuItem.ImageAlign = $Alignment
    $TempMenuItem.TextImageRelation = $TextImageRelation
  }
  #endregion $TempMenuItem = [System.Windows.Forms.ToolStripMenuItem]

  if ($PassThru.IsPresent)
  {
    $TempMenuItem
  }

  $TempMenuItem = $Null

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function New-MenuItem

#region function New-MenuLabel
function New-MenuLabel()
{
  <#
    .SYNOPSIS
      Creates and adds a new MenuLabel (ToolStripLabel) to a Menu or ToolStrip control.
    .DESCRIPTION
      This function creates a new System.Windows.Forms.ToolStripLabel with the specified properties and adds it to the provided Menu or ToolStrip control.
      It supports customization of text, name, tooltip, icon, display style, alignment, tag, enabled/disabled state, font, and colors.
      The new MenuLabel can optionally be returned via the PassThru switch.
    .PARAMETER Menu
      The Menu or ToolStrip control to which the new MenuLabel will be added. This parameter is mandatory.
    .PARAMETER Text
      The display text for the new MenuLabel. This parameter is mandatory.
    .PARAMETER Name
      The name of the new MenuLabel. If not specified, the Text value is used.  
    .PARAMETER ToolTip
      The tooltip text to display when the mouse hovers over the MenuLabel.
    .PARAMETER Icon
      The icon to display for the MenuLabel. If specified, the label will show the icon.
    .PARAMETER DisplayStyle
      Specifies how the MenuLabel displays its image and text. Defaults to 'Text'.
    .PARAMETER Alignment
      Specifies the alignment of the MenuLabel's text and image. Defaults to 'MiddleLeft'.
    .PARAMETER Tag
      An object to associate additional data with the new MenuLabel.
    .PARAMETER Disable
      If specified, disables the MenuLabel (sets Enabled to $false).
    .PARAMETER Font
      The font to use for the MenuLabel text. Defaults to [MyConfig]::Font.Regular.
    .PARAMETER BackColor
      The background color of the MenuLabel. Defaults to [MyConfig]::Colors.Back.
    .PARAMETER ForeColor
      The foreground (text) color of the MenuLabel. Defaults to [MyConfig]::Colors.Fore.
    .PARAMETER PassThru
      If specified, returns the newly created MenuLabel object.
    .EXAMPLE
      $NewItem = New-MenuLabel -Menu $menuStrip -Text "Info" -Tag "Information"
      Adds a new MenuLabel with text "Info" and tag "Information" to $menuStrip.
    .EXAMPLE
      $label = New-MenuLabel -Menu $contextMenu -Text "Help" -Icon $icon -PassThru
      Adds a new MenuLabel with an icon and returns the created MenuLabel object.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [Object]$Menu,
    [parameter(Mandatory = $True)]
    [String]$Text,
    [String]$Name,
    [String]$ToolTip,
    [System.Drawing.Icon]$Icon,
    [System.Windows.Forms.ToolStripItemDisplayStyle]$DisplayStyle = "Text",
    [System.Drawing.ContentAlignment]$Alignment = "MiddleLeft",
    [Object]$Tag,
    [Switch]$Disable,
    [System.Drawing.Font]$Font = [MyConfig]::Font.Regular,
    [System.Drawing.Color]$BackColor = [MyConfig]::Colors.Back,
    [System.Drawing.Color]$ForeColor = [MyConfig]::Colors.Fore,
    [switch]$PassThru
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  #region $TempMenuLabel = [System.Windows.Forms.ToolStripLabel]
  $TempMenuLabel = [System.Windows.Forms.ToolStripLabel]::New($Text)

  if ($Menu.GetType().Name -eq "ToolStripMenuItem")
  {
    [Void]$Menu.DropDownItems.Add($TempMenuLabel)
  }
  else
  {
    [Void]$Menu.Items.Add($TempMenuLabel)
  }

  if ($PSBoundParameters.ContainsKey("Name"))
  {
    $TempMenuLabel.Name = $Name
  }
  else
  {
    $TempMenuLabel.Name = $Text
  }

  $TempMenuLabel.TextAlign = $Alignment
  $TempMenuLabel.Tag = $Tag
  $TempMenuLabel.ToolTipText = $ToolTip
  $TempMenuLabel.DisplayStyle = $DisplayStyle
  $TempMenuLabel.Enabled = (-not $Disable.IsPresent)

  $TempMenuLabel.BackColor = $BackColor
  $TempMenuLabel.ForeColor = $ForeColor
  $TempMenuLabel.Font = $Font

  if ($PSBoundParameters.ContainsKey("Icon"))
  {
    $TempMenuLabel.Image = $Icon
    $TempMenuLabel.ImageAlign = $Alignment
    $TempMenuLabel.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
  }
  else
  {
    $TempMenuLabel.TextImageRelation = [System.Windows.Forms.TextImageRelation]::TextBeforeImage
  }
  #endregion $TempMenuLabel = [System.Windows.Forms.ToolStripLabel]

  if ($PassThru)
  {
    $TempMenuLabel
  }

  $TempMenuLabel = $Null

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function New-MenuLabel

#region function New-MenuSeparator
function New-MenuSeparator()
{
  <#
    .SYNOPSIS
      Creates and adds a new MenuSeparator (ToolStripSeparator) to a Menu or ToolStrip control.
    .DESCRIPTION
      This function creates a new System.Windows.Forms.ToolStripSeparator and adds it to the provided Menu or ToolStrip control.
      It supports customization of background and foreground colors.
      The separator is useful for visually grouping related menu items.
    .PARAMETER Menu
      The Menu or ToolStrip control to which the new MenuSeparator will be added. This parameter is mandatory.
    .PARAMETER BackColor
      The background color of the MenuSeparator. Defaults to [MyConfig]::Colors.Back.
    .PARAMETER ForeColor
      The foreground (line) color of the MenuSeparator. Defaults to [MyConfig]::Colors.Fore.
    .EXAMPLE
      New-MenuSeparator -Menu $Menu
      Adds a new separator to the specified $Menu control.
    .EXAMPLE
      New-MenuSeparator -Menu $contextMenu -BackColor ([System.Drawing.Color]::LightGray) -ForeColor ([System.Drawing.Color]::DarkGray)
      Adds a new separator to $contextMenu with custom colors.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param(
    [parameter(Mandatory = $True)]
    [Object]$Menu,
    [System.Drawing.Color]$BackColor = [MyConfig]::Colors.Back,
    [System.Drawing.Color]$ForeColor = [MyConfig]::Colors.Fore
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  #region $TempSeparator = [System.Windows.Forms.ToolStripSeparator]
  $TempSeparator = [System.Windows.Forms.ToolStripSeparator]::New()

  if ($Menu.GetType().Name -eq "ToolStripMenuItem")
  {
    [Void]$Menu.DropDownItems.Add($TempSeparator)
  }
  else
  {
    [Void]$Menu.Items.Add($TempSeparator)
  }

  $TempSeparator.Name = "TempSeparator"

  $TempSeparator.BackColor = $BackColor
  $TempSeparator.ForeColor = $ForeColor
  #endregion $TempSeparator = [System.Windows.Forms.ToolStripSeparator]

  $TempSeparator = $Null

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function New-MenuSeparator

#region function New-ColumnHeader
function New-ColumnHeader()
{
  <#
    .SYNOPSIS
      Creates and adds a new ColumnHeader to a ListView control.
    .DESCRIPTION
      This function creates a new System.Windows.Forms.ColumnHeader with the specified properties and adds it to the provided ListView control.
      It supports customization of the header text, name, tag, and width. The new ColumnHeader can optionally be returned via the PassThru switch.
    .PARAMETER ListView
      The ListView control to which the new ColumnHeader will be added. This parameter is mandatory.
    .PARAMETER Text
      The display text for the new ColumnHeader. This parameter is mandatory.
    .PARAMETER Name
      The name of the new ColumnHeader. If not specified, the Text value is used.
    .PARAMETER Tag
      An object to associate additional data with the new ColumnHeader.
    .PARAMETER Width
      The width of the new ColumnHeader in pixels. Defaults to -2 (auto size).
    .PARAMETER PassThru
      If specified, returns the newly created ColumnHeader object.
    .EXAMPLE
      $NewItem = New-ColumnHeader -ListView $listView -Text "Name" -Tag "UserName"
      Adds a new ColumnHeader with text "Name" and tag "UserName" to $listView.
    .EXAMPLE
      $col = New-ColumnHeader -ListView $listView -Text "Date" -Width 120 -PassThru
      Adds a new ColumnHeader with text "Date" and width 120 pixels, and returns the created ColumnHeader object.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param(
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ListView]$ListView,
    [parameter(Mandatory = $True)]
    [String]$Text,
    [String]$Name,
    [Object]$Tag,
    [Int]$Width = -2,
    [switch]$PassThru
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  #region $TempColumnHeader = [System.Windows.Forms.ColumnHeader]
  $TempColumnHeader = [System.Windows.Forms.ColumnHeader]::New()
  [Void]$ListView.Columns.Add($TempColumnHeader)
  $TempColumnHeader.Tag = $Tag
  $TempColumnHeader.Text = $Text
  if ($PSBoundParameters.ContainsKey("Name"))
  {
    $TempColumnHeader.Name = $Name
  }
  else
  {
    $TempColumnHeader.Name = $Text
  }
  $TempColumnHeader.Width = $Width
  #endregion $TempColumnHeader = [System.Windows.Forms.ColumnHeader]

  if ($PassThru.IsPresent)
  {
    $TempColumnHeader
  }

  $TempColumnHeader = $Null

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function New-ColumnHeader

#region Custom ListView Sort

$MyCode = @"
using System;
using System.Windows.Forms;
using System.Collections;

namespace MyCustom
{
  public class ListViewSort : IComparer
  {
    private int _Column = 0;
    private bool _Ascending = true;
    private bool _Enable = true;

    public ListViewSort()
    {
      _Column = 0;
      _Ascending = true;
    }

    public ListViewSort(int Column)
    {
      _Column = Column;
      _Ascending = true;
    }

    public ListViewSort(int Column, bool Order)
    {
      _Column = Column;
      _Ascending = Order;
    }

    public int Column
    {
      get { return _Column; }
      set { _Column = value; }
    }

    public bool Ascending
    {
      get { return _Ascending; }
      set { _Ascending = value; }
    }

    public bool Enable
    {
      get { return _Enable; }
      set { _Enable = value; }
    }

    public int Compare(object RowX, object RowY)
    {
      if (_Enable)
      {
        if (_Ascending)
        {
          return String.Compare(((System.Windows.Forms.ListViewItem)RowX).SubItems[_Column].Text, ((System.Windows.Forms.ListViewItem)RowY).SubItems[_Column].Text);
        }
        else
        {
          return String.Compare(((System.Windows.Forms.ListViewItem)RowY).SubItems[_Column].Text, ((System.Windows.Forms.ListViewItem)RowX).SubItems[_Column].Text);
        }
      }
      else
      {
        return 0;
      }
    }
  }
}
"@
Add-Type -TypeDefinition $MyCode -ReferencedAssemblies "System.Windows.Forms" -Debug:$False

#endregion My Custom ListView Sort

#region function Install-MyModule
Function Install-MyModule ()
{
  <#
    .SYNOPSIS
      Checks for, installs if required, and imports the specified PowerShell module.
    .DESCRIPTION
      This function checks if the specified module is imported or installed. If not, it installs the module from the given repository and imports it. Supports custom repositories and installation scopes.
    .PARAMETER Name
      The name of the module to check, install, and import.
    .PARAMETER Version
      The minimum required version of the module. Defaults to "0.0.0.0" (any version).
    .PARAMETER Scope
      Specifies whether to install/import the module for AllUsers or CurrentUser. Defaults to "AllUsers".
    .PARAMETER Repository
      The name of the repository to use for installation. Defaults to "sie-powershell".
    .PARAMETER Install
      If specified, performs the installation of the module if not present.
    .PARAMETER NoImport
      If specified, Do not Import the specified module
    .PARAMETER SourceLocation
      The URL of the repository source location. Used when registering a custom repository.
    .PARAMETER PublishLocation
      The URL of the repository publish location. Used when registering a custom repository.
    .EXAMPLE
      Install-MyModule -Name "MSAL.PS" -Version "2.0.0.0" -Scope "AllUsers" -Install
      Checks for MSAL.PS module, installs version 2.0.0.0 or higher for all users if required, and imports it.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  Param (
    [parameter(Mandatory = $True)]
    [String]$Name,
    [Version]$Version = "0.0.0.0",
    [ValidateSet("AllUsers", "CurrentUser")]
    [String]$Scope = "AllUsers",
    [parameter(Mandatory = $False, ParameterSetName = "Default")]
    [parameter(Mandatory = $True, ParameterSetName = "Custom")]
    [String]$Repository = "PSGallery",
    [Switch]$Install,
    [Switch]$NoImport,
    [parameter(Mandatory = $True, ParameterSetName = "Custom")]
    [String]$SourceLocation,
    [parameter(Mandatory = $True, ParameterSetName = "Custom")]
    [String]$PublishLocation
  )
  Write-Verbose -Message "Enter Function Install-MyModule"

  # Zero Verion for Checks
  $ZeroVersion = [Version]::new(0, 0, 0, 0)

  # Get Module Common Parameters
  $GMParams = @{
    "Name"          = $Name
    "WarningAction" = "SilentlyContinue"
    "ErrorAction"   = "SilentlyContinue"
    "Verbose"       = $False
  }

  # Install Module Parameters
  $IMParams = @{
    "Name"          = $Name
    "WarningAction" = "SilentlyContinue"
    "ErrorAction"   = "SilentlyContinue"
    "Verbose"       = $False
  }
  If ($PSBoundParameters.ContainsKey("Version"))
  {
    $IMParams.Add("RequiredVersion", $Version)
  }

  # Check if Module is Already Imported
  $ChkInstalled = Get-Module @GMParams | Sort-Object -Property Version -Descending | Select-Object -Property Version -First 1
  If ([String]::IsNullOrEmpty($ChkInstalled.Version))
  {
    # Get Installed Module Versions
    $ChkInstalled = Get-InstalledModule @GMParams -AllVersions | Where-Object -FilterScript { ($PSItem.Version -eq $Version) -or ($Version -eq $ZeroVersion) } | Sort-Object -Property Version -Descending | Select-Object -Property Version -First 1
    If ([String]::IsNullOrEmpty($ChkInstalled.Version))
    {
      If (((([Security.Principal.WindowsPrincipal]::New([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -and ($Scope -eq "AllUsers")) -or ($Scope -eq "CurrentUser")) -and $Install.IsPresent)
      {
        # Check if Repo Exists
        $ChkRepo = Get-PSRepository -Name $Repository -ErrorAction SilentlyContinue
        If ([String]::IsNullOrEmpty($ChkRepo.Name))
        {
          # Add Custom Repo
          Register-PSRepository -Name $Repository -SourceLocation $SourceLocation -PublishLocation $PublishLocation -InstallationPolicy "Trusted"
        }
        # Install / Update Module
        Install-Module @IMParams -Repository $Repository -Scope $Scope -Force -AllowClobber | Out-Null
        If ($Repository -ne "PSGallery")
        {
          # Remove Custom Repo
          Unregister-PSRepository -Name $Repository
        }
        $ChkInstalled = Get-InstalledModule @GMParams -AllVersions | Where-Object -FilterScript { ($PSItem.Version -eq $Version) -or ($Version -eq $ZeroVersion) } | Sort-Object -Property Version -Descending | Select-Object -Property Version -First 1
        If ([String]::IsNullOrEmpty($ChkInstalled.Version))
        {
          # Module Installed Failed
          [PSCustomObject]@{ "Success" = $False; "Message" = "Module Install Failed" }
        }
        Else
        {
          If ($NoImport.IsPresent)
          {
            [PSCustomObject]@{ "Success" = $True; "Message" = "Module Install Succeeded" }
          }
          Else
          {
            # Import Module
            Import-Module @IMParams
            # Verify Imported Module
            $ChkImported = Get-Module @GMParams | Sort-Object -Property Version -Descending | Select-Object -Property Version -First 1
            If ($ChkImported.Version -eq $ChkInstalled.Version)
            {
              # Module Install / Import Succeeded
              [PSCustomObject]@{ "Success" = $True; "Message" = "Module Install / Import Succeeded" }
            }
            Else
            {
              # Module Install / Import Failed
              [PSCustomObject]@{ "Success" = $False; "Message" = "Module Install / Import Failed" }
            }
          }
        }
      }
      Else
      {
        # Module Install / Import Failed
        [PSCustomObject]@{ "Success" = $False; "Message" = "Module Install / Import Not Installed" }
      }
    }
    Else
    {
      If ($NoImport.IsPresent)
      {
        [PSCustomObject]@{ "Success" = $True; "Message" = "Module Install Succeeded" }
      }
      Else
      {
        # Import Module
        Import-Module @IMParams
        # Verify Imported Module
        $ChkImported = Get-Module @GMParams | Sort-Object -Property Version -Descending | Select-Object -Property Version -First 1
        If ($ChkImported.Version -eq $ChkInstalled.Version)
        {
          # Module Import Succeeded
          [PSCustomObject]@{ "Success" = $True; "Message" = "Module Import Succeeded" }
        }
        Else
        {
          # Module Import Failed
          [PSCustomObject]@{ "Success" = $False; "Message" = "Module Import Failed" }
        }
      }
    }
  }
  Else
  {
    # Module Previously Imported
    If (($ChkInstalled.Version -eq $Version) -or ($Version -eq $ZeroVersion))
    {
      # Correct Module Version Imported
      [PSCustomObject]@{ "Success" = $True; "Message" = "Correct Version Previously Loaded" }
    }
    Else
    {
      # Wrong Module Version Imported
      [PSCustomObject]@{ "Success" = $False; "Message" = "Wrong Version Previously Loaded" }
    }
  }

  Write-Verbose -Message "Exit Function Install-MyModule"
}
#endregion function Install-MyModule

#region function Show-MyWebReport
function Show-MyWebReport
{
  <#
    .SYNOPSIS
      Opens a web report in the default browser (Edge or Chrome) as an app window.
    .DESCRIPTION
      This function launches the specified report URL in Microsoft Edge or Google Chrome as an app window. 
      It checks for the configured browser path in [MyConfig] and uses Edge if available, otherwise Chrome. 
      If neither is configured, the function does nothing.
    .PARAMETER ReportURL
      The URL of the web report to open. This parameter is mandatory.
    .EXAMPLE
      Show-MyWebReport -ReportURL "https://myreportserver/report1"
      Opens the specified report in the configured browser as an app window.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [String]$ReportURL
  )
  Write-Verbose -Message "Enter Function Show-MyWebReport"

  if ([String]::IsNullOrEmpty(([MyConfig]::EdgePath)))
  {
    if (-not [String]::IsNullOrEmpty(([MyConfig]::ChromePath)))
    {
      Start-Process -FilePath ([MyConfig]::ChromePath) -ArgumentList "--app=`"$($ReportURL)`""
    }
  }
  else
  {
    Start-Process -FilePath ([MyConfig]::EdgePath) -ArgumentList "--app=`"$($ReportURL)`""
  }

  Write-Verbose -Message "Exit Function Show-MyWebReport"
}
#endregion function Show-MyWebReport

#endregion ******** Functions Library ********

#region ******** Multiple Thread Functions ********

#region ******** Custom Objects MyRSPool / MyRSJob ********

$MyCode = @"
using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading;

public class MyRSJob
{
  private System.String _Name;
  private System.String _PoolName;
  private System.Guid _PoolID;
  private System.Management.Automation.PowerShell _PowerShell;
  private System.IAsyncResult _PowerShellAsyncResult;
  private System.Object _InputObject = null;

  public MyRSJob(System.String Name, System.Management.Automation.PowerShell PowerShell, System.IAsyncResult PowerShellAsyncResult, System.Object InputObject, System.String PoolName, System.Guid PoolID)
  {
    _Name = Name;
    _PoolName = PoolName;
    _PoolID = PoolID;
    _PowerShell = PowerShell;
    _PowerShellAsyncResult = PowerShellAsyncResult;
    _InputObject = InputObject;
  }

  public System.String Name
  {
    get
    {
      return _Name;
    }
  }

  public System.Guid InstanceID
  {
    get
    {
      return _PowerShell.InstanceId;
    }
  }

  public System.String PoolName
  {
    get
    {
      return _PoolName;
    }
  }

  public System.Guid PoolID
  {
    get
    {
      return _PoolID;
    }
  }

  public System.Management.Automation.PowerShell PowerShell
  {
    get
    {
      return _PowerShell;
    }
  }

  public System.Management.Automation.PSInvocationState State
  {
    get
    {
      return _PowerShell.InvocationStateInfo.State;
    }
  }

  public System.Exception Reason
  {
    get
    {
      return _PowerShell.InvocationStateInfo.Reason;
    }
  }

  public bool HadErrors
  {
    get
    {
      return _PowerShell.HadErrors;
    }
  }

  public System.String Command
  {
    get
    {
      return _PowerShell.Commands.Commands[0].ToString();
    }
  }

  public System.Management.Automation.Runspaces.RunspacePool RunspacePool
  {
    get
    {
      return _PowerShell.RunspacePool;
    }
  }

  public System.IAsyncResult PowerShellAsyncResult
  {
    get
    {
      return _PowerShellAsyncResult;
    }
  }

  public bool IsCompleted
  {
    get
    {
      return _PowerShellAsyncResult.IsCompleted;
    }
  }

  public System.Object InputObject
  {
    get
    {
      return _InputObject;
    }
  }

  public System.Management.Automation.PSDataCollection<System.Management.Automation.DebugRecord> Debug
  {
    get
    {
      return _PowerShell.Streams.Debug;
    }
  }

  public System.Management.Automation.PSDataCollection<System.Management.Automation.ErrorRecord> Error
  {
    get
    {
      return _PowerShell.Streams.Error;
    }
  }

  public System.Management.Automation.PSDataCollection<System.Management.Automation.ProgressRecord> Progress
  {
    get
    {
      return _PowerShell.Streams.Progress;
    }
  }

  public System.Management.Automation.PSDataCollection<System.Management.Automation.VerboseRecord> Verbose
  {
    get
    {
      return _PowerShell.Streams.Verbose;
    }
  }

  public System.Management.Automation.PSDataCollection<System.Management.Automation.WarningRecord> Warning
  {
    get
    {
      return _PowerShell.Streams.Warning;
    }
  }
}

public class MyRSPool
{
  private System.String _Name;  
  private System.Management.Automation.Runspaces.RunspacePool _RunspacePool;
  public System.Collections.Generic.List<MyRSJob> Jobs = new System.Collections.Generic.List<MyRSJob>();
  private System.Collections.Hashtable _SyncedHash;
  private System.Threading.Mutex _Mutex;  

  public MyRSPool(System.String Name, System.Management.Automation.Runspaces.RunspacePool RunspacePool, System.Collections.Hashtable SyncedHash) 
  {
    _Name = Name;
    _RunspacePool = RunspacePool;
    _SyncedHash = SyncedHash;
  }

  public MyRSPool(System.String Name, System.Management.Automation.Runspaces.RunspacePool RunspacePool, System.Collections.Hashtable SyncedHash, System.String Mutex) 
  {
    _Name = Name;
    _RunspacePool = RunspacePool;
    _SyncedHash = SyncedHash;
    _Mutex = new System.Threading.Mutex(false, Mutex);
  }

  public System.Collections.Hashtable SyncedHash
  {
    get
    {
      return _SyncedHash;
    }
  }

  public System.Threading.Mutex Mutex
  {
    get
    {
      return _Mutex;
    }
  }

  public System.String Name
  {
    get
    {
      return _Name;
    }
  }

  public System.Guid InstanceID
  {
    get
    {
      return _RunspacePool.InstanceId;
    }
  }

  public System.Management.Automation.Runspaces.RunspacePool RunspacePool
  {
    get
    {
      return _RunspacePool;
    }
  }

  public System.Management.Automation.Runspaces.RunspacePoolState State
  {
    get
    {
      return _RunspacePool.RunspacePoolStateInfo.State;
    }
  }
}
"@
Add-Type -TypeDefinition $MyCode -Debug:$False

$Script:MyHiddenRSPool = [System.Collections.Generic.Dictionary[[String], [MyRSPool]]]::New()

#endregion ******** Custom Objects MyRSPool / MyRSJob ********

#region function Start-MyRSPool
function Start-MyRSPool()
{
  <#
    .SYNOPSIS
      Creates or Updates a RunspacePool
    .DESCRIPTION
      Function to do something specific
    .PARAMETER PoolName
      Name of RunspacePool
    .PARAMETER Functions
      Functions to include in the initial Session State
    .PARAMETER Variables
      Variables to include in the initial Session State
    .PARAMETER Modules
      Modules to load in the initial Session State
    .PARAMETER PSSnapins
      PSSnapins to load in the initial Session State
    .PARAMETER Hashtable
      Synced Hasttable to pass values between threads
    .PARAMETER Mutex
      Protects access to a shared resource
    .PARAMETER MaxJobs
      Maximum Number of Jobs
    .PARAMETER PassThru
      Return the New RSPool to the Pipeline
    .EXAMPLE
      Start-MyRSPool

      Create the Default RunspacePool
    .EXAMPLE
      $MyRSPool = Start-MyRSPool -PoolName $PoolName -MaxJobs $MaxJobs -PassThru

      Create a New RunspacePool and Return the RSPool to the Pipeline
    .NOTES
      Original Script By Ken Sweet on 10/15/2017
      Updated Script By Ken Sweet on 02/04/2019
    .LINK
  #>
  [CmdletBinding()]
  param (
    [String]$PoolName = "MyDefaultRSPool",
    [Hashtable]$Functions,
    [Hashtable]$Variables,
    [String[]]$Modules,
    [String[]]$PSSnapins,
    [Hashtable]$Hashtable = @{ "Enabled" = $True },
    [String]$Mutex,
    [ValidateRange(1, 64)]
    [Int]$MaxJobs = 8,
    [Switch]$PassThru
  )
  Write-Verbose -Message "Enter Function Start-MyRSPool"

  # check if Runspace Pool already exists
  if ($Script:MyHiddenRSPool.ContainsKey($PoolName))
  {
    # Return Existing Runspace Pool
    [MyRSPool]($Script:MyHiddenRSPool[$PoolName])
  }
  else
  {
    # Create Default Session State
    $InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $InitialSessionState.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::RemoteSigned
    
    # Import Modules
    if ($PSBoundParameters.ContainsKey("Modules"))
    {
      [Void]$InitialSessionState.ImportPSModule($Modules)
    }

    # Import PSSnapins
    if ($PSBoundParameters.ContainsKey("PSSnapins"))
    {
      [Void]$InitialSessionState.ImportPSSnapIn($PSSnapins, [Ref]$Null)
    }

    # Add Common Functions
    if ($PSBoundParameters.ContainsKey("Functions"))
    {
      ForEach ($Key in $Functions.Keys)
      {
        $InitialSessionState.Commands.Add(([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::New($Key, $Functions[$Key])))
      }
    }

    # Add Default Variables
    if ($PSBoundParameters.ContainsKey("Variables"))
    {
      ForEach ($Key in $Variables.Keys)
      {
        $InitialSessionState.Variables.Add(([System.Management.Automation.Runspaces.SessionStateVariableEntry]::New($Key, $Variables[$Key], "$Key = $($Variables[$Key])", ([System.Management.Automation.ScopedItemOptions]::AllScope))))
      }
    }

    # Create and Open RunSpacePool
    $SyncedHash = [Hashtable]::Synchronized($Hashtable)
    $InitialSessionState.Variables.Add(([System.Management.Automation.Runspaces.SessionStateVariableEntry]::New("SyncedHash", $SyncedHash, "SyncedHash = Synced Hashtable", ([System.Management.Automation.ScopedItemOptions]::AllScope))))
    if ($PSBoundParameters.ContainsKey("Mutex"))
    {
      $InitialSessionState.Variables.Add(([System.Management.Automation.Runspaces.SessionStateVariableEntry]::New("Mutex", $Mutex, "Mutex = $Mutex", ([System.Management.Automation.ScopedItemOptions]::AllScope))))
      $CreateRunspacePool = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxJobs, $InitialSessionState, $Host)
      $RSPool = [MyRSPool]::New($PoolName, $CreateRunspacePool, $SyncedHash, $Mutex)
    }
    else
    {
      $CreateRunspacePool = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxJobs, $InitialSessionState, $Host)
      $RSPool = [MyRSPool]::New($PoolName, $CreateRunspacePool, $SyncedHash)
    }

    $RSPool.RunspacePool.ApartmentState = "STA"
    #$RSPool.RunspacePool.ApartmentState = "MTA"
    $RSPool.RunspacePool.CleanupInterval = [TimeSpan]::FromMinutes(2)
    $RSPool.RunspacePool.Open()

    $Script:MyHiddenRSPool.Add($PoolName, $RSPool)

    if ($PassThru.IsPresent)
    {
      $RSPool
    }
  }

  Write-Verbose -Message "Exit Function Start-MyRSPool"
}
#endregion function Start-MyRSPool

#region function Get-MyRSPool
function Get-MyRSPool()
{
  <#
    .SYNOPSIS
      Get RunspacePools that match specified criteria
    .DESCRIPTION
      Get RunspacePools that match specified criteria
    .PARAMETER PoolName
      Name of RSPool to search for
    .PARAMETER PoolID
      PoolID of Job to search for
    .PARAMETER State
      State of Jobs to search for
    .EXAMPLE
      $MyRSPools = Get-MyRSPool

      Get all RSPools
    .EXAMPLE
      $MyRSPools = Get-MyRSPool -PoolName $PoolName

      $MyRSPools = Get-MyRSPool -PoolID $PoolID

      Get Specified RSPools
    .NOTES
      Original Script By Ken Sweet on 10/15/2017
      Updated Script By Ken Sweet on 02/04/2019
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "All")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "PoolName")]
    [String[]]$PoolName = "MyDefaultRSPool",
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "PoolID")]
    [Guid[]]$PoolID,
    [parameter(Mandatory = $False, ParameterSetName = "All")]
    [parameter(Mandatory = $False, ParameterSetName = "PoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "PoolID")]
    [ValidateSet("BeforeOpen", "Opening", "Opened", "Closed", "Closing", "Broken", "Disconnecting", "Disconnected", "Connecting")]
    [String[]]$State
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Get-MyRSPool Begin Block"

    # Set Job State RegEx Pattern
    if ($PSBoundParameters.ContainsKey("State"))
    {
      $StatePattern = $State -join "|"
    }
    else
    {
      $StatePattern = ".*"
    }

    Write-Verbose -Message "Exit Function Get-MyRSPool Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function Get-MyRSPool Process Block"

    switch ($PSCmdlet.ParameterSetName)
    {
      "All" {
        # Return Matching Pools
        [MyRSPool[]]($Script:MyHiddenRSPool.Values | Where-Object -FilterScript { $PSItem.State -match $StatePattern })
        Break;
      }
      "PoolName" {
        # Set Pool Name and Return Matching Pools
        $NamePattern = $PoolName -join "|"
        [MyRSPool[]]($Script:MyHiddenRSPool.Values | Where-Object -FilterScript { $PSItem.State -match $StatePattern -and $PSItem.Name -match $NamePattern})
        Break;
      }
      "PoolID" {
        # Set PoolID and Return Matching Pools
        $IDPattern = $PoolID -join "|"
        [MyRSPool[]]($Script:MyHiddenRSPool.Values | Where-Object -FilterScript { $PSItem.State -match $StatePattern -and $PSItem.InstanceId -match $IDPattern })
        Break;
      }
    }

    Write-Verbose -Message "Exit Function Get-MyRSPool Process Block"
  }
}
#endregion function Get-MyRSPool

#region function Close-MyRSPool
function Close-MyRSPool()
{
  <#
    .SYNOPSIS
      Close RunspacePool and Stop all Running Jobs
    .DESCRIPTION
      Close RunspacePool and Stop all Running Jobs
    .PARAMETER RSPool
      RunspacePool to clsoe
    .PARAMETER PoolName
      Name of RSPool to close
    .PARAMETER PoolID
      PoolID of Job to close
    .PARAMETER State
      State of Jobs to close
    .EXAMPLE
      Close-MyRSPool

      Close the Default RSPool
    .EXAMPLE
      Close-MyRSPool -PoolName $PoolName

      Close-MyRSPool -PoolID $PoolID

      Close Specified RSPools
    .NOTES
      Original Script By Ken Sweet on 10/15/2017
      Updated Script By Ken Sweet on 02/04/2019
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "All")]
  param (
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "RSPool")]
    [MyRSPool[]]$RSPool,
    [parameter(Mandatory = $True, ParameterSetName = "PoolName")]
    [String[]]$PoolName = "MyDefaultRSPool",
    [parameter(Mandatory = $True, ParameterSetName = "PoolID")]
    [Guid[]]$PoolID,
    [parameter(Mandatory = $False, ParameterSetName = "All")]
    [parameter(Mandatory = $False, ParameterSetName = "PoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "PoolID")]
    [ValidateSet("BeforeOpen", "Opening", "Opened", "Closed", "Closing", "Broken", "Disconnecting", "Disconnected", "Connecting")]
    [String[]]$State
  )
  Process
  {
    Write-Verbose -Message "Enter Function Close-MyRSPool Process Block"

    If ($PSCmdlet.ParameterSetName -eq "RSPool")
    {
      $TempPools = $RSPool
    }
    else
    {
      $TempPools = [MyRSPool[]](Get-MyRSPool @PSBoundParameters)
    }

    # Close RunspacePools, This will Stop all Running Jobs
    ForEach ($TempPool in $TempPools)
    {
      if (-not [String]::IsNullOrEmpty($TempPool.Mutex))
      {
        $TempPool.Mutex.Close()
        $TempPool.Mutex.Dispose()
      }
      $TempPool.RunspacePool.Close()
      $TempPool.RunspacePool.Dispose()
      [Void]$Script:MyHiddenRSPool.Remove($TempPool.Name)
    }

    Write-Verbose -Message "Exit Function Close-MyRSPool Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function Close-MyRSPool End Block"

    # Garbage Collect, Recover Resources
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Function Close-MyRSPool End Block"
  }
}
#endregion function Close-MyRSPool

#region function Start-MyRSJob
Function Start-MyRSJob()
{
  <#
    .SYNOPSIS
      Creates or Updates a RunspacePool
    .DESCRIPTION
      Function to do something specific
    .PARAMETER RSPool
      RunspacePool to add new RunspacePool Jobs to
    .PARAMETER PoolName
      Name of RunspacePool
    .PARAMETER PoolID
      ID of RunspacePool
    .PARAMETER InputObject
      Object / Value to pass to the RunspacePool Job ScriptBlock
    .PARAMETER InputParam
      Paramter to pass the Object / Value as
    .PARAMETER JobName
      Name of RunspacePool Jobs
    .PARAMETER ScriptBlock
      RunspacePool Job ScriptBock to Execute
    .PARAMETER Parameters
      Common Paramaters to pass to the RunspacePool Job ScriptBlock
    .PARAMETER PassThru
      Return the New Jobs to the Pipeline
    .EXAMPLE
      Start-MyRSJob -ScriptBlock $ScriptBlock -JobName $JobName -InputObject $InputObject

      Add new RSJobs to the Default RSPool
    .EXAMPLE
      $InputObject | Start-MyRSJob -ScriptBlock $ScriptBlock -RSPool $RSPool -JobName $JobName
  
      $InputObject | Start-MyRSJob -ScriptBlock $ScriptBlock -PoolName $PoolName -JobName $JobName
  
      $InputObject | Start-MyRSJob -ScriptBlock $ScriptBlock -PoolID $PoolID -JobName $JobName

      Add new RSJobs to the Specified RSPool
    .NOTES
      Original Script By Ken Sweet on 10/15/2017
      Updated Script By Ken Sweet on 02/04/2019
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "PoolName")]
  Param (
    [parameter(Mandatory = $True, ParameterSetName = "RSPool")]
    [MyRSPool]$RSPool,
    [parameter(Mandatory = $False, ParameterSetName = "PoolName")]
    [String]$PoolName = "MyDefaultRSPool",
    [parameter(Mandatory = $True, ParameterSetName = "PoolID")]
    [Guid]$PoolID,
    [parameter(Mandatory = $False, ValueFromPipeline = $True)]
    [Object]$InputObject,
    [String]$InputParam = "InputObject",
    [String]$JobName = "Job Name",
    [parameter(Mandatory = $True)]
    [ScriptBlock]$ScriptBlock,
    [Hashtable]$Parameters,
    [Switch]$PassThru
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Start-MyRSJob Begin Block"
    
    Switch ($PSCmdlet.ParameterSetName)
    {
      "RSPool" {
        # Set Pool
        $TempPool = $RSPool
        Break;
      }
      "PoolName" {
        # Set Pool Name and Return Matching Pools
        $TempPool = [MyRSPool](Start-MyRSPool -PoolName $PoolName -PassThru)
        Break;
      }
      "PoolID" {
        # Set PoolID Return Matching Pools
        $TempPool = [MyRSPool](Get-MyRSPool -PoolID $PoolID)
        Break;
      }
    }
    
    # List for New Jobs
    $NewJobs = [System.Collections.Generic.List[MyRSJob]]::New()
    
    Write-Verbose -Message "Exit Function Start-MyRSJob Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function Start-MyRSJob Process Block"
    
    If ($PSBoundParameters.ContainsKey("InputObject"))
    {
      # Create New PowerShell Instance with ScriptBlock
      $PowerShell = ([Management.Automation.PowerShell]::Create()).AddScript($ScriptBlock)
      # Set RunspacePool
      $PowerShell.RunspacePool = $TempPool.RunspacePool
      # Add Parameters
      [Void]$PowerShell.AddParameter($InputParam, $InputObject)
      If ($PSBoundParameters.ContainsKey("Parameters"))
      {
        [Void]$PowerShell.AddParameters($Parameters)
      }
      # set Job Name
      If (($Object -is [String]) -or ($Object -is [ValueType]))
      {
        $TempJobName = "$JobName - $($Object)"
      }
      Else
      {
        $TempJobName = $($Object.$JobName)
      }
      [Void]$NewJobs.Add(([MyRSjob]::New($TempJobName, $PowerShell, $PowerShell.BeginInvoke(), $Object, $TempPool.Name, $TempPool.InstanceID)))
    }
    Else
    {
      # Create New PowerShell Instance with ScriptBlock
      $PowerShell = ([Management.Automation.PowerShell]::Create()).AddScript($ScriptBlock)
      # Set RunspacePool
      $PowerShell.RunspacePool = $TempPool.RunspacePool
      # Add Parameters
      If ($PSBoundParameters.ContainsKey("Parameters"))
      {
        [Void]$PowerShell.AddParameters($Parameters)
      }
      [Void]$NewJobs.Add(([MyRSjob]::New($JobName, $PowerShell, $PowerShell.BeginInvoke(), $Null, $TempPool.Name, $TempPool.InstanceID)))
    }
    
    Write-Verbose -Message "Exit Function Start-MyRSJob Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function Start-MyRSJob End Block"
    
    If ($NewJobs.Count)
    {
      $TempPool.Jobs.AddRange($NewJobs)
      # Return Jobs only if New RunspacePool
      If ($PassThru.IsPresent)
      {
        $NewJobs
      }
      $NewJobs.Clear()
    }
    
    Write-Verbose -Message "Exit Function Start-MyRSJob End Block"
  }
}
#endregion

#region function Get-MyRSJob
function Get-MyRSJob()
{
  <#
    .SYNOPSIS
      Get Jobs from RunspacePool that match specified criteria
    .DESCRIPTION
      Get Jobs from RunspacePool that match specified criteria
    .PARAMETER RSPool
      RunspacePool to search
    .PARAMETER PoolName
      Name of Pool to Get Jobs From
    .PARAMETER PoolID
      ID of Pool to Get Jobs From
    .PARAMETER JobName
      Name of Jobs to Get
    .PARAMETER JobID
      ID of Jobs to Get
    .PARAMETER State
      State of Jobs to search for
    .EXAMPLE
      $MyRSJobs = Get-MyRSJob

      Get RSJobs from the Default RSPool
    .EXAMPLE
      $MyRSJobs = Get-MyRSJob -RSPool $RSPool

      $MyRSJobs = Get-MyRSJob -PoolName $PoolName

      $MyRSJobs = Get-MyRSJob -PoolID $PoolID

      Get RSJobs from the Specified RSPool
    .NOTES
      Original Script By Ken Sweet on 10/15/2017
      Updated Script By Ken Sweet on 02/04/2019
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "JobNamePoolName")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePool")]
    [MyRSPool[]]$RSPool,
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [String]$PoolName = "MyDefaultRSPool",
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolID")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePoolID")]
    [Guid]$PoolID,
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolID")]
    [String[]]$JobName = ".*",
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "JobIDPoolID")]
    [Guid[]]$JobID,
    [ValidateSet("NotStarted", "Running", "Stopping", "Stopped", "Completed", "Failed", "Disconnected")]
    [String[]]$State
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Get-MyRSJob Begin Block"

    # Set Job State RegEx Pattern
    if ($PSBoundParameters.ContainsKey("State"))
    {
      $StatePattern = $State -join "|"
    }
    else
    {
      $StatePattern = ".*"
    }

    Switch -regex ($PSCmdlet.ParameterSetName)
    {
      "Pool$" {
        # Set Pool
        $TempPools = $RSPool
        Break;
      }
      "PoolName$" {
        # Set Pool Name and Return Matching Pools
        $TempPools = [MyRSPool[]](Get-MyRSPool -PoolName $PoolName)
        Break;
      }
      "PoolID$" {
        # Set PoolID Return Matching Pools
        $TempPools = [MyRSPool[]](Get-MyRSPool -PoolID $PoolID)
        Break;
      }
    }

    Write-Verbose -Message "Exit Function Get-MyRSJob Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function Get-MyRSJob Process Block"

    Switch -regex ($PSCmdlet.ParameterSetName)
    {
      "^JobName" {
        # Set Job Name RegEx Pattern and Return Matching Jobs
        $NamePattern = $JobName -join "|"
        [MyRSJob[]]($TempPools | ForEach-Object -Process { $PSItem.Jobs | Where-Object -FilterScript { $PSItem.State -match $StatePattern -and $PSItem.Name -match $NamePattern } })
        Break;
      }
      "^JobID" {
        # Set Job ID RegEx Pattern and Return Matching Jobs
        $IDPattern = $JobID -join "|"
        [MyRSJob[]]($TempPools | ForEach-Object -Process { $PSItem.Jobs | Where-Object -FilterScript { $PSItem.State -match $StatePattern -and $PSItem.InstanceId -match $IDPattern } })
        Break;
      }
    }

    Write-Verbose -Message "Exit Function Get-MyRSJob Process Block"
  }
}
#endregion function Get-MyRSJob

#region function Wait-MyRSJob
function Wait-MyRSJob()
{
  <#
    .SYNOPSIS
      Wait for RSJob to Finish
    .DESCRIPTION
      Wait for RSJob to Finish
    .PARAMETER RSPool
      RunspacePool to search
    .PARAMETER PoolName
      Name of Pool to Get Jobs From
    .PARAMETER PoolID
      ID of Pool to Get Jobs From
    .PARAMETER JobName
      Name of Jobs to Get
    .PARAMETER JobID
      ID of Jobs to Get
    .PARAMETER State
      State of Jobs to search for
    .PARAMETER ScriptBlock
      ScriptBlock to invoke while waiting

      For windows Forms scripts add the DoEvents method in to the Wait ScritpBlock

      [System.Windows.Forms.Application]::DoEvents()
      [System.Threading.Thread]::Sleep(250)
    .PARAMETER Wait
      TimeSpace to wait
    .PARAMETER NoWait
      No Wait, Return when any Job states changes to Stopped, Completed, or Failed
    .PARAMETER PassThru
      Return the New Jobs to the Pipeline
    .EXAMPLE
      $MyRSJobs = Wait-MyRSJob -PassThru

      Wait for and Get RSJobs from the Default RSPool
    .EXAMPLE
      $MyRSJobs = Wait-MyRSJob -RSPool $RSPool -PassThru

      $MyRSJobs = Wait-MyRSJob -PoolName $PoolName -PassThru

      $MyRSJobs = Wait-MyRSJob -PoolID $PoolID -PassThru

      Wait for and Get RSJobs from the Specified RSPool
    .NOTES
      Original Script By Ken Sweet on 10/15/2017
      Updated Script By Ken Sweet on 02/04/2019
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "JobNamePoolName")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePool")]
    [MyRSPool[]]$RSPool,
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [String]$PoolName = "MyDefaultRSPool",
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolID")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePoolID")]
    [Guid]$PoolID,
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolID")]
    [String[]]$JobName = ".*",
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolID")]
    [Guid[]]$JobID,
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "RSJob")]
    [MyRSJob[]]$RSJob,
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolID")]
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolID")]
    [ValidateSet("NotStarted", "Running", "Stopping", "Stopped", "Completed", "Failed", "Disconnected")]
    [String[]]$State,
    [ScriptBlock]$SciptBlock = { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 200 },
    [UInt16]$Wait = 60,
    [Switch]$NoWait,
    [Switch]$PassThru
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Wait-MyRSJob Begin Block"

    # Remove Invalid Get-MyRSJob Parameters
    if ($PSCmdlet.ParameterSetName -ne "RSJob")
    {
      if ($PSBoundParameters.ContainsKey("PassThru"))
      {
        [Void]$PSBoundParameters.Remove("PassThru")
      }
      if ($PSBoundParameters.ContainsKey("Wait"))
      {
        [Void]$PSBoundParameters.Remove("Wait")
      }
      if ($PSBoundParameters.ContainsKey("NoWait"))
      {
        [Void]$PSBoundParameters.Remove("NoWait")
      }
      if ($PSBoundParameters.ContainsKey("ScriptBlock"))
      {
        [Void]$PSBoundParameters.Remove("ScriptBlock")
      }
    }

    # List for Wait Jobs
    $WaitJobs = [System.Collections.Generic.List[MyRSJob]]::New()

    Write-Verbose -Message "Exit Function Wait-MyRSJob Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function Wait-MyRSJob Process Block"

    # Add Passed RSJobs to $Jobs
    if ($PSCmdlet.ParameterSetName -eq "RSJob")
    {
      $WaitJobs.AddRange([MyRSJob[]]($RSJob))
    }
    else
    {
      $WaitJobs.AddRange([MyRSJob[]](Get-MyRSJob @PSBoundParameters))
    }

    Write-Verbose -Message "Exit Function Wait-MyRSJob Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function Wait-MyRSJob End Block"

    # Wait for Jobs to be Finshed
    if ($NoWait.IsPresent)
    {
      While (@(($WaitJobs | Where-Object -FilterScript { $PSItem.State -notmatch "Stopped|Completed|Failed" })).Count -eq $WaitJobs.Count)
      {
        $SciptBlock.Invoke()
      }
    }
    else
    {
      [Object[]]$CheckJobs = $WaitJobs.ToArray()
      $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
      While (@(($CheckJobs = $CheckJobs | Where-Object -FilterScript { $PSItem.State -notmatch "Stopped|Completed|Failed" })).Count -and (($StopWatch.TotalSeconds -le $Wait) -or ($Wait -eq 0)))
      {
        $SciptBlock.Invoke()
      }
      $StopWatch.Stop()
    }
    
    if ($PassThru.IsPresent)
    {
      # Return Completed Jobs
      [MyRSJob[]]($WaitJobs | Where-Object -FilterScript { $PSItem.State -match "Stopped|Completed|Failed" })
    }
    $WaitJobs.Clear()

    Write-Verbose -Message "Exit Function Wait-MyRSJob End Block"
  }
}
#endregion function Wait-MyRSJob

#region function Stop-MyRSJob
function Stop-MyRSJob()
{
  <#
    .SYNOPSIS
      Function to do something specific
    .DESCRIPTION
      Function to do something specific
    .PARAMETER RSPool
      RunspacePool to search
    .PARAMETER Name
      Name of Job to search for
    .PARAMETER InstanceId
      InstanceId of Job to search for
    .PARAMETER RSJob
      RunspacePool Jobs to Process
    .PARAMETER State
      State of Jobs to search for
    .EXAMPLE
      Stop-MyRSJob

      Stop all RSJobs in the Default RSPool
    .EXAMPLE
      Stop-MyRSJob -RSPool $RSPool

      Stop-MyRSJob -PoolName $PoolName

      Stop-MyRSJob -PoolID $PoolID

      Stop all RSJobs in the Specified RSPool
    .NOTES
      Original Script By Ken Sweet on 10/15/2017
      Updated Script By Ken Sweet on 02/04/2019
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "JobNamePoolName")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePool")]
    [MyRSPool[]]$RSPool,
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [String]$PoolName = "MyDefaultRSPool",
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolID")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePoolID")]
    [Guid]$PoolID,
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolID")]
    [String[]]$JobName = ".*",
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolID")]
    [Guid[]]$JobID,
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "RSJob")]
    [MyRSJob[]]$RSJob,
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolID")]
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolID")]
    [ValidateSet("NotStarted", "Running", "Stopping", "Stopped", "Completed", "Failed", "Disconnected")]
    [String[]]$State
  )
  Process
  {
    Write-Verbose -Message "Enter Function Stop-MyRSJob Process Block"

    # Add Passed RSJobs to $Jobs
    if ($PSCmdlet.ParameterSetName -eq "RSJob")
    {
      $TempJobs = $RSJob
    }
    else
    {
      $TempJobs = [MyRSJob[]](Get-MyRSJob @PSBoundParameters)
    }

    # Stop all Jobs that have not Finished
    ForEach ($TempJob in $TempJobs)
    {
      if ($TempJob.State -notmatch "Stopped|Completed|Failed")
      {
        $TempJob.PowerShell.Stop()
      }
    }

    Write-Verbose -Message "Exit Function Stop-MyRSJob Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function Stop-MyRSJob End Block"

    # Garbage Collect, Recover Resources
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Function Stop-MyRSJob End Block"
  }
}
#endregion function Stop-MyRSJob

#region function Receive-MyRSJob
function Receive-MyRSJob()
{
  <#
    .SYNOPSIS
      Receive Output from Completed Jobs
    .DESCRIPTION
      Receive Output from Completed Jobs
    .PARAMETER RSPool
      RunspacePool to search
    .PARAMETER PoolName
      Name of Pool to Get Jobs From
    .PARAMETER PoolID
      ID of Pool to Get Jobs From
    .PARAMETER JobName
      Name of Jobs to Get
    .PARAMETER JobID
      ID of Jobs to Get
    .PARAMETER RSJob
      Jobs to Process
    .PARAMETER AutoRemove
      Remove Jobs after Receiving Output
    .EXAMPLE
      $MyResults = Receive-MyRSJob -AutoRemove

      Receive Results from RSJobs in the Default RSPool
    .EXAMPLE
      $MyResults = Receive-MyRSJob -RSPool $RSPool -AutoRemove

      $MyResults = Receive-MyRSJob -PoolName $PoolName -AutoRemove

      $MyResults = Receive-MyRSJob -PoolID $PoolID -AutoRemove

      Receive Results from RSJobs in the Specified RSPool
    .NOTES
      Original Script By Ken Sweet on 10/15/2017
      Updated Script By Ken Sweet on 02/04/2019
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "JobNamePoolName")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePool")]
    [MyRSPool[]]$RSPool,
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [String]$PoolName = "MyDefaultRSPool",
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolID")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePoolID")]
    [Guid]$PoolID,
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolID")]
    [String[]]$JobName = ".*",
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "JobIDPoolID")]
    [Guid[]]$JobID,
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "RSJob")]
    [MyRSJob[]]$RSJob,
    [Switch]$AutoRemove,
    [Switch]$Force
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Receive-MyRSJob Begin Block"

    # Remove Invalid Get-MyRSJob Parameters
    if ($PSCmdlet.ParameterSetName -ne "RSJob")
    {
      if ($PSBoundParameters.ContainsKey("AutoRemove"))
      {
        [Void]$PSBoundParameters.Remove("AutoRemove")
      }
    }

    # List for Remove Jobs
    $RemoveJobs = [System.Collections.Generic.List[MyRSJob]]::New()

    Write-Verbose -Message "Exit Function Receive-MyRSJob Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function Receive-MyRSJob Process Block"

    # Add Passed RSJobs to $Jobs
    if ($PSCmdlet.ParameterSetName -eq "RSJob")
    {
      $TempJobs = $RSJob
    }
    else
    {
      [Void]$PSBoundParameters.Add("State", "Completed")
      $TempJobs = [MyRSJob[]](Get-MyRSJob @PSBoundParameters)
    }

    # Receive all Complted Jobs, Remove Job if Required
    ForEach ($TempJob in $TempJobs)
    {
      if ($TempJob.IsCompleted)
      {
        Try
        {
          $TempJob.PowerShell.EndInvoke($TempJob.PowerShellAsyncResult)
          # Add Job to Remove List
          [Void]$RemoveJobs.Add($TempJob)
        }
        Catch
        {
        }
      }
    }

    Write-Verbose -Message "Exit Function Receive-MyRSJob Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function Receive-MyRSJob End Block"

    if ($AutoRemove.IsPresent)
    {
      # Remove RSJobs
      foreach ($RemoveJob in $RemoveJobs)
      {
        $RemoveJob.PowerShell.Dispose()
        [Void]$Script:MyHiddenRSPool[$RemoveJob.PoolName].Jobs.Remove($RemoveJob)
      }
      $RemoveJobs.Clear()
    }

    # Garbage Collect, Recover Resources
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Function Receive-MyRSJob End Block"
  }
}
#endregion function Receive-MyRSJob

#region function Remove-MyRSJob
function Remove-MyRSJob()
{
  <#
    .SYNOPSIS
      Function to do something specific
    .DESCRIPTION
      Function to do something specific
    .PARAMETER RSPool
      RunspacePool to search
    .PARAMETER Name
      Name of Job to search for
    .PARAMETER InstanceId
      InstanceId of Job to search for
    .PARAMETER RSJob
      RunspacePool Jobs to Process
    .PARAMETER State
      State of Jobs to search for
    .PARAMETER Force
      Force the Job to stop
    .EXAMPLE
      Remove-MyRSJob

      Remove all RSJobs in the Default RSPool
    .EXAMPLE
      Remove-MyRSJob -RSPool $RSPool

      Remove-MyRSJob -PoolName $PoolName

      Remove-MyRSJob -PoolID $PoolID

      Remove all RSJobs in the Specified RSPool
    .NOTES
      Original Script By Ken Sweet on 10/15/2017 at 06:53 AM
      Updated Script By Ken Sweet on 02/04/2019 at 06:53 AM
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "JobNamePoolName")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePool")]
    [MyRSPool[]]$RSPool,
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [String]$PoolName = "MyDefaultRSPool",
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolID")]
    [parameter(Mandatory = $True, ParameterSetName = "JobNamePoolID")]
    [Guid]$PoolID,
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolID")]
    [String[]]$JobName = ".*",
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $True, ParameterSetName = "JobIDPoolID")]
    [Guid[]]$JobID,
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ParameterSetName = "RSJob")]
    [MyRSJob[]]$RSJob,
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobNamePoolID")]
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPool")]
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolName")]
    [parameter(Mandatory = $False, ParameterSetName = "JobIDPoolID")]
    [ValidateSet("NotStarted", "Running", "Stopping", "Stopped", "Completed", "Failed", "Disconnected")]
    [String[]]$State,
    [Switch]$Force
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Remove-MyRSJob Begin Block"

    # Remove Invalid Get-MyRSJob Parameters
    if ($PSCmdlet.ParameterSetName -ne "RSJob")
    {
      if ($PSBoundParameters.ContainsKey("Force"))
      {
        [Void]$PSBoundParameters.Remove("Force")
      }
    }

    # List for Remove Jobs
    $RemoveJobs = [System.Collections.Generic.List[MyRSJob]]::New()

    Write-Verbose -Message "Exit Function Remove-MyRSJob Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function Remove-MyRSJob Process Block"

    # Add Passed RSJobs to $Jobs
    if ($PSCmdlet.ParameterSetName -eq "RSJob")
    {
      $TempJobs = $RSJob
    }
    else
    {
      $TempJobs = [MyRSJob[]](Get-MyRSJob @PSBoundParameters)
    }

    # Remove all Jobs, Stop all Running if Forced
    ForEach ($TempJob in $TempJobs)
    {
      if ($Force -and $TempJob.State -notmatch "Stopped|Completed|Failed")
      {
        $TempJob.PowerShell.Stop()
      }
      if ($TempJob.State -match "Stopped|Completed|Failed")
      {
        # Add Job to Remove List
        [Void]$RemoveJobs.Add($TempJob)
      }
    }

    Write-Verbose -Message "Exit Function Remove-MyRSJob Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function Remove-MyRSJob End Block"

    # Remove RSJobs
    foreach ($RemoveJob in $RemoveJobs)
    {
      $RemoveJob.PowerShell.Dispose()
      [Void]$Script:MyHiddenRSPool[$RemoveJob.PoolName].Jobs.Remove($RemoveJob)
    }
    $RemoveJobs.Clear()

    # Garbage Collect, Recover Resources
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Function Remove-MyRSJob End Block"
  }
}
#endregion function Remove-MyRSJob

#endregion ******** Multiple Thread Functions ********

#region ******** PIL Common Dialogs ********

# --------------------------
# Show ChangeLog Function
# --------------------------
#region function Show-ChangeLog
Function Show-ChangeLog ()
{
  <#
    .SYNOPSIS
      Shows Show-ChangeLog
    .DESCRIPTION
      Shows Show-ChangeLog
    .PARAMETER Title
      Title of the Show-ChangeLog Dialog Window
    .PARAMETER ChangeText
      Change Log Text
    .PARAMETER Width
      Width of Show-ChangeLog Dialog Window
    .PARAMETER Height
      Height of Show-ChangeLog Dialog Window
    .EXAMPLE
      $TmpContent = ($Script:MyInvocation.MyCommand.ScriptBlock).ToString()
      $CLogStart = ($TmpContent.IndexOf("<") + 3)
      $CLogEnd = ($TmpContent.IndexOf(">") - 1)
      Show-ChangeLog -ChangeText ($TmpContent.SubString($CLogStart, ($CLogEnd - $CLogStart)))
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  Param (
    [String]$Title = "Change Log - $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)",
    [parameter(Mandatory = $True)]
    [String]$ChangeText,
    [Int]$Width = 60,
    [Int]$Height = 30
  )
  Write-Verbose -Message "Enter Function Show-ChangeLog"

  #region ******** Begin **** ChangeLog **** Begin ********

  # ************************************************
  # ChangeLog Form
  # ************************************************
  #region $ChangeLogForm = [System.Windows.Forms.Form]::New()
  $ChangeLogForm = [System.Windows.Forms.Form]::New()
  $ChangeLogForm.BackColor = [MyConfig]::Colors.Back
  $ChangeLogForm.Font = [MyConfig]::Font.Regular
  $ChangeLogForm.ForeColor = [MyConfig]::Colors.Fore
  $ChangeLogForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $ChangeLogForm.Icon = $PILForm.Icon
  $ChangeLogForm.KeyPreview = $True
  $ChangeLogForm.MaximizeBox = $False
  $ChangeLogForm.MinimizeBox = $False
  $ChangeLogForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * $Height))
  $ChangeLogForm.Name = "ChangeLogForm"
  $ChangeLogForm.Owner = $PILForm
  $ChangeLogForm.ShowInTaskbar = $False
  $ChangeLogForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $ChangeLogForm.Tag = $False
  $ChangeLogForm.Text = $Title
  #endregion $ChangeLogForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-ChangeLogFormKeyDown ********
  Function Start-ChangeLogFormKeyDown
  {
  <#
    .SYNOPSIS
      KeyDown Event for the ChangeLog Form Control
    .DESCRIPTION
      KeyDown Event for the ChangeLog Form Control
    .PARAMETER Sender
       The Form Control that fired the KeyDown Event
    .PARAMETER EventArg
       The Event Arguments for the Form KeyDown Event
    .EXAMPLE
       Start-ChangeLogFormKeyDown -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$ChangeLogForm"

    [MyConfig]::AutoExit = 0

    If ($EventArg.KeyCode -in ([System.Windows.Forms.Keys]::Enter, [System.Windows.Forms.Keys]::Space, [System.Windows.Forms.Keys]::Escape))
    {
      $ChangeLogForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }

    Write-Verbose -Message "Exit KeyDown Event for `$ChangeLogForm"
  }
  #endregion ******** Function Start-ChangeLogFormKeyDown ********
  $ChangeLogForm.add_KeyDown({ Start-ChangeLogFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ChangeLogFormShown ********
  Function Start-ChangeLogFormShown
  {
    <#
     .SYNOPSIS
       Shown Event for the ChangeLog Form Control
     .DESCRIPTION
       Shown Event for the ChangeLog Form Control
     .PARAMETER Sender
        The Form Control that fired the Shown Event
     .PARAMETER EventArg
         The Event Arguments for the Form Shown Event
      .EXAMPLE
         Start-ChangeLogFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$ChangeLogForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    $ChangeLogTextBox.AppendText($ChangeText)

    $ChangeLogTextBox.SelectionLength = 0
    $ChangeLogTextBox.SelectionStart = 0
    $ChangeLogTextBox.ScrollToCaret()
    $Sender.Refresh()
    $Sender.Activate()
    [System.Windows.Forms.Application]::DoEvents()

    Write-Verbose -Message "Exit Shown Event for `$ChangeLogForm"
  }
  #endregion ******** Function Start-ChangeLogFormShown ********
  $ChangeLogForm.add_Shown({ Start-ChangeLogFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for ChangeLog Form ********

  # ************************************************
  # ChangeLog Panel
  # ************************************************
  #region $ChangeLogPanel = [System.Windows.Forms.Panel]::New()
  $ChangeLogPanel = [System.Windows.Forms.Panel]::New()
  $ChangeLogForm.Controls.Add($ChangeLogPanel)
  $ChangeLogPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ChangeLogPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $ChangeLogPanel.Name = "ChangeLogPanel"
  #endregion $ChangeLogPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ChangeLogPanel Controls ********

  #region $ChangeLogTextBox = [System.Windows.Forms.TextBox]::New()
  $ChangeLogTextBox = [System.Windows.Forms.TextBox]::New()
  $ChangeLogPanel.Controls.Add($ChangeLogTextBox)
  $ChangeLogTextBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
  $ChangeLogTextBox.BackColor = [MyConfig]::Colors.TextBack
  $ChangeLogTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
  $ChangeLogTextBox.Font = [System.Drawing.Font]::New("Courier New", [MyConfig]::FontSize, [System.Drawing.FontStyle]::Regular)
  $ChangeLogTextBox.ForeColor = [MyConfig]::Colors.TextFore
  $ChangeLogTextBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $ChangeLogTextBox.MaxLength = [Int]::MaxValue
  $ChangeLogTextBox.Multiline = $True
  $ChangeLogTextBox.Name = "ChangeLogTextBox"
  $ChangeLogTextBox.ReadOnly = $True
  $ChangeLogTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
  $ChangeLogTextBox.Size = [System.Drawing.Size]::New(($ChangeLogPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($ChangeLogPanel.ClientSize.Height - ($ChangeLogTextBox.Top + [MyConfig]::FormSpacer)))
  $ChangeLogTextBox.TabStop = $False
  $ChangeLogTextBox.Text = $Null
  $ChangeLogTextBox.WordWrap = $False
  #endregion $ChangeLogTextBox = [System.Windows.Forms.TextBox]::New()

  #endregion ******** $ChangeLogPanel Controls ********

  # ************************************************
  # ChangeLogBtm Panel
  # ************************************************
  #region $ChangeLogBtmPanel = [System.Windows.Forms.Panel]::New()
  $ChangeLogBtmPanel = [System.Windows.Forms.Panel]::New()
  $ChangeLogForm.Controls.Add($ChangeLogBtmPanel)
  $ChangeLogBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ChangeLogBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $ChangeLogBtmPanel.Name = "ChangeLogBtmPanel"
  #endregion $ChangeLogBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ChangeLogBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($ChangeLogBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $ChangeLogBtmMidButton = [System.Windows.Forms.Button]::New()
  $ChangeLogBtmMidButton = [System.Windows.Forms.Button]::New()
  $ChangeLogBtmPanel.Controls.Add($ChangeLogBtmMidButton)
  $ChangeLogBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $ChangeLogBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ChangeLogBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ChangeLogBtmMidButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
  $ChangeLogBtmMidButton.Enabled = $True
  $ChangeLogBtmMidButton.Font = [MyConfig]::Font.Bold
  $ChangeLogBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ChangeLogBtmMidButton.Location = [System.Drawing.Point]::New(($TempWidth + ([MyConfig]::FormSpacer * 2)), 0)
  $ChangeLogBtmMidButton.Name = "ChangeLogBtmMidButton"
  $ChangeLogBtmMidButton.TabStop = $True
  $ChangeLogBtmMidButton.Text = "&Ok"
  $ChangeLogBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $ChangeLogBtmMidButton.PreferredSize.Height)
  #endregion $ChangeLogBtmMidButton = [System.Windows.Forms.Button]::New()

  $ChangeLogBtmPanel.ClientSize = [System.Drawing.Size]::New(($ChangeLogTextBox.Right + [MyConfig]::FormSpacer), (($ChangeLogBtmPanel.Controls[$ChangeLogBtmPanel.Controls.Count - 1]).Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ChangeLogBtmPanel Controls ********

  #endregion ******** Controls for ChangeLog Form ********

  #endregion ******** End **** Show-ChangeLog **** End ********

  $DialogResult = $ChangeLogForm.ShowDialog($PILForm)
  $ChangeLogForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Show-ChangeLog"
}
#endregion function Show-ChangeLog

# --------------------------
# Get UserResponse Function
# --------------------------
#region UserResponse Result Class
Class UserResponse
{
  [Bool]$Success
  [Object]$DialogResult
  
  UserResponse ([Bool]$Success, [Object]$DialogResult)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
  }
}
#endregion UserResponse Result Class

#region function Get-UserResponse
Function Get-UserResponse ()
{
  <#
    .SYNOPSIS
      Shows Get-UserResponse
    .DESCRIPTION
      Shows Get-UserResponse
    .PARAMETER Title
      Title of the Get-UserResponse Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER Width
      Width of the Get-UserResponse Dialog Window
    .PARAMETER Icon
      Message Icon
    .PARAMETER ButtonDefault
      The Default Button
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $Response = Get-UserResponse -Title "Get User Text - Single" -Message "Show this Sample Message Prompt to the User"
      if ($Response.Success)
      {
        # Success
      }
      else
      {
        # Failed
      }
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "One")]
  Param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [parameter(Mandatory = $True)]
    [String]$Message,
    [Int]$Width = 25,
    [System.Drawing.Icon]$Icon = [System.Drawing.SystemIcons]::Information,
    [System.Windows.Forms.DialogResult]$ButtonDefault = "OK",
    [parameter(Mandatory = $True, ParameterSetName = "Two")]
    [parameter(Mandatory = $True, ParameterSetName = "Three")]
    [System.Windows.Forms.DialogResult]$ButtonLeft,
    [parameter(Mandatory = $False, ParameterSetName = "One")]
    [parameter(Mandatory = $True, ParameterSetName = "Three")]
    [System.Windows.Forms.DialogResult]$ButtonMid = "OK",
    [parameter(Mandatory = $True, ParameterSetName = "Two")]
    [parameter(Mandatory = $True, ParameterSetName = "Three")]
    [System.Windows.Forms.DialogResult]$ButtonRight
  )
  Write-Verbose -Message "Enter Function Get-UserResponse"
  
  #region ******** Begin **** $UserResponse **** Begin ********
  
  # ************************************************
  # $UserResponse Form
  # ************************************************
  #region $UserResponseForm = [System.Windows.Forms.Form]::New()
  $UserResponseForm = [System.Windows.Forms.Form]::New()
  $UserResponseForm.BackColor = [MyConfig]::Colors.Back
  $UserResponseForm.Font = [MyConfig]::Font.Regular
  $UserResponseForm.ForeColor = [MyConfig]::Colors.Fore
  $UserResponseForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $UserResponseForm.Icon = $PILForm.Icon
  $UserResponseForm.KeyPreview = $AllowControl.IsPresent
  $UserResponseForm.MaximizeBox = $False
  $UserResponseForm.MinimizeBox = $False
  $UserResponseForm.Name = "UserResponseForm"
  $UserResponseForm.Owner = $PILForm
  $UserResponseForm.ShowInTaskbar = $False
  $UserResponseForm.Size = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * 25))
  $UserResponseForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $UserResponseForm.Tag = @{ "Cancel" = $False; "Pause" = $False }
  $UserResponseForm.Text = $Title
  #endregion $UserResponseForm = [System.Windows.Forms.Form]::New()
  
  #region ******** Function Start-UserResponseFormKeyDown ********
  Function Start-UserResponseFormKeyDown
  {
  <#
    .SYNOPSIS
      KeyDown Event for the UserResponse Form Control
    .DESCRIPTION
      KeyDown Event for the UserResponse Form Control
    .PARAMETER Sender
       The Form Control that fired the KeyDown Event
    .PARAMETER EventArg
       The Event Arguments for the Form KeyDown Event
    .EXAMPLE
       Start-UserResponseFormKeyDown -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$UserResponseForm"
    
    [MyConfig]::AutoExit = 0
    
    If ($EventArg.KeyCode -in ([System.Windows.Forms.Keys]::Escape))
    {
      $UserResponseForm.Close()
    }
    
    Write-Verbose -Message "Exit KeyDown Event for `$UserResponseForm"
  }
  #endregion ******** Function Start-UserResponseFormKeyDown ********
  $UserResponseForm.add_KeyDown({ Start-UserResponseFormKeyDown -Sender $This -EventArg $PSItem })
  
  #region ******** Function Start-UserResponseFormShown ********
  Function Start-UserResponseFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the $UserResponse Form Control
      .DESCRIPTION
        Shown Event for the $UserResponse Form Control
      .PARAMETER Sender
         The Form Control that fired the Shown Event
      .PARAMETER EventArg
         The Event Arguments for the Form Shown Event
      .EXAMPLE
         Start-UserResponseFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$UserResponseForm"
    
    [MyConfig]::AutoExit = 0
    
    $Sender.Refresh()
    
    Write-Verbose -Message "Exit Shown Event for `$UserResponseForm"
  }
  #endregion ******** Function Start-UserResponseFormShown ********
  $UserResponseForm.add_Shown({ Start-UserResponseFormShown -Sender $This -EventArg $PSItem })
  
  #region ******** Controls for $UserResponse Form ********
  
  # ************************************************
  # $UserResponse Panel
  # ************************************************
  #region $UserResponsePanel = [System.Windows.Forms.Panel]::New()
  $UserResponsePanel = [System.Windows.Forms.Panel]::New()
  $UserResponseForm.Controls.Add($UserResponsePanel)
  $UserResponsePanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $UserResponsePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $UserResponsePanel.Name = "UserResponsePanel"
  #endregion $UserResponsePanel = [System.Windows.Forms.Panel]::New()
  
  #region ******** $UserResponsePanel Controls ********
  
  #region $UserResponsePictureBox = [System.Windows.Forms.PictureBox]::New()
  $UserResponsePictureBox = [System.Windows.Forms.PictureBox]::New()
  $UserResponsePanel.Controls.Add($UserResponsePictureBox)
  $UserResponsePictureBox.AutoSize = $False
  $UserResponsePictureBox.BackColor = [MyConfig]::Colors.Back
  $UserResponsePictureBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $UserResponsePictureBox.Image = $Icon
  $UserResponsePictureBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
  $UserResponsePictureBox.Name = "UserResponsePictureBox"
  $UserResponsePictureBox.Size = [System.Drawing.Size]::New(32, 32)
  $UserResponsePictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::CenterImage
  #endregion $UserResponsePictureBox = [System.Windows.Forms.PictureBox]::New()
  
  #region $UserResponseLabel = [System.Windows.Forms.Label]::New()
  $UserResponseLabel = [System.Windows.Forms.Label]::New()
  $UserResponsePanel.Controls.Add($UserResponseLabel)
  $UserResponseLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $UserResponseLabel.Font = [MyConfig]::Font.Regular
  $UserResponseLabel.ForeColor = [MyConfig]::Colors.LabelFore
  $UserResponseLabel.Location = [System.Drawing.Point]::New(($UserResponsePictureBox.Right + [MyConfig]::FormSpacer), $UserResponsePictureBox.Top)
  $UserResponseLabel.Name = "UserResponseLabel"
  $UserResponseLabel.Size = [System.Drawing.Size]::New(($UserResponsePanel.ClientSize.Width - ($UserResponseLabel.Left + ([MyConfig]::FormSpacer * 3))), $UserResponsePanel.ClientSize.Width)
  $UserResponseLabel.Text = $Message
  $UserResponseLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  #endregion $UserResponseLabel = [System.Windows.Forms.Label]::New()
  
  # Returns the minimum size required to display the text
  $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($UserResponseLabel.Text, [MyConfig]::Font.Regular, $UserResponseLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
  $UserResponseLabel.Size = [System.Drawing.Size]::New(($UserResponsePanel.ClientSize.Width - ($UserResponseLabel.Left + ([MyConfig]::FormSpacer * 3))), ($TmpSize.Height + [MyConfig]::Font.Height))
  
  #endregion ******** $UserResponsePanel Controls ********
  
  Switch ($PSCmdlet.ParameterSetName)
  {
    "One"
    {
      $UserResponseButtons = 1
      Break
    }
    "Two"
    {
      $UserResponseButtons = 2
      Break
    }
    "Three"
    {
      $UserResponseButtons = 3
      Break
    }
  }
  
  # Evenly Space Buttons - Move Size to after Text
  # ************************************************
  # $UserResponseBtm Panel
  # ************************************************
  #region $UserResponseBtmPanel = [System.Windows.Forms.Panel]::New()
  $UserResponseBtmPanel = [System.Windows.Forms.Panel]::New()
  $UserResponseForm.Controls.Add($UserResponseBtmPanel)
  $UserResponseBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $UserResponseBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $UserResponseBtmPanel.Name = "UserResponseBtmPanel"
  #endregion $UserResponseBtmPanel = [System.Windows.Forms.Panel]::New()
  
  #region ******** $UserResponseBtmPanel Controls ********
  
  $NumButtons = 3
  $TempSpace = [Math]::Floor($UserResponseBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons
  
  #region $UserResponseBtmLeftButton = [System.Windows.Forms.Button]::New()
  If (($UserResponseButtons -eq 2) -or ($UserResponseButtons -eq 3))
  {
    $UserResponseBtmLeftButton = [System.Windows.Forms.Button]::New()
    $UserResponseBtmPanel.Controls.Add($UserResponseBtmLeftButton)
    $UserResponseBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
    $UserResponseBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $UserResponseBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
    $UserResponseBtmLeftButton.DialogResult = $ButtonLeft
    $UserResponseBtmLeftButton.Font = [MyConfig]::Font.Bold
    $UserResponseBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
    $UserResponseBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
    $UserResponseBtmLeftButton.Name = "UserResponseBtmLeftButton"
    $UserResponseBtmLeftButton.TabIndex = 0
    $UserResponseBtmLeftButton.TabStop = $True
    $UserResponseBtmLeftButton.Text = "&$($ButtonLeft.ToString())"
    $UserResponseBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $UserResponseBtmLeftButton.PreferredSize.Height)
    if ($ButtonLeft -eq $ButtonDefault)
    {
      $UserResponseBtmLeftButton.Select()
    }
  }
  #endregion $UserResponseBtmLeftButton = [System.Windows.Forms.Button]::New()
  
  #region $UserResponseBtmMidButton = [System.Windows.Forms.Button]::New()
  If (($UserResponseButtons -eq 1) -or ($UserResponseButtons -eq 3))
  {
    $UserResponseBtmMidButton = [System.Windows.Forms.Button]::New()
    $UserResponseBtmPanel.Controls.Add($UserResponseBtmMidButton)
    $UserResponseBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
    $UserResponseBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $UserResponseBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
    $UserResponseBtmMidButton.DialogResult = $ButtonMid
    $UserResponseBtmMidButton.Font = [MyConfig]::Font.Bold
    $UserResponseBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
    $UserResponseBtmMidButton.Location = [System.Drawing.Point]::New(($TempWidth + ([MyConfig]::FormSpacer * 2)), [MyConfig]::FormSpacer)
    $UserResponseBtmMidButton.Name = "UserResponseBtmMidButton"
    $UserResponseBtmMidButton.TabStop = $True
    $UserResponseBtmMidButton.Text = "&$($ButtonMid.ToString())"
    $UserResponseBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $UserResponseBtmMidButton.PreferredSize.Height)
    if ($ButtonMid -eq $ButtonDefault)
    {
      $UserResponseBtmMidButton.Select()
    }
  }
  #endregion $UserResponseBtmMidButton = [System.Windows.Forms.Button]::New()
  
  #region $UserResponseBtmRightButton = [System.Windows.Forms.Button]::New()
  If (($UserResponseButtons -eq 2) -or ($UserResponseButtons -eq 3))
  {
    $UserResponseBtmRightButton = [System.Windows.Forms.Button]::New()
    $UserResponseBtmPanel.Controls.Add($UserResponseBtmRightButton)
    $UserResponseBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
    $UserResponseBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $UserResponseBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
    $UserResponseBtmRightButton.DialogResult = $ButtonRight
    $UserResponseBtmRightButton.Font = [MyConfig]::Font.Bold
    $UserResponseBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
    $UserResponseBtmRightButton.Location = [System.Drawing.Point]::New(($UserResponseBtmLeftButton.Right + $TempWidth + $TempMod + ([MyConfig]::FormSpacer * 2)), [MyConfig]::FormSpacer)
    $UserResponseBtmRightButton.Name = "UserResponseBtmRightButton"
    $UserResponseBtmRightButton.TabIndex = 1
    $UserResponseBtmRightButton.TabStop = $True
    $UserResponseBtmRightButton.Text = "&$($ButtonRight.ToString())"
    $UserResponseBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $UserResponseBtmRightButton.PreferredSize.Height)
    if ($ButtonRight -eq $ButtonDefault)
    {
      $UserResponseBtmRightButton.Select()
    }
  }
  #endregion $UserResponseBtmRightButton = [System.Windows.Forms.Button]::New()
  
  $UserResponseBtmPanel.ClientSize = [System.Drawing.Size]::New(($UserResponseTextBox.Right + [MyConfig]::FormSpacer), (($UserResponseBtmPanel.Controls[$UserResponseBtmPanel.Controls.Count - 1]).Bottom + [MyConfig]::FormSpacer))
  
  #endregion ******** $UserResponseBtmPanel Controls ********
  
  $UserResponseForm.ClientSize = [System.Drawing.Size]::New($UserResponseForm.ClientSize.Width, ($UserResponseForm.ClientSize.Height - ($UserResponsePanel.ClientSize.Height - ([Math]::Max($UserResponsePictureBox.Bottom, $UserResponseLabel.Bottom) + ([MyConfig]::FormSpacer * 2)))))
  
  #endregion ******** Controls for $UserResponse Form ********
  
  #endregion ******** End **** $Get-UserResponse **** End ********
  
  $DialogResult = $UserResponseForm.ShowDialog($PILForm)
  [UserResponse]::New(($DialogResult -eq $ButtonDefault), $DialogResult)
  
  $UserResponseForm.Dispose()
  
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
  
  Write-Verbose -Message "Exit Function Get-UserResponse"
}
#endregion function Get-UserResponse

# -----------------------
# Get TextBoxInput Function
# -----------------------
#region TextBoxInput Result Class
Class TextBoxInput
{
  [Bool]$Success
  [Object]$DialogResult
  [String[]]$Items

  TextBoxInput ([Bool]$Success, [Object]$DialogResult, [String[]]$Items)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.Items = $Items
  }
}
#endregion TextBoxInput Result Class

#region function Get-TextBoxInput
function Get-TextBoxInput ()
{
  <#
    .SYNOPSIS
      Shows Get-TextBoxInput
    .DESCRIPTION
      Shows Get-TextBoxInput
    .PARAMETER Title
      Title of the Get-TextBoxInput Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER HintText
      Hint Text to Display
    .PARAMETER ValidChars
      RegEx Allowed Valid Characters for Input
    .PARAMETER ValidOutput
      RegEx Validate Output Format
    .PARAMETER Items
      Default Items / Text
    .PARAMETER MaxLength
      Maximum Length of Text Input
    .PARAMETER Multi
      Allow Multiple Lines of TExt
    .PARAMETER NoDuplicates
      Do Not Allow Duplicate Values
    .PARAMETER Width
      Width of the Get-TextBoxInput Dialog Window
    .PARAMETER Height
      Height of the Get-TextBoxInput Dialog Window
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $DialogResult = Get-TextBoxInput -Title "Get User Text - Multi" -Message "Show this Sample Message Prompt to the User" -Multi -Items @("Computer Name 01", "Computer Name 02", "Computer Name 03")
      if ($DialogResult.Success)
      {
        # Success
      }
      else
      {
        # Failed
      }
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Single")]
  param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [String]$Message = "Status Message",
    [String]$HintText = "Enter Value Here",
    [String]$ValidChars = "[\s\w\d\.\-_,;]",
    [String]$ValidOutput = ".+",
    [String[]]$Items = @(),
    [Int]$MaxLength = [Int]::MaxValue,
    [Int]$Width = 35,
    [parameter(Mandatory = $True, ParameterSetName = "Multi")]
    [Switch]$Multi,
    [parameter(Mandatory = $True, ParameterSetName = "Multi")]
    [Switch]$NoDuplicates,
    [parameter(Mandatory = $False, ParameterSetName = "Multi")]
    [Int]$Height = 18,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel"
  )
  Write-Verbose -Message "Enter Function Get-TextBoxInput"

  #region ******** Begin **** TextBoxInput **** Begin ********

  # ************************************************
  # TextBoxInput Form
  # ************************************************
  #region $TextBoxInputForm = [System.Windows.Forms.Form]::New()
  $TextBoxInputForm = [System.Windows.Forms.Form]::New()
  $TextBoxInputForm.BackColor = [MyConfig]::Colors.Back
  $TextBoxInputForm.Font = [MyConfig]::Font.Regular
  $TextBoxInputForm.ForeColor = [MyConfig]::Colors.Fore
  $TextBoxInputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $TextBoxInputForm.Icon = $PILForm.Icon
  $TextBoxInputForm.KeyPreview = $True
  $TextBoxInputForm.MaximizeBox = $False
  $TextBoxInputForm.MinimizeBox = $False
  if ($Multi.IsPresent)
  {
    $TextBoxInputForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * $Height))
  }
  else
  {
    $TextBoxInputForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), 0)
  }
  $TextBoxInputForm.Name = "TextBoxInputForm"
  $TextBoxInputForm.Owner = $PILForm
  $TextBoxInputForm.ShowInTaskbar = $False
  $TextBoxInputForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $TextBoxInputForm.Text = $Title
  #endregion $TextBoxInputForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-TextBoxInputFormKeyDown ********
  function Start-TextBoxInputFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the TextBoxInput Form Control
      .DESCRIPTION
        KeyDown Event for the TextBoxInput Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-TextBoxInputFormKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$TextBoxInputForm"

    [MyConfig]::AutoExit = 0
    if ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $TextBoxInputForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$TextBoxInputForm"
  }
  #endregion ******** Function Start-TextBoxInputFormKeyDown ********
  $TextBoxInputForm.add_KeyDown({ Start-TextBoxInputFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-TextBoxInputFormShown ********
  function Start-TextBoxInputFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the TextBoxInput Form Control
      .DESCRIPTION
        Shown Event for the TextBoxInput Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-TextBoxInputFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$TextBoxInputForm"

    [MyConfig]::AutoExit = 0

    $TextBoxInputTextBox.DeselectAll()

    $Sender.Refresh()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Shown Event for `$TextBoxInputForm"
  }
  #endregion ******** Function Start-TextBoxInputFormShown ********
  $TextBoxInputForm.add_Shown({ Start-TextBoxInputFormShown -Sender $This -EventArg $PSItem })
  
  #region ******** Controls for TextBoxInput Form ********

  # ************************************************
  # TextBoxInput Panel
  # ************************************************
  #region $TextBoxInputPanel = [System.Windows.Forms.Panel]::New()
  $TextBoxInputPanel = [System.Windows.Forms.Panel]::New()
  $TextBoxInputForm.Controls.Add($TextBoxInputPanel)
  $TextBoxInputPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $TextBoxInputPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $TextBoxInputPanel.Name = "TextBoxInputPanel"
  #endregion $TextBoxInputPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $TextBoxInputPanel Controls ********

  if ($PSBoundParameters.ContainsKey("Message"))
  {
    #region $TextBoxInputLabel = [System.Windows.Forms.Label]::New()
    $TextBoxInputLabel = [System.Windows.Forms.Label]::New()
    $TextBoxInputPanel.Controls.Add($TextBoxInputLabel)
    $TextBoxInputLabel.ForeColor = [MyConfig]::Colors.LabelFore
    $TextBoxInputLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
    $TextBoxInputLabel.Name = "TextBoxInputLabel"
    $TextBoxInputLabel.Size = [System.Drawing.Size]::New(($TextBoxInputPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
    $TextBoxInputLabel.Text = $Message
    $TextBoxInputLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #endregion $TextBoxInputLabel = [System.Windows.Forms.Label]::New()

    # Returns the minimum size required to display the text
    $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($TextBoxInputLabel.Text, [MyConfig]::Font.Regular, $TextBoxInputLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
    $TextBoxInputLabel.Size = [System.Drawing.Size]::New(($TextBoxInputPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

    $TmpBottom = $TextBoxInputLabel.Bottom + [MyConfig]::FormSpacer
  }
  else
  {
    $TmpBottom = 0
  }
  
  # ************************************************
  # TextBoxInput GroupBox
  # ************************************************
  #region $TextBoxInputGroupBox = [System.Windows.Forms.GroupBox]::New()
  $TextBoxInputGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $TextBoxInputPanel.Controls.Add($TextBoxInputGroupBox)
  $TextBoxInputGroupBox.BackColor = [MyConfig]::Colors.Back
  $TextBoxInputGroupBox.Font = [MyConfig]::Font.Regular
  $TextBoxInputGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $TextBoxInputGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($TmpBottom + [MyConfig]::FormSpacer))
  $TextBoxInputGroupBox.Name = "TextBoxInputGroupBox"
  $TextBoxInputGroupBox.Size = [System.Drawing.Size]::New(($TextBoxInputPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TextBoxInputPanel.ClientSize.Height - ($TextBoxInputGroupBox.Top + [MyConfig]::FormSpacer)))
  $TextBoxInputGroupBox.Text = $Null
  #endregion $TextBoxInputGroupBox = [System.Windows.Forms.GroupBox]::New()
  
  #region ******** $TextBoxInputGroupBox Controls ********
  
  #region $TextBoxInputTextBox = [System.Windows.Forms.TextBox]::New()
  $TextBoxInputTextBox = [System.Windows.Forms.TextBox]::New()
  $TextBoxInputGroupBox.Controls.Add($TextBoxInputTextBox)
  $TextBoxInputTextBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
  $TextBoxInputTextBox.AutoSize = $True
  $TextBoxInputTextBox.BackColor = [MyConfig]::Colors.TextBack
  $TextBoxInputTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
  $TextBoxInputTextBox.Font = [MyConfig]::Font.Regular
  $TextBoxInputTextBox.ForeColor = [MyConfig]::Colors.TextFore
  $TextBoxInputTextBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $TextBoxInputTextBox.MaxLength = $MaxLength
  $TextBoxInputTextBox.Multiline = $Multi.IsPresent
  $TextBoxInputTextBox.Name = "TextBoxInputTextBox"
  if ($Multi.IsPresent)
  {
    $TextBoxInputTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    If ($Items.Count)
    {
      $TextBoxInputTextBox.Lines = $Items
      $TextBoxInputTextBox.Tag = @{ "HintText" = $HintText; "HintEnabled" = $False; "Items" = $Items }
    }
    Else
    {
      $TextBoxInputTextBox.Lines = ""
      $TextBoxInputTextBox.Tag = @{ "HintText" = $HintText; "HintEnabled" = $True; "Items" = $Items }
    }
    $TextBoxInputTextBox.Size = [System.Drawing.Size]::New(($TextBoxInputGroupBox.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TextBoxInputGroupBox.ClientSize.Height - ($TextBoxInputTextBox.Top + [MyConfig]::FormSpacer)))
  }
  else
  {
    $TextBoxInputTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::None
    if ($Items.Count)
    {
      $TextBoxInputTextBox.Text = $Items[0]
      $TextBoxInputTextBox.Tag = @{ "HintText" = $HintText; "HintEnabled" = $False; "Items" = $Items[0] } 
    }
    else
    {
      $TextBoxInputTextBox.Text = ""
      $TextBoxInputTextBox.Tag = @{ "HintText" = $HintText; "HintEnabled" = $True; "Items" = "" }
    }
    $TextBoxInputTextBox.Size = [System.Drawing.Size]::New(($TextBoxInputGroupBox.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), $TextBoxInputTextBox.PreferredHeight)
  }
  $TextBoxInputTextBox.TabIndex = 0
  $TextBoxInputTextBox.TabStop = $True
  $TextBoxInputTextBox.WordWrap = $False
  #endregion $TextBoxInputTextBox = [System.Windows.Forms.TextBox]::New()
  
  #region ******** Function Start-TextBoxInputTextBoxGotFocus ********
  Function Start-TextBoxInputTextBoxGotFocus
  {
  <#
    .SYNOPSIS
      GotFocus Event for the TextBoxInput TextBox Control
    .DESCRIPTION
      GotFocus Event for the TextBoxInput TextBox Control
    .PARAMETER Sender
       The TextBox Control that fired the GotFocus Event
    .PARAMETER EventArg
       The Event Arguments for the TextBox GotFocus Event
    .EXAMPLE
       Start-TextBoxInputTextBoxGotFocus -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter GotFocus Event for `$TextBoxInputTextBox"
    
    [MyConfig]::AutoExit = 0
    
    # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
    If ($Sender.Tag.HintEnabled)
    {
      $Sender.Text = ""
      $Sender.Font = [MyConfig]::Font.Regular
      $Sender.ForeColor = [MyConfig]::Colors.TextFore
    }
    
    Write-Verbose -Message "Exit GotFocus Event for `$TextBoxInputTextBox"
  }
  #endregion ******** Function Start-TextBoxInputTextBoxGotFocus ********
  $TextBoxInputTextBox.add_GotFocus({ Start-TextBoxInputTextBoxGotFocus -Sender $This -EventArg $PSItem })
  
  #region ******** Function Start-TextBoxInputTextBoxKeyDown ********
  function Start-TextBoxInputTextBoxKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the TextBoxInput TextBox Control
      .DESCRIPTION
        KeyDown Event for the TextBoxInput TextBox Control
      .PARAMETER Sender
        The TextBox Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the TextBox KeyDown Event
      .EXAMPLE
        Start-TextBoxInputTextBoxKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$TextBoxInputTextBox"

    [MyConfig]::AutoExit = 0
    
    if ((-not $Sender.Multiline) -and ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Return))
    {
      $TextBoxInputBtmLeftButton.PerformClick()
    }
    
    Write-Verbose -Message "Exit KeyDown Event for `$TextBoxInputTextBox"
  }
  #endregion ******** Function Start-TextBoxInputTextBoxKeyDown ********
  $TextBoxInputTextBox.add_KeyDown({ Start-TextBoxInputTextBoxKeyDown -Sender $This -EventArg $PSItem })
  
  #region ******** Function Start-TextBoxInputTextBoxKeyPress ********
  Function Start-TextBoxInputTextBoxKeyPress
  {
    <#
      .SYNOPSIS
        KeyPress Event for the TextBoxInput TextBox Control
      .DESCRIPTION
        KeyPress Event for the TextBoxInput TextBox Control
      .PARAMETER Sender
         The TextBox Control that fired the KeyPress Event
      .PARAMETER EventArg
         The Event Arguments for the TextBox KeyPress Event
      .EXAMPLE
         Start-TextBoxInputTextBoxKeyPress -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyPress Event for `$TextBoxInputTextBox"
    
    [MyConfig]::AutoExit = 0
    
    # 1 = Ctrl-A, 3 = Ctrl-C, 8 = Backspace, 22 = Ctrl-V, 24 = Ctrl-X
    $EventArg.Handled = (($EventArg.KeyChar -notmatch $ValidChars) -and ([Int]($EventArg.KeyChar) -notin (1, 3, 8, 22, 24)))
    
    Write-Verbose -Message "Exit KeyPress Event for `$TextBoxInputTextBox"
  }
  #endregion ******** Function Start-TextBoxInputTextBoxKeyPress ********
  $TextBoxInputTextBox.add_KeyPress({Start-TextBoxInputTextBoxKeyPress -Sender $This -EventArg $PSItem})
  
  #region ******** Function Start-TextBoxInputTextBoxKeyUp ********
  Function Start-TextBoxInputTextBoxKeyUp
  {
  <#
    .SYNOPSIS
      KeyUp Event for the TextBoxInput TextBox Control
    .DESCRIPTION
      KeyUp Event for the TextBoxInput TextBox Control
    .PARAMETER Sender
       The TextBox Control that fired the KeyUp Event
    .PARAMETER EventArg
       The Event Arguments for the TextBox KeyUp Event
    .EXAMPLE
       Start-TextBoxInputTextBoxKeyUp -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyUp Event for `$TextBoxInputTextBox"
    
    [MyConfig]::AutoExit = 0
    
    # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
    $Sender.Tag.HintEnabled = ($Sender.Text.Trim().Length -eq 0)
    
    Write-Verbose -Message "Exit KeyUp Event for `$TextBoxInputTextBox"
  }
  #endregion ******** Function Start-TextBoxInputTextBoxKeyUp ********
  $TextBoxInputTextBox.add_KeyUp({ Start-TextBoxInputTextBoxKeyUp -Sender $This -EventArg $PSItem })
  
  #region ******** Function Start-TextBoxInputTextBoxLostFocus ********
  Function Start-TextBoxInputTextBoxLostFocus
  {
  <#
    .SYNOPSIS
      LostFocus Event for the TextBoxInput TextBox Control
    .DESCRIPTION
      LostFocus Event for the TextBoxInput TextBox Control
    .PARAMETER Sender
       The TextBox Control that fired the LostFocus Event
    .PARAMETER EventArg
       The Event Arguments for the TextBox LostFocus Event
    .EXAMPLE
       Start-TextBoxInputTextBoxLostFocus -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter LostFocus Event for `$TextBoxInputTextBox"
    
    [MyConfig]::AutoExit = 0
    
    # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
    If ([String]::IsNullOrEmpty(($Sender.Text.Trim())))
    {
      $Sender.Text = $Sender.Tag.HintText
      $Sender.Tag.HintEnabled = $True
      $Sender.Font = [MyConfig]::Font.Hint
      $Sender.ForeColor = [MyConfig]::Colors.TextHint
    }
    Else
    {
      $Sender.Tag.HintEnabled = $False
      $Sender.Font = [MyConfig]::Font.Regular
      $Sender.ForeColor = [MyConfig]::Colors.TextFore
    }
    
    Write-Verbose -Message "Exit LostFocus Event for `$TextBoxInputTextBox"
  }
  #endregion ******** Function Start-TextBoxInputTextBoxLostFocus ********
  $TextBoxInputTextBox.add_LostFocus({ Start-TextBoxInputTextBoxLostFocus -Sender $This -EventArg $PSItem })
  
  $TextBoxInputGroupBox.ClientSize = [System.Drawing.Size]::New($TextBoxInputGroupBox.ClientSize.Width, ($TextBoxInputTextBox.Bottom + ([MyConfig]::FormSpacer * 2)))
  
  #endregion ******** $TextBoxInputGroupBox Controls ********
  
  $TempClientSize = [System.Drawing.Size]::New(($TextBoxInputGroupBox.Right + [MyConfig]::FormSpacer), ($TextBoxInputGroupBox.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $TextBoxInputPanel Controls ********

  # ************************************************
  # TextBoxInputBtm Panel
  # ************************************************
  #region $TextBoxInputBtmPanel = [System.Windows.Forms.Panel]::New()
  $TextBoxInputBtmPanel = [System.Windows.Forms.Panel]::New()
  $TextBoxInputForm.Controls.Add($TextBoxInputBtmPanel)
  $TextBoxInputBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $TextBoxInputBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $TextBoxInputBtmPanel.Name = "TextBoxInputBtmPanel"
  #endregion $TextBoxInputBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $TextBoxInputBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($TextBoxInputBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $TextBoxInputBtmLeftButton = [System.Windows.Forms.Button]::New()
  $TextBoxInputBtmLeftButton = [System.Windows.Forms.Button]::New()
  $TextBoxInputBtmPanel.Controls.Add($TextBoxInputBtmLeftButton)
  $TextBoxInputBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $TextBoxInputBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $TextBoxInputBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $TextBoxInputBtmLeftButton.Font = [MyConfig]::Font.Bold
  $TextBoxInputBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $TextBoxInputBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $TextBoxInputBtmLeftButton.Name = "TextBoxInputBtmLeftButton"
  $TextBoxInputBtmLeftButton.TabIndex = 1
  $TextBoxInputBtmLeftButton.TabStop = $True
  $TextBoxInputBtmLeftButton.Text = $ButtonLeft
  $TextBoxInputBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $TextBoxInputBtmLeftButton.PreferredSize.Height)
  #endregion $TextBoxInputBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-TextBoxInputBtmLeftButtonClick ********
  function Start-TextBoxInputBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the TextBoxInputBtmLeft Button Control
      .DESCRIPTION
        Click Event for the TextBoxInputBtmLeft Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-TextBoxInputBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$TextBoxInputBtmLeftButton"

    [MyConfig]::AutoExit = 0

    If ((-not $TextBoxInputTextBox.Tag.HintEnabled) -and ("$($TextBoxInputTextBox.Text.Trim())".Length -gt 0))
    {
      $ChkOutput = $True
      ($TextBoxInputTextBox.Text -replace "\s*[\n,;]+\s*", ",").Split(",", [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object -Process { $ChkOutput = ($ChkOutput -and $PSItem -match $ValidOutput) }
      If ($ChkOutput)
      {
        $TextBoxInputForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
      }
      Else
      {
        [Void][System.Windows.Forms.MessageBox]::Show($TextBoxInputForm, "Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
      }
    }
    Else
    {
      [Void][System.Windows.Forms.MessageBox]::Show($TextBoxInputForm, "Missing Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }

    Write-Verbose -Message "Exit Click Event for `$TextBoxInputBtmLeftButton"
  }
  #endregion ******** Function Start-TextBoxInputBtmLeftButtonClick ********
  $TextBoxInputBtmLeftButton.add_Click({ Start-TextBoxInputBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $TextBoxInputBtmMidButton = [System.Windows.Forms.Button]::New()
  $TextBoxInputBtmMidButton = [System.Windows.Forms.Button]::New()
  $TextBoxInputBtmPanel.Controls.Add($TextBoxInputBtmMidButton)
  $TextBoxInputBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $TextBoxInputBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $TextBoxInputBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $TextBoxInputBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $TextBoxInputBtmMidButton.Font = [MyConfig]::Font.Bold
  $TextBoxInputBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $TextBoxInputBtmMidButton.Location = [System.Drawing.Point]::New(($TextBoxInputBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $TextBoxInputBtmMidButton.Name = "TextBoxInputBtmMidButton"
  $TextBoxInputBtmMidButton.TabIndex = 2
  $TextBoxInputBtmMidButton.TabStop = $True
  $TextBoxInputBtmMidButton.Text = $ButtonMid
  $TextBoxInputBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $TextBoxInputBtmMidButton.PreferredSize.Height)
  #endregion $TextBoxInputBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-TextBoxInputBtmMidButtonClick ********
  function Start-TextBoxInputBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the TextBoxInputBtmMid Button Control
      .DESCRIPTION
        Click Event for the TextBoxInputBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-TextBoxInputBtmMidButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$TextBoxInputBtmMidButton"

    [MyConfig]::AutoExit = 0

    if ($Multi.IsPresent)
    {
      $TextBoxInputTextBox.Lines = $TextBoxInputTextBox.Tag.Items
    }
    else
    {
      $TextBoxInputTextBox.Text = $TextBoxInputTextBox.Tag.Items
    }
    
    $TextBoxInputTextBox.Tag.HintEnabled = ($TextBoxInputTextBox.TextLength -gt 0)
    Start-TextBoxInputTextBoxLostFocus -Sender $TextBoxInputTextBox -EventArg "LostFocus"
    
    Write-Verbose -Message "Exit Click Event for `$TextBoxInputBtmMidButton"
  }
  #endregion ******** Function Start-TextBoxInputBtmMidButtonClick ********
  $TextBoxInputBtmMidButton.add_Click({ Start-TextBoxInputBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $TextBoxInputBtmRightButton = [System.Windows.Forms.Button]::New()
  $TextBoxInputBtmRightButton = [System.Windows.Forms.Button]::New()
  $TextBoxInputBtmPanel.Controls.Add($TextBoxInputBtmRightButton)
  $TextBoxInputBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $TextBoxInputBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $TextBoxInputBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $TextBoxInputBtmRightButton.Font = [MyConfig]::Font.Bold
  $TextBoxInputBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $TextBoxInputBtmRightButton.Location = [System.Drawing.Point]::New(($TextBoxInputBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $TextBoxInputBtmRightButton.Name = "TextBoxInputBtmRightButton"
  $TextBoxInputBtmRightButton.TabIndex = 3
  $TextBoxInputBtmRightButton.TabStop = $True
  $TextBoxInputBtmRightButton.Text = $ButtonRight
  $TextBoxInputBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $TextBoxInputBtmRightButton.PreferredSize.Height)
  #endregion $TextBoxInputBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-TextBoxInputBtmRightButtonClick ********
  function Start-TextBoxInputBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the TextBoxInputBtmRight Button Control
      .DESCRIPTION
        Click Event for the TextBoxInputBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-TextBoxInputBtmRightButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$TextBoxInputBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here
    
    $TextBoxInputForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$TextBoxInputBtmRightButton"
  }
  #endregion ******** Function Start-TextBoxInputBtmRightButtonClick ********
  $TextBoxInputBtmRightButton.add_Click({ Start-TextBoxInputBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $TextBoxInputBtmPanel.ClientSize = [System.Drawing.Size]::New(($TextBoxInputBtmRightButton.Right + [MyConfig]::FormSpacer), ($TextBoxInputBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $TextBoxInputBtmPanel Controls ********

  $TextBoxInputForm.ClientSize = [System.Drawing.Size]::New($TextBoxInputForm.ClientSize.Width, ($TempClientSize.Height + $TextBoxInputBtmPanel.Height))

  #endregion ******** Controls for TextBoxInput Form ********

  #endregion ******** End **** Get-TextBoxInput **** End ********

  $DialogResult = $TextBoxInputForm.ShowDialog($PILForm)
  If ($Multi.IsPresent)
  {
    If ($NoDuplicates.IsPresent)
    {
      $TmpItems = @(($TextBoxInputTextBox.Text -replace "\s*[\n,;]+\s*", ",").Split(",", [System.StringSplitOptions]::RemoveEmptyEntries) | Select-Object -Unique)
    }
    Else
    {
      $TmpItems = @(($TextBoxInputTextBox.Text -replace "\s*[\n,;]+\s*", ",").Split(",", [System.StringSplitOptions]::RemoveEmptyEntries))
    }
    [TextBoxInput]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, $TmpItems)
  }
  Else
  {
    [TextBoxInput]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, $TextBoxInputTextBox.Text)
  }
  
  $TextBoxInputForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-TextBoxInput"
}
#endregion function Get-TextBoxInput

# ------------------------------
# Get MultiTextBoxInput Function
# ------------------------------
#region MultiTextBoxInput Result Class
Class MultiTextBoxInput
{
  [Bool]$Success
  [Object]$DialogResult
  [System.Collections.Specialized.OrderedDictionary]$OrderedItems

  MultiTextBoxInput ([Bool]$Success, [Object]$DialogResult, [System.Collections.Specialized.OrderedDictionary]$OrderedItems)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.OrderedItems = $OrderedItems
  }
}
#endregion MultiTextBoxInput Result Class

#region function Get-MultiTextBoxInput
Function Get-MultiTextBoxInput ()
{
  <#
    .SYNOPSIS
      Shows Get-MultiTextBoxInput
    .DESCRIPTION
      Shows Get-MultiTextBoxInput
    .PARAMETER Title
      Title of the Get-MultiTextBoxInput Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER ReturnTitle
      Title of Values Group Box
    .PARAMETER OrderedItems
      Ordered List (HashTable) if Names and Values
    .PARAMETER ValidCars
      Valid Inputy Chatacters
    .PARAMETER MaxLength
      Maximum Length of Text Input
    .PARAMETER Width
      With of Get-MultiTextBoxInput Dialog Window
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .PARAMETER AllRequired
      All Values are Required
    .EXAMPLE
      $DialogResult = Get-MultiTextBoxInput -Title "Get Multi Text Input" -Message "Show this Sample Message Prompt to the User" -OrderedItems $OrderedItems -AllRequired
      if ($DialogResult.Success)
      {
        # Success
      }
      else
      {
        # Failed
      }
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  Param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [String]$Message,
    [String]$ReturnTitle,
    [parameter(Mandatory = $True)]
    [System.Collections.Specialized.OrderedDictionary]$OrderedItems,
    [String]$ValidChars = "[\s\w\d\.\-_]",
    [Int]$MaxLength = [Int]::MaxValue,
    [Int]$Width = 35,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel",
    [Switch]$AllRequired
  )
  Write-Verbose -Message "Enter Function Get-MultiTextBoxInput"

  #region ******** Begin **** MultiTextBoxInput **** Begin ********

  # ************************************************
  # MultiTextBoxInput Form
  # ************************************************
  #region $MultiTextBoxInputForm = [System.Windows.Forms.Form]::New()
  $MultiTextBoxInputForm = [System.Windows.Forms.Form]::New()
  $MultiTextBoxInputForm.BackColor = [MyConfig]::Colors.Back
  $MultiTextBoxInputForm.Font = [MyConfig]::Font.Regular
  $MultiTextBoxInputForm.ForeColor = [MyConfig]::Colors.Fore
  $MultiTextBoxInputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $MultiTextBoxInputForm.Icon = $PILForm.Icon
  $MultiTextBoxInputForm.KeyPreview = $True
  $MultiTextBoxInputForm.MaximizeBox = $False
  $MultiTextBoxInputForm.MinimizeBox = $False
  $MultiTextBoxInputForm.Name = "MultiTextBoxInputForm"
  $MultiTextBoxInputForm.Owner = $PILForm
  $MultiTextBoxInputForm.ShowInTaskbar = $False
  $MultiTextBoxInputForm.Size = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * 25))
  $MultiTextBoxInputForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $MultiTextBoxInputForm.Tag = $AllRequired.IsPresent
  $MultiTextBoxInputForm.Text = $Title
  #endregion $MultiTextBoxInputForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-MultiTextBoxInputFormKeyDown ********
  Function Start-MultiTextBoxInputFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the MultiTextBoxInput Form Control
      .DESCRIPTION
        KeyDown Event for the MultiTextBoxInput Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-MultiTextBoxInputFormKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$MultiTextBoxInputForm"

    [MyConfig]::AutoExit = 0

    If ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $MultiTextBoxInputForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$MultiTextBoxInputForm"
  }
  #endregion ******** Function Start-MultiTextBoxInputFormKeyDown ********
  $MultiTextBoxInputForm.add_KeyDown({ Start-MultiTextBoxInputFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-MultiTextBoxInputFormShown ********
  Function Start-MultiTextBoxInputFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the MultiTextBoxInput Form Control
      .DESCRIPTION
        Shown Event for the MultiTextBoxInput Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-MultiTextBoxInputFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$MultiTextBoxInputForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Shown Event for `$MultiTextBoxInputForm"
  }
  #endregion ******** Function Start-MultiTextBoxInputFormShown ********
  $MultiTextBoxInputForm.add_Shown({ Start-MultiTextBoxInputFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for MultiTextBoxInput Form ********

  # ************************************************
  # MultiTextBoxInput Panel
  # ************************************************
  #region $MultiTextBoxInputPanel = [System.Windows.Forms.Panel]::New()
  $MultiTextBoxInputPanel = [System.Windows.Forms.Panel]::New()
  $MultiTextBoxInputForm.Controls.Add($MultiTextBoxInputPanel)
  $MultiTextBoxInputPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $MultiTextBoxInputPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $MultiTextBoxInputPanel.Name = "MultiTextBoxInputPanel"
  $MultiTextBoxInputPanel.Text = "MultiTextBoxInputPanel"
  #endregion $MultiTextBoxInputPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $MultiTextBoxInputPanel Controls ********

  If ($PSBoundParameters.ContainsKey("Message"))
  {
    #region $MultiTextBoxInputLabel = [System.Windows.Forms.Label]::New()
    $MultiTextBoxInputLabel = [System.Windows.Forms.Label]::New()
    $MultiTextBoxInputPanel.Controls.Add($MultiTextBoxInputLabel)
    $MultiTextBoxInputLabel.ForeColor = [MyConfig]::Colors.LabelFore
    $MultiTextBoxInputLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
    $MultiTextBoxInputLabel.Name = "SearchTextMainLabel"
    $MultiTextBoxInputLabel.Size = [System.Drawing.Size]::New(($MultiTextBoxInputPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
    $MultiTextBoxInputLabel.Text = $Message
    $MultiTextBoxInputLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #endregion $MultiTextBoxInputLabel = [System.Windows.Forms.Label]::New()

    # Returns the minimum size required to display the text
    $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($MultiTextBoxInputLabel.Text, [MyConfig]::Font.Regular, $MultiTextBoxInputLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
    $MultiTextBoxInputLabel.Size = [System.Drawing.Size]::New(($MultiTextBoxInputPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

    $TempBottom = $MultiTextBoxInputLabel.Bottom
  }
  Else
  {
    $TempBottom = 0
  }

  # ************************************************
  # MultiTextBoxInput GroupBox
  # ************************************************
  #region $MultiTextBoxInputGroupBox = [System.Windows.Forms.GroupBox]::New()
  $MultiTextBoxInputGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $MultiTextBoxInputPanel.Controls.Add($MultiTextBoxInputGroupBox)
  $MultiTextBoxInputGroupBox.BackColor = [MyConfig]::Colors.Back
  $MultiTextBoxInputGroupBox.Font = [MyConfig]::Font.Bold
  $MultiTextBoxInputGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $MultiTextBoxInputGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($TempBottom + [MyConfig]::FormSpacer))
  $MultiTextBoxInputGroupBox.Name = "MultiTextBoxInputGroupBox"
  $MultiTextBoxInputGroupBox.Text = $ReturnTitle
  $MultiTextBoxInputGroupBox.Width = ($MultiTextBoxInputPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2))
  #endregion $MultiTextBoxInputGroupBox = [System.Windows.Forms.GroupBox]::New()

  $TmpLabelWidth = 0
  $Count = 0
  ForEach ($Key In $OrderedItems.Keys)
  {
    #region $MultiTextBoxInputLabel = [System.Windows.Forms.Label]::New()
    $MultiTextBoxInputLabel = [System.Windows.Forms.Label]::New()
    $MultiTextBoxInputGroupBox.Controls.Add($MultiTextBoxInputLabel)
    $MultiTextBoxInputLabel.AutoSize = $True
    $MultiTextBoxInputLabel.BackColor = [MyConfig]::Colors.Back
    $MultiTextBoxInputLabel.Font = [MyConfig]::Font.Regular
    $MultiTextBoxInputLabel.ForeColor = [MyConfig]::Colors.Fore
    $MultiTextBoxInputLabel.Location = [System.Drawing.Size]::New([MyConfig]::FormSpacer, ([MyConfig]::Font.Height + (($MultiTextBoxInputLabel.PreferredHeight + [MyConfig]::FormSpacer) * $Count)))
    $MultiTextBoxInputLabel.Name = "$($Key)Label"
    $MultiTextBoxInputLabel.Tag = $Null
    $MultiTextBoxInputLabel.Text = "$($Key):"
    $MultiTextBoxInputLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    #endregion $MultiTextBoxInputLabel = [System.Windows.Forms.Label]::New()

    $TmpLabelWidth = [Math]::Max($TmpLabelWidth, $MultiTextBoxInputLabel.Width)
    $Count += 1
  }

  #region ******** Function Start-MultiTextBoxInputTextBoxGotFocus ********
  Function Start-MultiTextBoxInputTextBoxGotFocus
  {
  <#
    .SYNOPSIS
      GotFocus Event for the MultiTextBoxInput TextBox Control
    .DESCRIPTION
      GotFocus Event for the MultiTextBoxInput TextBox Control
    .PARAMETER Sender
       The TextBox Control that fired the GotFocus Event
    .PARAMETER EventArg
       The Event Arguments for the TextBox GotFocus Event
    .EXAMPLE
       Start-MultiTextBoxInputTextBoxGotFocus -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter GotFocus Event for `$MultiTextBoxInputTextBox"

    [MyConfig]::AutoExit = 0

    # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
    If ($Sender.Tag.HintEnabled)
    {
      $Sender.Text = ""
      $Sender.Font = [MyConfig]::Font.Regular
      $Sender.ForeColor = [MyConfig]::Colors.TextFore
    }

    Write-Verbose -Message "Exit GotFocus Event for `$MultiTextBoxInputTextBox"
  }
  #endregion ******** Function Start-MultiTextBoxInputTextBoxGotFocus ********

  #region ******** Function Start-MultiTextBoxInputTextBoxKeyDown ********
  function Start-MultiTextBoxInputTextBoxKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the GetMultiValueMain TextBox Control
      .DESCRIPTION
        KeyDown Event for the GetMultiValueMain TextBox Control
      .PARAMETER Sender
        The TextBox Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the TextBox KeyDown Event
      .EXAMPLE
        Start-MultiTextBoxInputTextBoxKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$MultiTextBoxInputTextBox"

    [MyConfig]::AutoExit = 0

    if ((-not $Sender.Multiline) -and ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Return))
    {
      $MultiTextBoxInputBtmLeftButton.PerformClick()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$MultiTextBoxInputTextBox"
  }
  #endregion ******** Function Start-MultiTextBoxInputTextBoxKeyDown ********

  #region ******** Function Start-MultiTextBoxInputTextBoxKeyPress ********
  Function Start-MultiTextBoxInputTextBoxKeyPress
  {
    <#
      .SYNOPSIS
        KeyPress Event for the MultiTextBoxInput TextBox Control
      .DESCRIPTION
        KeyPress Event for the MultiTextBoxInput TextBox Control
      .PARAMETER Sender
         The TextBox Control that fired the KeyPress Event
      .PARAMETER EventArg
         The Event Arguments for the TextBox KeyPress Event
      .EXAMPLE
         Start-MultiTextBoxInputTextBoxKeyPress -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyPress Event for `$MultiTextBoxInputTextBox"

    [MyConfig]::AutoExit = 0

    # 1 = Ctrl-A, 3 = Ctrl-C, 8 = Backspace, 22 = Ctrl-V, 24 = Ctrl-X
    $EventArg.Handled = (($EventArg.KeyChar -notmatch $ValidChars) -and ([Int]($EventArg.KeyChar) -notin (1, 3, 8, 22, 24)))

    Write-Verbose -Message "Exit KeyPress Event for `$MultiTextBoxInputTextBox"
  }
  #endregion ******** Function Start-MultiTextBoxInputTextBoxKeyPress ********

  #region ******** Function Start-MultiTextBoxInputTextBoxKeyUp ********
  Function Start-MultiTextBoxInputTextBoxKeyUp
  {
  <#
    .SYNOPSIS
      KeyUp Event for the MultiTextBoxInput TextBox Control
    .DESCRIPTION
      KeyUp Event for the MultiTextBoxInput TextBox Control
    .PARAMETER Sender
       The TextBox Control that fired the KeyUp Event
    .PARAMETER EventArg
       The Event Arguments for the TextBox KeyUp Event
    .EXAMPLE
       Start-MultiTextBoxInputTextBoxKeyUp -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyUp Event for `$MultiTextBoxInputTextBox"

    [MyConfig]::AutoExit = 0

    # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
    $Sender.Tag.HintEnabled = ($Sender.Text.Trim().Length -eq 0)

    Write-Verbose -Message "Exit KeyUp Event for `$MultiTextBoxInputTextBox"
  }
  #endregion ******** Function Start-MultiTextBoxInputTextBoxKeyUp ********

  #region ******** Function Start-MultiTextBoxInputTextBoxLostFocus ********
  Function Start-MultiTextBoxInputTextBoxLostFocus
  {
  <#
    .SYNOPSIS
      LostFocus Event for the MultiTextBoxInput TextBox Control
    .DESCRIPTION
      LostFocus Event for the MultiTextBoxInput TextBox Control
    .PARAMETER Sender
       The TextBox Control that fired the LostFocus Event
    .PARAMETER EventArg
       The Event Arguments for the TextBox LostFocus Event
    .EXAMPLE
       Start-MultiTextBoxInputTextBoxLostFocus -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter LostFocus Event for `$MultiTextBoxInputTextBox"

    [MyConfig]::AutoExit = 0

    # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
    If ([String]::IsNullOrEmpty(($Sender.Text.Trim())))
    {
      $Sender.Text = $Sender.Tag.HintText
      $Sender.Tag.HintEnabled = $True
      $Sender.Font = [MyConfig]::Font.Hint
      $Sender.ForeColor = [MyConfig]::Colors.TextHint
    }
    Else
    {
      $Sender.Tag.HintEnabled = $False
      $Sender.Font = [MyConfig]::Font.Regular
      $Sender.ForeColor = [MyConfig]::Colors.TextFore
    }

    Write-Verbose -Message "Exit LostFocus Event for `$MultiTextBoxInputTextBox"
  }
  #endregion ******** Function Start-MultiTextBoxInputTextBoxLostFocus ********

  ForEach ($Key In $OrderedItems.Keys)
  {
    $TmpLabel = $MultiTextBoxInputGroupBox.Controls["$($Key)Label"]
    $TmpLabel.AutoSize = $False
    $TmpLabel.Size = [System.Drawing.Size]::New($TmpLabelWidth, $TmpLabel.PreferredHeight)

    #region $MultiTextBoxInputTextBox = [System.Windows.Forms.TextBox]::New()
    $MultiTextBoxInputTextBox = [System.Windows.Forms.TextBox]::New()
    $MultiTextBoxInputGroupBox.Controls.Add($MultiTextBoxInputTextBox)
    $MultiTextBoxInputTextBox.AutoSize = $False
    $MultiTextBoxInputTextBox.BackColor = [MyConfig]::Colors.TextBack
    $MultiTextBoxInputTextBox.Font = [MyConfig]::Font.Regular
    $MultiTextBoxInputTextBox.ForeColor = [MyConfig]::Colors.TextFore
    $MultiTextBoxInputTextBox.Location = [System.Drawing.Size]::New(($TmpLabel.Right + [MyConfig]::FormSpacer), $TmpLabel.Top)
    $MultiTextBoxInputTextBox.MaxLength = $MaxLength
    $MultiTextBoxInputTextBox.Name = "$($Key)"
    $MultiTextBoxInputTextBox.TabStop = $True
    $MultiTextBoxInputTextBox.Text = $OrderedItems[$Key]
    $MultiTextBoxInputTextBox.Tag = @{ "HintText" = "Enter Value for '$($Key)'"; "HintEnabled" = ($MultiTextBoxInputTextBox.TextLength -eq 0); "Value" = $OrderedItems[$Key] }
    $MultiTextBoxInputTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Left
    $MultiTextBoxInputTextBox.Size = [System.Drawing.Size]::New(($MultiTextBoxInputGroupBox.ClientSize.Width - ($TmpLabel.Right + ([MyConfig]::FormSpacer) * 2)), $TmpLabel.Height)
    #endregion $MultiTextBoxInputTextBox = [System.Windows.Forms.TextBox]::New()

    $MultiTextBoxInputTextBox.add_GotFocus({ Start-MultiTextBoxInputTextBoxGotFocus -Sender $This -EventArg $PSItem})
    $MultiTextBoxInputTextBox.add_KeyDown({ Start-MultiTextBoxInputTextBoxKeyDown -Sender $This -EventArg $PSItem })
    $MultiTextBoxInputTextBox.add_KeyPress({ Start-MultiTextBoxInputTextBoxKeyPress -Sender $This -EventArg $PSItem })
    $MultiTextBoxInputTextBox.add_KeyUp({ Start-MultiTextBoxInputTextBoxKeyUp -Sender $This -EventArg $PSItem })
    $MultiTextBoxInputTextBox.add_LostFocus({ Start-MultiTextBoxInputTextBoxLostFocus -Sender $This -EventArg $PSItem })
    Start-MultiTextBoxInputTextBoxLostFocus -Sender $MultiTextBoxInputTextBox -EventArg $EventArg
  }

  $MultiTextBoxInputGroupBox.ClientSize = [System.Drawing.Size]::New($MultiTextBoxInputGroupBox.ClientSize.Width, (($MultiTextBoxInputGroupBox.Controls[$MultiTextBoxInputGroupBox.Controls.Count - 1]).Bottom + [MyConfig]::FormSpacer))

  $TempClientSize = [System.Drawing.Size]::New(($MultiTextBoxInputTextBox.Right + [MyConfig]::FormSpacer), ($MultiTextBoxInputGroupBox.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $MultiTextBoxInputPanel Controls ********

  # ************************************************
  # MultiTextBoxInputBtm Panel
  # ************************************************
  #region $MultiTextBoxInputBtmPanel = [System.Windows.Forms.Panel]::New()
  $MultiTextBoxInputBtmPanel = [System.Windows.Forms.Panel]::New()
  $MultiTextBoxInputForm.Controls.Add($MultiTextBoxInputBtmPanel)
  $MultiTextBoxInputBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $MultiTextBoxInputBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $MultiTextBoxInputBtmPanel.Name = "MultiTextBoxInputBtmPanel"
  $MultiTextBoxInputBtmPanel.Text = "MultiTextBoxInputBtmPanel"
  #endregion $MultiTextBoxInputBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $MultiTextBoxInputBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($MultiTextBoxInputBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $MultiTextBoxInputBtmLeftButton = [System.Windows.Forms.Button]::New()
  $MultiTextBoxInputBtmLeftButton = [System.Windows.Forms.Button]::New()
  $MultiTextBoxInputBtmPanel.Controls.Add($MultiTextBoxInputBtmLeftButton)
  $MultiTextBoxInputBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $MultiTextBoxInputBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $MultiTextBoxInputBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $MultiTextBoxInputBtmLeftButton.Font = [MyConfig]::Font.Bold
  $MultiTextBoxInputBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $MultiTextBoxInputBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $MultiTextBoxInputBtmLeftButton.Name = "MultiTextBoxInputBtmLeftButton"
  $MultiTextBoxInputBtmLeftButton.TabIndex = 1
  $MultiTextBoxInputBtmLeftButton.TabStop = $True
  $MultiTextBoxInputBtmLeftButton.Text = $ButtonLeft
  $MultiTextBoxInputBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $MultiTextBoxInputBtmLeftButton.PreferredSize.Height)
  #endregion $MultiTextBoxInputBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-MultiTextBoxInputBtmLeftButtonClick ********
  Function Start-MultiTextBoxInputBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the MultiTextBoxInputBtmLeft Button Control
      .DESCRIPTION
        Click Event for the MultiTextBoxInputBtmLeft Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-MultiTextBoxInputBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$MultiTextBoxInputBtmLeftButton"

    [MyConfig]::AutoExit = 0

    $TmpValidCheck = $MultiTextBoxInputForm.Tag
    ForEach ($Key In @($OrderedItems.Keys))
    {
      $TmpItemValue = "$($MultiTextBoxInputGroupBox.Controls[$Key].Text)".Trim()
      $ChkItemValue = (-not (([String]::IsNullOrEmpty($TmpItemValue) -or $MultiTextBoxInputGroupBox.Controls[$Key].Tag.HintEnabled)))
      if ($ChkItemValue)
      {
        $OrderedItems[$Key] = $TmpItemValue
      }
      else
      {
        $OrderedItems[$Key] = $Null
      }

      if ($MultiTextBoxInputForm.Tag)
      {
        $TmpValidCheck = $ChkItemValue -and $TmpValidCheck
      }
      else
      {
        $TmpValidCheck = $ChkItemValue -or $TmpValidCheck
      }
    }

    If ($TmpValidCheck)
    {
      $MultiTextBoxInputForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    Else
    {
      [Void][System.Windows.Forms.MessageBox]::Show($MultiTextBoxInputForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }

    Write-Verbose -Message "Exit Click Event for `$MultiTextBoxInputBtmLeftButton"
  }
  #endregion ******** Function Start-MultiTextBoxInputBtmLeftButtonClick ********
  $MultiTextBoxInputBtmLeftButton.add_Click({ Start-MultiTextBoxInputBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $MultiTextBoxInputBtmMidButton = [System.Windows.Forms.Button]::New()
  $MultiTextBoxInputBtmMidButton = [System.Windows.Forms.Button]::New()
  $MultiTextBoxInputBtmPanel.Controls.Add($MultiTextBoxInputBtmMidButton)
  $MultiTextBoxInputBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $MultiTextBoxInputBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $MultiTextBoxInputBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $MultiTextBoxInputBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $MultiTextBoxInputBtmMidButton.Font = [MyConfig]::Font.Bold
  $MultiTextBoxInputBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $MultiTextBoxInputBtmMidButton.Location = [System.Drawing.Point]::New(($MultiTextBoxInputBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $MultiTextBoxInputBtmMidButton.Name = "MultiTextBoxInputBtmMidButton"
  $MultiTextBoxInputBtmMidButton.TabIndex = 2
  $MultiTextBoxInputBtmMidButton.TabStop = $True
  $MultiTextBoxInputBtmMidButton.Text = $ButtonMid
  $MultiTextBoxInputBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $MultiTextBoxInputBtmMidButton.PreferredSize.Height)
  #endregion $MultiTextBoxInputBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-MultiTextBoxInputBtmMidButtonClick ********
  Function Start-MultiTextBoxInputBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the MultiTextBoxInputBtmMid Button Control
      .DESCRIPTION
        Click Event for the MultiTextBoxInputBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-MultiTextBoxInputBtmMidButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$MultiTextBoxInputBtmMidButton"

    [MyConfig]::AutoExit = 0

    ForEach ($Key In @($OrderedItems.Keys))
    {
      $MultiTextBoxInputGroupBox.Controls[$Key].Text = $MultiTextBoxInputGroupBox.Controls[$Key].Tag.Value
      $MultiTextBoxInputGroupBox.Controls[$Key].Tag.HintEnabled = ($MultiTextBoxInputGroupBox.TextLength -eq 0)
      Start-MultiTextBoxInputTextBoxLostFocus -Sender $MultiTextBoxInputGroupBox.Controls[$Key] -EventArg $EventArg
    }

    Write-Verbose -Message "Exit Click Event for `$MultiTextBoxInputBtmMidButton"
  }
  #endregion ******** Function Start-MultiTextBoxInputBtmMidButtonClick ********
  $MultiTextBoxInputBtmMidButton.add_Click({ Start-MultiTextBoxInputBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $MultiTextBoxInputBtmRightButton = [System.Windows.Forms.Button]::New()
  $MultiTextBoxInputBtmRightButton = [System.Windows.Forms.Button]::New()
  $MultiTextBoxInputBtmPanel.Controls.Add($MultiTextBoxInputBtmRightButton)
  $MultiTextBoxInputBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $MultiTextBoxInputBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $MultiTextBoxInputBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $MultiTextBoxInputBtmRightButton.Font = [MyConfig]::Font.Bold
  $MultiTextBoxInputBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $MultiTextBoxInputBtmRightButton.Location = [System.Drawing.Point]::New(($MultiTextBoxInputBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $MultiTextBoxInputBtmRightButton.Name = "MultiTextBoxInputBtmRightButton"
  $MultiTextBoxInputBtmRightButton.TabIndex = 3
  $MultiTextBoxInputBtmRightButton.TabStop = $True
  $MultiTextBoxInputBtmRightButton.Text = $ButtonRight
  $MultiTextBoxInputBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $MultiTextBoxInputBtmRightButton.PreferredSize.Height)
  #endregion $MultiTextBoxInputBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-MultiTextBoxInputBtmRightButtonClick ********
  Function Start-MultiTextBoxInputBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the MultiTextBoxInputBtmRight Button Control
      .DESCRIPTION
        Click Event for the MultiTextBoxInputBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-MultiTextBoxInputBtmRightButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$MultiTextBoxInputBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here

    $MultiTextBoxInputForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$MultiTextBoxInputBtmRightButton"
  }
  #endregion ******** Function Start-MultiTextBoxInputBtmRightButtonClick ********
  $MultiTextBoxInputBtmRightButton.add_Click({ Start-MultiTextBoxInputBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $MultiTextBoxInputBtmPanel.ClientSize = [System.Drawing.Size]::New(($MultiTextBoxInputBtmRightButton.Right + [MyConfig]::FormSpacer), ($MultiTextBoxInputBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $MultiTextBoxInputBtmPanel Controls ********

  $MultiTextBoxInputForm.ClientSize = [System.Drawing.Size]::New($MultiTextBoxInputForm.ClientSize.Width, ($TempClientSize.Height + $MultiTextBoxInputBtmPanel.Height))

  #endregion ******** Controls for MultiTextBoxInput Form ********

  #endregion ******** End **** MultiTextBoxInput **** End ********

  $DialogResult = $MultiTextBoxInputForm.ShowDialog($PILForm)
  [MultiTextBoxInput]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, $OrderedItems)

  $MultiTextBoxInputForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-MultiTextBoxInput"
}
#endregion function Get-MultiTextBoxInput

# --------------------------------
# Get CheckedListBoxOprion Function
# --------------------------------
#region CheckedListBoxOprion Result Class
Class CheckedListBoxOprion
{
  [Bool]$Success
  [Object]$DialogResult
  [Object[]]$Items

  CheckedListBoxOprion ([Bool]$Success, [Object]$DialogResult, [Object[]]$Items)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.Items = $Items
  }
}
#endregion CheckedListBoxOprion Result Class

#region function Get-CheckedListBoxOprion
function Get-CheckedListBoxOprion ()
{
  <#
    .SYNOPSIS
      Shows Get-CheckedListBoxOprion
    .DESCRIPTION
      Shows Get-CheckedListBoxOprion
    .PARAMETER Title
      Title of the Get-CheckedListBoxOprion Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER Items
      Items to show in the ListView
    .PARAMETER Sorted
      Sort ListView
    .PARAMETER DisplayMember
      Name of the Property to Display in the CheckedListBox
    .PARAMETER ValueMember
      Name of the Property for the Value
    .PARAMETER Selected
      Default Selected CheckedListBox Items
    .PARAMETER Width
      Width of Get-CheckedListBoxOprion Dialog Window
    .PARAMETER Height
      Height of Get-CheckedListBoxOprion Dialog Window
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $Items = Get-Service
      $DialogResult = CheckedGet-CheckedListBoxOprion -Title "Get CheckListBox Option" -Message "Show this Sample Message Prompt to the User" -DisplayMember "DisplayName" -ValueMember "Name" -Items $Items -Selected $Items[1, 3, 5, 7]
      If ($DialogResult.Success)
      {
        # Success
      }
      Else
      {
        # Failed
      }
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [String]$Message = "Status Message",
    [parameter(Mandatory = $True)]
    [Object[]]$Items = @(),
    [Switch]$Sorted,
    [String]$DisplayMember = "Text",
    [String]$ValueMember = "Value",
    [Object[]]$Selected,
    [Int]$Width = 25,
    [Int]$Height = 20,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel",
    [Switch]$Required
  )
  Write-Verbose -Message "Enter Function Get-CheckedListBoxOprion"

  #region ******** Begin **** CheckedListBoxOprion **** Begin ********

  # ************************************************
  # CheckedListBoxOprion Form
  # ************************************************
  #region $CheckedListBoxOprionForm = [System.Windows.Forms.Form]::New()
  $CheckedListBoxOprionForm = [System.Windows.Forms.Form]::New()
  $CheckedListBoxOprionForm.BackColor = [MyConfig]::Colors.Back
  $CheckedListBoxOprionForm.Font = [MyConfig]::Font.Regular
  $CheckedListBoxOprionForm.ForeColor = [MyConfig]::Colors.Fore
  $CheckedListBoxOprionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $CheckedListBoxOprionForm.Icon = $PILForm.Icon
  $CheckedListBoxOprionForm.KeyPreview = $True
  $CheckedListBoxOprionForm.MaximizeBox = $False
  $CheckedListBoxOprionForm.MinimizeBox = $False
  $CheckedListBoxOprionForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * $Height))
  $CheckedListBoxOprionForm.Name = "CheckedListBoxOprionForm"
  $CheckedListBoxOprionForm.Owner = $PILForm
  $CheckedListBoxOprionForm.ShowInTaskbar = $False
  $CheckedListBoxOprionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $CheckedListBoxOprionForm.Text = $Title
  #endregion $CheckedListBoxOprionForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-CheckedListBoxOprionFormKeyDown ********
  function Start-CheckedListBoxOprionFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the CheckedListBoxOprion Form Control
      .DESCRIPTION
        KeyDown Event for the CheckedListBoxOprion Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-CheckedListBoxOprionFormKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$CheckedListBoxOprionForm"

    [MyConfig]::AutoExit = 0

    if ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $CheckedListBoxOprionForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$CheckedListBoxOprionForm"
  }
  #endregion ******** Function Start-CheckedListBoxOprionFormKeyDown ********
  $CheckedListBoxOprionForm.add_KeyDown({ Start-CheckedListBoxOprionFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-CheckedListBoxOprionFormShown ********
  function Start-CheckedListBoxOprionFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the CheckedListBoxOprion Form Control
      .DESCRIPTION
        Shown Event for the CheckedListBoxOprion Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-CheckedListBoxOprionFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$CheckedListBoxOprionForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    Write-Verbose -Message "Exit Shown Event for `$CheckedListBoxOprionForm"
  }
  #endregion ******** Function Start-CheckedListBoxOprionFormShown ********
  $CheckedListBoxOprionForm.add_Shown({ Start-CheckedListBoxOprionFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for CheckedListBoxOprion Form ********

  # ************************************************
  # CheckedListBoxOprion Panel
  # ************************************************
  #region $CheckedListBoxOprionPanel = [System.Windows.Forms.Panel]::New()
  $CheckedListBoxOprionPanel = [System.Windows.Forms.Panel]::New()
  $CheckedListBoxOprionForm.Controls.Add($CheckedListBoxOprionPanel)
  $CheckedListBoxOprionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $CheckedListBoxOprionPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $CheckedListBoxOprionPanel.Name = "CheckedListBoxOprionPanel"
  #endregion $CheckedListBoxOprionPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $CheckedListBoxOprionPanel Controls ********

  if ($PSBoundParameters.ContainsKey("Message"))
  {
    #region $CheckedListBoxOprionLabel = [System.Windows.Forms.Label]::New()
    $CheckedListBoxOprionLabel = [System.Windows.Forms.Label]::New()
    $CheckedListBoxOprionPanel.Controls.Add($CheckedListBoxOprionLabel)
    $CheckedListBoxOprionLabel.ForeColor = [MyConfig]::Colors.LabelFore
    $CheckedListBoxOprionLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
    $CheckedListBoxOprionLabel.Name = "CheckedListBoxOprionLabel"
    $CheckedListBoxOprionLabel.Size = [System.Drawing.Size]::New(($CheckedListBoxOprionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
    $CheckedListBoxOprionLabel.Text = $Message
    $CheckedListBoxOprionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #endregion $CheckedListBoxOprionLabel = [System.Windows.Forms.Label]::New()

    # Returns the minimum size required to display the text
    $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($CheckedListBoxOprionLabel.Text, [MyConfig]::Font.Regular, $CheckedListBoxOprionLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
    $CheckedListBoxOprionLabel.Size = [System.Drawing.Size]::New(($CheckedListBoxOprionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

    $TmpBottom = $CheckedListBoxOprionLabel.Bottom + [MyConfig]::FormSpacer
  }
  else
  {
    $TmpBottom = 0
  }

  # ************************************************
  # CheckedListBoxOprion GroupBox
  # ************************************************
  #region $CheckedListBoxOprionGroupBox = [System.Windows.Forms.GroupBox]::New()
  $CheckedListBoxOprionGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $CheckedListBoxOprionPanel.Controls.Add($CheckedListBoxOprionGroupBox)
  $CheckedListBoxOprionGroupBox.BackColor = [MyConfig]::Colors.Back
  $CheckedListBoxOprionGroupBox.Font = [MyConfig]::Font.Regular
  $CheckedListBoxOprionGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $CheckedListBoxOprionGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($TmpBottom + [MyConfig]::FormSpacer))
  $CheckedListBoxOprionGroupBox.Name = "CheckedListBoxOprionGroupBox"
  $CheckedListBoxOprionGroupBox.Size = [System.Drawing.Size]::New(($CheckedListBoxOprionPanel.Width - ([MyConfig]::FormSpacer * 2)), ($CheckedListBoxOprionPanel.Height - ($CheckedListBoxOprionGroupBox.Top + [MyConfig]::FormSpacer)))
  #endregion $CheckedListBoxOprionGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $CheckedListBoxOprionGroupBox Controls ********

  #region $CheckedListBoxOprionCheckedListBox = [System.Windows.Forms.CheckedListBox]::New()
  $CheckedListBoxOprionCheckedListBox = [System.Windows.Forms.CheckedListBox]::New()
  $CheckedListBoxOprionGroupBox.Controls.Add($CheckedListBoxOprionCheckedListBox)
  $CheckedListBoxOprionCheckedListBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
  $CheckedListBoxOprionCheckedListBox.AutoSize = $True
  $CheckedListBoxOprionCheckedListBox.BackColor = [MyConfig]::Colors.TextBack
  $CheckedListBoxOprionCheckedListBox.CheckOnClick = $True
  $CheckedListBoxOprionCheckedListBox.DisplayMember = $DisplayMember
  $CheckedListBoxOprionCheckedListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $CheckedListBoxOprionCheckedListBox.Font = [MyConfig]::Font.Regular
  $CheckedListBoxOprionCheckedListBox.ForeColor = [MyConfig]::Colors.TextFore
  $CheckedListBoxOprionCheckedListBox.Name = "CheckedListBoxOprionCheckedListBox"
  $CheckedListBoxOprionCheckedListBox.Sorted = $Sorted.IsPresent
  $CheckedListBoxOprionCheckedListBox.TabIndex = 0
  $CheckedListBoxOprionCheckedListBox.TabStop = $True
  $CheckedListBoxOprionCheckedListBox.Tag = $Null
  $CheckedListBoxOprionCheckedListBox.ValueMember = $ValueMember
  #endregion $CheckedListBoxOprionCheckedListBox = [System.Windows.Forms.CheckedListBox]::New()

  $CheckedListBoxOprionCheckedListBox.Items.AddRange($Items)

  if ($PSBoundParameters.ContainsKey("Selected"))
  {
    $CheckedListBoxOprionCheckedListBox.Tag = @($Items | Where-Object -FilterScript { $PSItem -in $Selected})
    $CheckedListBoxOprionCheckedListBox.Tag | ForEach-Object -Process { $CheckedListBoxOprionCheckedListBox.SetItemChecked($CheckedListBoxOprionCheckedListBox.Items.IndexOf($PSItem), $True) }
  }
  else
  {
    $CheckedListBoxOprionCheckedListBox.Tag = @()
  }

  #region ******** Function Start-CheckedListBoxOprionCheckedListBoxMouseDown ********
  function Start-CheckedListBoxOprionCheckedListBoxMouseDown
  {
    <#
      .SYNOPSIS
        MouseDown Event for the IDP TreeView Control
      .DESCRIPTION
        MouseDown Event for the IDP TreeView Control
      .PARAMETER Sender
         The TreeView Control that fired the MouseDown Event
      .PARAMETER EventArg
         The Event Arguments for the TreeView MouseDown Event
      .EXAMPLE
         Start-CheckedListBoxOprionCheckedListBoxMouseDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.CheckedListBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter MouseDown Event for `$CheckedListBoxOprionCheckedListBox"

    [MyConfig]::AutoExit = 0

    If ($EventArg.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
      if ($CheckedListBoxOprionCheckedListBox.Items.Count -gt 0)
      {
        $CheckedListBoxOprionContextMenuStrip.Show($CheckedListBoxOprionCheckedListBox, $EventArg.Location)
      }
    }

    Write-Verbose -Message "Exit MouseDown Event for `$CheckedListBoxOprionCheckedListBox"
  }
  #endregion ******** Function Start-CheckedListBoxOprionCheckedListBoxMouseDown ********
  $CheckedListBoxOprionCheckedListBox.add_MouseDown({ Start-CheckedListBoxOprionCheckedListBoxMouseDown -Sender $This -EventArg $PSItem })

  $CheckedListBoxOprionGroupBox.ClientSize = [System.Drawing.Size]::New($CheckedListBoxOprionGroupBox.ClientSize.Width, ($CheckedListBoxOprionCheckedListBox.Bottom + ([MyConfig]::FormSpacer * 2)))

  # ************************************************
  # CheckedListBoxOprion ContextMenuStrip
  # ************************************************
  #region $CheckedListBoxOprionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  $CheckedListBoxOprionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  #$CheckedListBoxOprionListView.Controls.Add($CheckedListBoxOprionContextMenuStrip)
  $CheckedListBoxOprionContextMenuStrip.BackColor = [MyConfig]::Colors.Back
  #$CheckedListBoxOprionContextMenuStrip.Enabled = $True
  $CheckedListBoxOprionContextMenuStrip.Font = [MyConfig]::Font.Regular
  $CheckedListBoxOprionContextMenuStrip.ForeColor = [MyConfig]::Colors.Fore
  $CheckedListBoxOprionContextMenuStrip.ImageList = $PILSmallImageList
  $CheckedListBoxOprionContextMenuStrip.Name = "CheckedListBoxOprionContextMenuStrip"
  #endregion $CheckedListBoxOprionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()

  #region ******** Function Start-CheckedListBoxOprionContextMenuStripOpening ********
  function Start-CheckedListBoxOprionContextMenuStripOpening
  {
    <#
      .SYNOPSIS
        Opening Event for the CheckedListBoxOprion ContextMenuStrip Control
      .DESCRIPTION
        Opening Event for the CheckedListBoxOprion ContextMenuStrip Control
      .PARAMETER Sender
         The ContextMenuStrip Control that fired the Opening Event
      .PARAMETER EventArg
         The Event Arguments for the ContextMenuStrip Opening Event
      .EXAMPLE
         Start-CheckedListBoxOprionContextMenuStripOpening -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ContextMenuStrip]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Opening Event for `$CheckedListBoxOprionContextMenuStrip"

    [MyConfig]::AutoExit = 0

    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

    Write-Verbose -Message "Exit Opening Event for `$CheckedListBoxOprionContextMenuStrip"
  }
  #endregion ******** Function Start-CheckedListBoxOprionContextMenuStripOpening ********
  $CheckedListBoxOprionContextMenuStrip.add_Opening({Start-CheckedListBoxOprionContextMenuStripOpening -Sender $This -EventArg $PSItem})

  #region ******** Function Start-CheckedListBoxOprionContextMenuStripItemClick ********
  function Start-CheckedListBoxOprionContextMenuStripItemClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckedListBoxOprion ToolStripItem Control
      .DESCRIPTION
        Click Event for the CheckedListBoxOprion ToolStripItem Control
      .PARAMETER Sender
         The ToolStripItem Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the ToolStripItem Click Event
      .EXAMPLE
         Start-CheckedListBoxOprionContextMenuStripItemClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ToolStripItem]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$CheckedListBoxOprionContextMenuStripItem"

    [MyConfig]::AutoExit = 0

    switch ($Sender.Name)
    {
      "CheckAll"
      {
        $TmpCheckedItems = @($CheckedListBoxOprionCheckedListBox.CheckedIndices)
        (0..$($CheckedListBoxOprionCheckedListBox.Items.Count - 1)) | Where-Object -FilterScript { $PSItem -notin $TmpCheckedItems } | ForEach-Object -Process { $CheckedListBoxOprionCheckedListBox.SetItemChecked($PSItem, $True) }
        Break
      }
      "UnCheckAll"
      {
        $TmpCheckedItems = @($CheckedListBoxOprionCheckedListBox.CheckedIndices)
        $TmpCheckedItems | ForEach-Object -Process { $CheckedListBoxOprionCheckedListBox.SetItemChecked($PSItem, $False) }
        Break
      }
    }

    Write-Verbose -Message "Exit Click Event for `$CheckedListBoxOprionContextMenuStripItem"
  }
  #endregion ******** Function Start-CheckedListBoxOprionContextMenuStripItemClick ********

  (New-MenuItem -Menu $CheckedListBoxOprionContextMenuStrip -Text "Check All" -Name "CheckAll" -Tag "CheckAll" -DisplayStyle "ImageAndText" -ImageKey "CheckIcon" -PassThru).add_Click({Start-CheckedListBoxOprionContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $CheckedListBoxOprionContextMenuStrip -Text "Uncheck All" -Name "UnCheckAll" -Tag "UnCheckAll" -DisplayStyle "ImageAndText" -ImageKey "UncheckIcon" -PassThru).add_Click({Start-CheckedListBoxOprionContextMenuStripItemClick -Sender $This -EventArg $PSItem})

  #endregion ******** $CheckedListBoxOprionGroupBox Controls ********

  $TempClientSize = [System.Drawing.Size]::New(($CheckedListBoxOprionGroupBox.Right + [MyConfig]::FormSpacer), ($CheckedListBoxOprionGroupBox.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $CheckedListBoxOprionPanel Controls ********

  # ************************************************
  # CheckedListBoxOprionBtm Panel
  # ************************************************
  #region $CheckedListBoxOprionBtmPanel = [System.Windows.Forms.Panel]::New()
  $CheckedListBoxOprionBtmPanel = [System.Windows.Forms.Panel]::New()
  $CheckedListBoxOprionForm.Controls.Add($CheckedListBoxOprionBtmPanel)
  $CheckedListBoxOprionBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $CheckedListBoxOprionBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $CheckedListBoxOprionBtmPanel.Name = "CheckedListBoxOprionBtmPanel"
  #endregion $CheckedListBoxOprionBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $CheckedListBoxOprionBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($CheckedListBoxOprionBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $CheckedListBoxOprionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOprionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOprionBtmPanel.Controls.Add($CheckedListBoxOprionBtmLeftButton)
  $CheckedListBoxOprionBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $CheckedListBoxOprionBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $CheckedListBoxOprionBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $CheckedListBoxOprionBtmLeftButton.Font = [MyConfig]::Font.Bold
  $CheckedListBoxOprionBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $CheckedListBoxOprionBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $CheckedListBoxOprionBtmLeftButton.Name = "CheckedListBoxOprionBtmLeftButton"
  $CheckedListBoxOprionBtmLeftButton.TabIndex = 1
  $CheckedListBoxOprionBtmLeftButton.TabStop = $True
  $CheckedListBoxOprionBtmLeftButton.Text = $ButtonLeft
  $CheckedListBoxOprionBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $CheckedListBoxOprionBtmLeftButton.PreferredSize.Height)
  #endregion $CheckedListBoxOprionBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-CheckedListBoxOprionBtmLeftButtonClick ********
  function Start-CheckedListBoxOprionBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckedListBoxOprionBtmLeft Button Control
      .DESCRIPTION
        Click Event for the CheckedListBoxOprionBtmLeft Button Control
      .PARAMETER Sender
         The Button Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the Button Click Event
      .EXAMPLE
         Start-CheckedListBoxOprionBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$CheckedListBoxOprionBtmLeftButton"

    [MyConfig]::AutoExit = 0

    if ($CheckedListBoxOprionCheckedListBox.CheckedItems.Count -gt 0)
    {
      $CheckedListBoxOprionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    else
    {
      [Void][System.Windows.Forms.MessageBox]::Show($CheckedListBoxOprionForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }

    Write-Verbose -Message "Exit Click Event for `$CheckedListBoxOprionBtmLeftButton"
  }
  #endregion ******** Function Start-CheckedListBoxOprionBtmLeftButtonClick ********
  $CheckedListBoxOprionBtmLeftButton.add_Click({ Start-CheckedListBoxOprionBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $CheckedListBoxOprionBtmMidButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOprionBtmMidButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOprionBtmPanel.Controls.Add($CheckedListBoxOprionBtmMidButton)
  $CheckedListBoxOprionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $CheckedListBoxOprionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $CheckedListBoxOprionBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $CheckedListBoxOprionBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $CheckedListBoxOprionBtmMidButton.Font = [MyConfig]::Font.Bold
  $CheckedListBoxOprionBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $CheckedListBoxOprionBtmMidButton.Location = [System.Drawing.Point]::New(($CheckedListBoxOprionBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $CheckedListBoxOprionBtmMidButton.Name = "CheckedListBoxOprionBtmMidButton"
  $CheckedListBoxOprionBtmMidButton.TabIndex = 2
  $CheckedListBoxOprionBtmMidButton.TabStop = $True
  $CheckedListBoxOprionBtmMidButton.Text = $ButtonMid
  $CheckedListBoxOprionBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $CheckedListBoxOprionBtmMidButton.PreferredSize.Height)
  #endregion $CheckedListBoxOprionBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-CheckedListBoxOprionBtmMidButtonClick ********
  function Start-CheckedListBoxOprionBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckedListBoxOprionBtmMid Button Control
      .DESCRIPTION
        Click Event for the CheckedListBoxOprionBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-CheckedListBoxOprionBtmMidButtonClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By MyUserName)
  #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$CheckedListBoxOprionBtmMidButton"

    [MyConfig]::AutoExit = 0

    $TmpCheckedItems = @($CheckedListBoxOprionCheckedListBox.CheckedIndices)
    $TmpCheckedItems | ForEach-Object -Process { $CheckedListBoxOprionCheckedListBox.SetItemChecked($PSItem, $False) }
    if ($CheckedListBoxOprionCheckedListBox.Tag.Count -gt 0)
    {
      $CheckedListBoxOprionCheckedListBox.Tag | ForEach-Object -Process { $CheckedListBoxOprionCheckedListBox.SetItemChecked($CheckedListBoxOprionCheckedListBox.Items.IndexOf($PSItem), $True) }
    }

    Write-Verbose -Message "Exit Click Event for `$CheckedListBoxOprionBtmMidButton"
  }
  #endregion ******** Function Start-CheckedListBoxOprionBtmMidButtonClick ********
  $CheckedListBoxOprionBtmMidButton.add_Click({ Start-CheckedListBoxOprionBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $CheckedListBoxOprionBtmRightButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOprionBtmRightButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOprionBtmPanel.Controls.Add($CheckedListBoxOprionBtmRightButton)
  $CheckedListBoxOprionBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $CheckedListBoxOprionBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $CheckedListBoxOprionBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $CheckedListBoxOprionBtmRightButton.Font = [MyConfig]::Font.Bold
  $CheckedListBoxOprionBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $CheckedListBoxOprionBtmRightButton.Location = [System.Drawing.Point]::New(($CheckedListBoxOprionBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $CheckedListBoxOprionBtmRightButton.Name = "CheckedListBoxOprionBtmRightButton"
  $CheckedListBoxOprionBtmRightButton.TabIndex = 3
  $CheckedListBoxOprionBtmRightButton.TabStop = $True
  $CheckedListBoxOprionBtmRightButton.Text = $ButtonRight
  $CheckedListBoxOprionBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $CheckedListBoxOprionBtmRightButton.PreferredSize.Height)
  #endregion $CheckedListBoxOprionBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-CheckedListBoxOprionBtmRightButtonClick ********
  function Start-CheckedListBoxOprionBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckedListBoxOprionBtmRight Button Control
      .DESCRIPTION
        Click Event for the CheckedListBoxOprionBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-CheckedListBoxOprionBtmRightButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$CheckedListBoxOprionBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here

    $CheckedListBoxOprionForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$CheckedListBoxOprionBtmRightButton"
  }
  #endregion ******** Function Start-CheckedListBoxOprionBtmRightButtonClick ********
  $CheckedListBoxOprionBtmRightButton.add_Click({ Start-CheckedListBoxOprionBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $CheckedListBoxOprionBtmPanel.ClientSize = [System.Drawing.Size]::New(($CheckedListBoxOprionBtmRightButton.Right + [MyConfig]::FormSpacer), ($CheckedListBoxOprionBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $CheckedListBoxOprionBtmPanel Controls ********

  $CheckedListBoxOprionForm.ClientSize = [System.Drawing.Size]::New($CheckedListBoxOprionForm.ClientSize.Width, ($TempClientSize.Height + $CheckedListBoxOprionBtmPanel.Height))

  #endregion ******** Controls for CheckedListBoxOprion Form ********

  #endregion ******** End **** CheckedListBoxOprion **** End ********

  $DialogResult = $CheckedListBoxOprionForm.ShowDialog()
  if ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
  {
    [CheckedListBoxOprion]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, $CheckedListBoxOprionCheckedListBox.CheckedItems)
  }
  else
  {
    [CheckedListBoxOprion]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, @())
  }

  $CheckedListBoxOprionForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-CheckedListBoxOprion"
}
#endregion function Get-CheckedListBoxOprion

# --------------------------------
# Get ComboBoxOption Function
# --------------------------------
#region ComboBoxOption Result Class
Class ComboBoxOption
{
  [Bool]$Success
  [Object]$DialogResult
  [Object]$Item

  ComboBoxOption ([Bool]$Success, [Object]$DialogResult, [Object]$Item)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.Item = $Item
  }
}
#endregion ComboBoxOption Result Class

#region function Get-ComboBoxOption
function Get-ComboBoxOption ()
{
  <#
    .SYNOPSIS
      Shows Get-ComboBoxOption
    .DESCRIPTION
      Shows Get-ComboBoxOption
    .PARAMETER Title
      Title of the Get-ComboBoxOption Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER Items
      Items to show in the ComboBox
    .PARAMETER Sorted
      Sort ComboBox
    .PARAMETER SelectText
      The Default Selected Item when no Value is Selected
    .PARAMETER DisplayMember
      Name of the Property to Display in the CheckedListBox
    .PARAMETER ValueMember
      Name of the Property for the Value
    .PARAMETER Selected
      Default Selected ComboBox Item
    .PARAMETER Width
      Width of Get-ComboBoxOption Dialog Window
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $Variables = @(Get-ChildItem -Path "Variable:\")
      $DialogResult = Get-ComboBoxOption -Title "Combo Choice Dialog 01" -Message "Show this Sample Message Prompt to the User" -Items $Variables -DisplayMember "Name" -ValueMember "Value" -Selected ($Variables[4])
      If ($DialogResult.Success)
      {
        # Success
      }
      Else
      {
        # Failed
      }
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [String]$Message = "Status Message",
    [parameter(Mandatory = $True)]
    [Object[]]$Items = @(),
    [Switch]$Sorted,
    [String]$SelectText = "Select Value",
    [String]$DisplayMember = "Text",
    [String]$ValueMember = "Value",
    [Object]$Selected,
    [Int]$Width = 35,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel"
  )
  Write-Verbose -Message "Enter Function Get-ComboBoxOption"

  #region ******** Begin **** ComboBoxOption **** Begin ********

  # ************************************************
  # ComboBoxOption Form
  # ************************************************
  #region $ComboBoxOptionForm = [System.Windows.Forms.Form]::New()
  $ComboBoxOptionForm = [System.Windows.Forms.Form]::New()
  $ComboBoxOptionForm.BackColor = [MyConfig]::Colors.Back
  $ComboBoxOptionForm.Font = [MyConfig]::Font.Regular
  $ComboBoxOptionForm.ForeColor = [MyConfig]::Colors.Fore
  $ComboBoxOptionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $ComboBoxOptionForm.Icon = $PILForm.Icon
  $ComboBoxOptionForm.KeyPreview = $True
  $ComboBoxOptionForm.MaximizeBox = $False
  $ComboBoxOptionForm.MinimizeBox = $False
  $ComboBoxOptionForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), 0)
  $ComboBoxOptionForm.Name = "ComboBoxOptionForm"
  $ComboBoxOptionForm.Owner = $PILForm
  $ComboBoxOptionForm.ShowInTaskbar = $False
  $ComboBoxOptionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $ComboBoxOptionForm.Text = $Title
  #endregion $ComboBoxOptionForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-ComboBoxOptionFormKeyDown ********
  function Start-ComboBoxOptionFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the ComboBoxOption Form Control
      .DESCRIPTION
        KeyDown Event for the ComboBoxOption Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-ComboBoxOptionFormKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$ComboBoxOptionForm"

    [MyConfig]::AutoExit = 0

    if ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $ComboBoxOptionForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$ComboBoxOptionForm"
  }
  #endregion ******** Function Start-ComboBoxOptionFormKeyDown ********
  $ComboBoxOptionForm.add_KeyDown({ Start-ComboBoxOptionFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ComboBoxOptionFormShown ********
  function Start-ComboBoxOptionFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the ComboBoxOption Form Control
      .DESCRIPTION
        Shown Event for the ComboBoxOption Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-ComboBoxOptionFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$ComboBoxOptionForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Shown Event for `$ComboBoxOptionForm"
  }
  #endregion ******** Function Start-ComboBoxOptionFormShown ********
  $ComboBoxOptionForm.add_Shown({ Start-ComboBoxOptionFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for ComboBoxOption Form ********

  # ************************************************
  # ComboBoxOption Panel
  # ************************************************
  #region $ComboBoxOptionPanel = [System.Windows.Forms.Panel]::New()
  $ComboBoxOptionPanel = [System.Windows.Forms.Panel]::New()
  $ComboBoxOptionForm.Controls.Add($ComboBoxOptionPanel)
  $ComboBoxOptionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ComboBoxOptionPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $ComboBoxOptionPanel.Name = "ComboBoxOptionPanel"
  #endregion $ComboBoxOptionPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ComboBoxOptionPanel Controls ********

  if ($PSBoundParameters.ContainsKey("Message"))
  {
    #region $ComboBoxOptionLabel = [System.Windows.Forms.Label]::New()
    $ComboBoxOptionLabel = [System.Windows.Forms.Label]::New()
    $ComboBoxOptionPanel.Controls.Add($ComboBoxOptionLabel)
    $ComboBoxOptionLabel.ForeColor = [MyConfig]::Colors.LabelFore
    $ComboBoxOptionLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
    $ComboBoxOptionLabel.Name = "ComboBoxOptionLabel"
    $ComboBoxOptionLabel.Size = [System.Drawing.Size]::New(($ComboBoxOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
    $ComboBoxOptionLabel.Text = $Message
    $ComboBoxOptionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #endregion $ComboBoxOptionLabel = [System.Windows.Forms.Label]::New()

    # Returns the minimum size required to display the text
    $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($ComboBoxOptionLabel.Text, [MyConfig]::Font.Regular, $ComboBoxOptionLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
    $ComboBoxOptionLabel.Size = [System.Drawing.Size]::New(($ComboBoxOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

    $TmpBottom = $ComboBoxOptionLabel.Bottom + [MyConfig]::FormSpacer
  }
  else
  {
    $TmpBottom = 0
  }

  # ************************************************
  # ComboBoxOption GroupBox
  # ************************************************
  #region $ComboBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  $ComboBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $ComboBoxOptionPanel.Controls.Add($ComboBoxOptionGroupBox)
  $ComboBoxOptionGroupBox.BackColor = [MyConfig]::Colors.Back
  $ComboBoxOptionGroupBox.Font = [MyConfig]::Font.Regular
  $ComboBoxOptionGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $ComboBoxOptionGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($TmpBottom + [MyConfig]::FormSpacer))
  $ComboBoxOptionGroupBox.Name = "ComboBoxOptionGroupBox"
  $ComboBoxOptionGroupBox.Size = [System.Drawing.Size]::New(($ComboBoxOptionPanel.Width - ([MyConfig]::FormSpacer * 2)), 50)
  #endregion $ComboBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $ComboBoxOptionGroupBox Controls ********

  #region $GetComboChoiceComboBox = [System.Windows.Forms.ComboBox]::New()
  $GetComboChoiceComboBox = [System.Windows.Forms.ComboBox]::New()
  $ComboBoxOptionGroupBox.Controls.Add($GetComboChoiceComboBox)
  $GetComboChoiceComboBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
  $GetComboChoiceComboBox.AutoSize = $True
  $GetComboChoiceComboBox.BackColor = [MyConfig]::Colors.TextBack
  $GetComboChoiceComboBox.DisplayMember = $DisplayMember
  $GetComboChoiceComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  $GetComboChoiceComboBox.Font = [MyConfig]::Font.Regular
  $GetComboChoiceComboBox.ForeColor = [MyConfig]::Colors.TextFore
  [void]$GetComboChoiceComboBox.Items.Add([PSCustomObject]@{ $DisplayMember = " - $($SelectText) - "; $ValueMember = " - $($SelectText) - "})
  $GetComboChoiceComboBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $GetComboChoiceComboBox.Name = "GetComboChoiceComboBox"
  $GetComboChoiceComboBox.SelectedIndex = 0
  $GetComboChoiceComboBox.Size = [System.Drawing.Size]::New(($ComboBoxOptionGroupBox.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), $GetComboChoiceComboBox.PreferredHeight)
  $GetComboChoiceComboBox.Sorted = $Sorted.IsPresent
  $GetComboChoiceComboBox.TabIndex = 0
  $GetComboChoiceComboBox.TabStop = $True
  $GetComboChoiceComboBox.Tag = $Null
  $GetComboChoiceComboBox.ValueMember = $ValueMember
  #endregion $GetComboChoiceComboBox = [System.Windows.Forms.ComboBox]::New()

  $GetComboChoiceComboBox.Items.AddRange($Items)

  if ($PSBoundParameters.ContainsKey("Selected"))
  {
    $GetComboChoiceComboBox.Tag = $Items | Where-Object -FilterScript { $PSItem -eq $Selected}
    $GetComboChoiceComboBox.SelectedItem = $GetComboChoiceComboBox.Tag
  }
  else
  {
    $GetComboChoiceComboBox.SelectedIndex = 0
  }

  $ComboBoxOptionGroupBox.ClientSize = [System.Drawing.Size]::New($ComboBoxOptionGroupBox.ClientSize.Width, ($GetComboChoiceComboBox.Bottom + ([MyConfig]::FormSpacer * 2)))

  #endregion ******** $ComboBoxOptionGroupBox Controls ********

  $TempClientSize = [System.Drawing.Size]::New(($ComboBoxOptionGroupBox.Right + [MyConfig]::FormSpacer), ($ComboBoxOptionGroupBox.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ComboBoxOptionPanel Controls ********

  # ************************************************
  # ComboBoxOptionBtm Panel
  # ************************************************
  #region $ComboBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $ComboBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $ComboBoxOptionForm.Controls.Add($ComboBoxOptionBtmPanel)
  $ComboBoxOptionBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ComboBoxOptionBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $ComboBoxOptionBtmPanel.Name = "ComboBoxOptionBtmPanel"
  #endregion $ComboBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ComboBoxOptionBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($ComboBoxOptionBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $ComboBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ComboBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ComboBoxOptionBtmPanel.Controls.Add($ComboBoxOptionBtmLeftButton)
  $ComboBoxOptionBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $ComboBoxOptionBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ComboBoxOptionBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ComboBoxOptionBtmLeftButton.Font = [MyConfig]::Font.Bold
  $ComboBoxOptionBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ComboBoxOptionBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $ComboBoxOptionBtmLeftButton.Name = "ComboBoxOptionBtmLeftButton"
  $ComboBoxOptionBtmLeftButton.TabIndex = 1
  $ComboBoxOptionBtmLeftButton.TabStop = $True
  $ComboBoxOptionBtmLeftButton.Text = $ButtonLeft
  $ComboBoxOptionBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $ComboBoxOptionBtmLeftButton.PreferredSize.Height)
  #endregion $ComboBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ComboBoxOptionBtmLeftButtonClick ********
  function Start-ComboBoxOptionBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ComboBoxOptionBtmLeft Button Control
      .DESCRIPTION
        Click Event for the ComboBoxOptionBtmLeft Button Control
      .PARAMETER Sender
         The Button Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the Button Click Event
      .EXAMPLE
         Start-ComboBoxOptionBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ComboBoxOptionBtmLeftButton"

    [MyConfig]::AutoExit = 0

    if ($GetComboChoiceComboBox.SelectedIndex -gt 0)
    {
      $ComboBoxOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    else
    {
      [Void][System.Windows.Forms.MessageBox]::Show($ComboBoxOptionForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }

    Write-Verbose -Message "Exit Click Event for `$ComboBoxOptionBtmLeftButton"
  }
  #endregion ******** Function Start-ComboBoxOptionBtmLeftButtonClick ********
  $ComboBoxOptionBtmLeftButton.add_Click({ Start-ComboBoxOptionBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $ComboBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $ComboBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $ComboBoxOptionBtmPanel.Controls.Add($ComboBoxOptionBtmMidButton)
  $ComboBoxOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $ComboBoxOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $ComboBoxOptionBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ComboBoxOptionBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ComboBoxOptionBtmMidButton.Font = [MyConfig]::Font.Bold
  $ComboBoxOptionBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ComboBoxOptionBtmMidButton.Location = [System.Drawing.Point]::New(($ComboBoxOptionBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ComboBoxOptionBtmMidButton.Name = "ComboBoxOptionBtmMidButton"
  $ComboBoxOptionBtmMidButton.TabIndex = 2
  $ComboBoxOptionBtmMidButton.TabStop = $True
  $ComboBoxOptionBtmMidButton.Text = $ButtonMid
  $ComboBoxOptionBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $ComboBoxOptionBtmMidButton.PreferredSize.Height)
  #endregion $ComboBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ComboBoxOptionBtmMidButtonClick ********
  function Start-ComboBoxOptionBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ComboBoxOptionBtmMid Button Control
      .DESCRIPTION
        Click Event for the ComboBoxOptionBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ComboBoxOptionBtmMidButtonClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By MyUserName)
  #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ComboBoxOptionBtmMidButton"

    [MyConfig]::AutoExit = 0

    if ([String]::IsNullOrEmpty($GetComboChoiceComboBox.Tag))
    {
      $GetComboChoiceComboBox.SelectedIndex = 0
    }
    else
    {
      $GetComboChoiceComboBox.SelectedItem = $GetComboChoiceComboBox.Tag
    }

    Write-Verbose -Message "Exit Click Event for `$ComboBoxOptionBtmMidButton"
  }
  #endregion ******** Function Start-ComboBoxOptionBtmMidButtonClick ********
  $ComboBoxOptionBtmMidButton.add_Click({ Start-ComboBoxOptionBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $ComboBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $ComboBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $ComboBoxOptionBtmPanel.Controls.Add($ComboBoxOptionBtmRightButton)
  $ComboBoxOptionBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $ComboBoxOptionBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ComboBoxOptionBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ComboBoxOptionBtmRightButton.Font = [MyConfig]::Font.Bold
  $ComboBoxOptionBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ComboBoxOptionBtmRightButton.Location = [System.Drawing.Point]::New(($ComboBoxOptionBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ComboBoxOptionBtmRightButton.Name = "ComboBoxOptionBtmRightButton"
  $ComboBoxOptionBtmRightButton.TabIndex = 3
  $ComboBoxOptionBtmRightButton.TabStop = $True
  $ComboBoxOptionBtmRightButton.Text = $ButtonRight
  $ComboBoxOptionBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $ComboBoxOptionBtmRightButton.PreferredSize.Height)
  #endregion $ComboBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ComboBoxOptionBtmRightButtonClick ********
  function Start-ComboBoxOptionBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ComboBoxOptionBtmRight Button Control
      .DESCRIPTION
        Click Event for the ComboBoxOptionBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ComboBoxOptionBtmRightButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ComboBoxOptionBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here

    $ComboBoxOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$ComboBoxOptionBtmRightButton"
  }
  #endregion ******** Function Start-ComboBoxOptionBtmRightButtonClick ********
  $ComboBoxOptionBtmRightButton.add_Click({ Start-ComboBoxOptionBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $ComboBoxOptionBtmPanel.ClientSize = [System.Drawing.Size]::New(($ComboBoxOptionBtmRightButton.Right + [MyConfig]::FormSpacer), ($ComboBoxOptionBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ComboBoxOptionBtmPanel Controls ********

  $ComboBoxOptionForm.ClientSize = [System.Drawing.Size]::New($ComboBoxOptionForm.ClientSize.Width, ($TempClientSize.Height + $ComboBoxOptionBtmPanel.Height))

  #endregion ******** Controls for ComboBoxOption Form ********

  #endregion ******** End **** ComboBoxOption **** End ********

  $DialogResult = $ComboBoxOptionForm.ShowDialog()
  [ComboBoxOption]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, $GetComboChoiceComboBox.SelectedItem)

  $ComboBoxOptionForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-ComboBoxOption"
}
#endregion function Get-ComboBoxOption

# --------------------------------
# Get ListViewOption Function
# --------------------------------
#region ListViewOption Result Class
Class ListViewOption
{
  [Bool]$Success
  [Object]$DialogResult
  [Object]$Item

  ListViewOption ([Bool]$Success, [Object]$DialogResult, [Object]$Item)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.Item = $Item
  }
}
#endregion ListViewOption Result Class

#region function Get-ListViewOption
function Get-ListViewOption ()
{
  <#
    .SYNOPSIS
      Shows Get-ListViewOption
    .DESCRIPTION
      Shows Get-ListViewOption
    .PARAMETER Title
      Title of the Get-ListViewOption Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER Items
      Items to show in the ListVieww
    .PARAMETER Property
      Name of the Properties to Display
    .PARAMETER Tooltip
      ToolTip to Displays
    .PARAMETER SelectText
      Selected Text
    .PARAMETER Selected
      Selected ListView Items
    .PARAMETER Multi
      Allow Select Multiple Rows
    .PARAMETER Width
      Width of Get-ListViewOption Dialog Window
    .PARAMETER Height
      Height of Get-ListViewOption Dialog Window
    .PARAMETER Filter
      Show Filter TextBox
    .PARAMETER Resize
      Make Get-ListViewOption Dialog Window ReSixeable
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $Functions = @(Get-ChildItem -Path "Function:\")
      $DialogResult = Get-ListViewOption -Title "ListView Choice Dialog 01" -Message "Show this Sample Message Prompt to the User" -Items $Functions -Property "Name", "Version", "Source" -Selected ($Functions[2]) -Tooltip "Show this ToolTip" -Resize -Multi
      If ($DialogResult.Success)
      {
        # Success
      }
      Else
      {
        # Failed
      }
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [String]$Message = "Status Message",
    [parameter(Mandatory = $True)]
    [Object[]]$Items = @(),
    [parameter(Mandatory = $True)]
    [String[]]$Property,
    [String]$Tooltip,
    [Object[]]$Selected = "xX NONE Xx",
    [Switch]$Multi,
    [Switch]$AllowSort,
    [Int]$Width = 50,
    [Int]$Height = 12,
    [Switch]$Filter,
    [Switch]$Resize,
    [Switch]$Required,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel"
  )
  Write-Verbose -Message "Enter Function Get-ListViewOption"

  #region ******** Begin **** ListViewOption **** Begin ********

  # ************************************************
  # ListViewOption Form
  # ************************************************
  #region $ListViewOptionForm = [System.Windows.Forms.Form]::New()
  $ListViewOptionForm = [System.Windows.Forms.Form]::New()
  $ListViewOptionForm.BackColor = [MyConfig]::Colors.Back
  $ListViewOptionForm.Font = [MyConfig]::Font.Regular
  $ListViewOptionForm.ForeColor = [MyConfig]::Colors.Fore
  if ($Resize.IsPresent)
  {
    $ListViewOptionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
  }
  else
  {
    $ListViewOptionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  }
  $ListViewOptionForm.Icon = $PILForm.Icon
  $ListViewOptionForm.KeyPreview = $True
  $ListViewOptionForm.MaximizeBox = $False
  $ListViewOptionForm.MinimizeBox = $False
  $ListViewOptionForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), 0)
  $ListViewOptionForm.Name = "ListViewOptionForm"
  $ListViewOptionForm.Owner = $PILForm
  $ListViewOptionForm.ShowInTaskbar = $False
  $ListViewOptionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $ListViewOptionForm.Text = $Title
  #endregion $ListViewOptionForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-ListViewOptionFormKeyDown ********
  function Start-ListViewOptionFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the ListViewOption Form Control
      .DESCRIPTION
        KeyDown Event for the ListViewOption Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-ListViewOptionFormKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$ListViewOptionForm"

    [MyConfig]::AutoExit = 0

    if ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $ListViewOptionForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$ListViewOptionForm"
  }
  #endregion ******** Function Start-ListViewOptionFormKeyDown ********
  $ListViewOptionForm.add_KeyDown({ Start-ListViewOptionFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ListViewOptionFormShown ********
  function Start-ListViewOptionFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the ListViewOption Form Control
      .DESCRIPTION
        Shown Event for the ListViewOption Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-ListViewOptionFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$ListViewOptionForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Shown Event for `$ListViewOptionForm"
  }
  #endregion ******** Function Start-ListViewOptionFormShown ********
  $ListViewOptionForm.add_Shown({ Start-ListViewOptionFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for ListViewOption Form ********

  # ************************************************
  # ListViewOption Panel
  # ************************************************
  #region $ListViewOptionPanel = [System.Windows.Forms.Panel]::New()
  $ListViewOptionPanel = [System.Windows.Forms.Panel]::New()
  $ListViewOptionForm.Controls.Add($ListViewOptionPanel)
  $ListViewOptionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ListViewOptionPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $ListViewOptionPanel.Name = "ListViewOptionPanel"
  $ListViewOptionPanel.Text = "ListViewOptionPanel"
  #endregion $ListViewOptionPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ListViewOptionPanel Controls ********

  if ($PSBoundParameters.ContainsKey("Message"))
  {
    #region $ListViewOptionLabel = [System.Windows.Forms.Label]::New()
    $ListViewOptionLabel = [System.Windows.Forms.Label]::New()
    $ListViewOptionPanel.Controls.Add($ListViewOptionLabel)
    $ListViewOptionLabel.ForeColor = [MyConfig]::Colors.LabelFore
    $ListViewOptionLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
    $ListViewOptionLabel.Name = "ListViewOptionLabel"
    $ListViewOptionLabel.Size = [System.Drawing.Size]::New(($ListViewOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
    $ListViewOptionLabel.Text = $Message
    $ListViewOptionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #endregion $ListViewOptionLabel = [System.Windows.Forms.Label]::New()

    # Returns the minimum size required to display the text
    $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($ListViewOptionLabel.Text, [MyConfig]::Font.Regular, $ListViewOptionLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
    $ListViewOptionLabel.Size = [System.Drawing.Size]::New(($ListViewOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

    $TempBottom = $ListViewOptionLabel.Bottom + [MyConfig]::FormSpacer
  }
  else
  {
    $TempBottom = 0
  }

  #region $ListViewOptionListView = [System.Windows.Forms.ListView]::New()
  $ListViewOptionListView = [System.Windows.Forms.ListView]::New()
  $ListViewOptionPanel.Controls.Add($ListViewOptionListView)
  $ListViewOptionListView.BackColor = [MyConfig]::Colors.TextBack
  $ListViewOptionListView.CheckBoxes = $Multi.IsPresent
  $ListViewOptionListView.Font = [MyConfig]::Font.Bold
  $ListViewOptionListView.ForeColor = [MyConfig]::Colors.TextFore
  $ListViewOptionListView.FullRowSelect = $True
  $ListViewOptionListView.GridLines = $True
  $ListViewOptionListView.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
  $ListViewOptionListView.HideSelection = $False
  $ListViewOptionListView.ListViewItemSorter = [MyCustom.ListViewSort]::New()
  $ListViewOptionListView.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($TempBottom + [MyConfig]::FormSpacer))
  $ListViewOptionListView.MultiSelect = $Multi.IsPresent
  $ListViewOptionListView.Name = "LAUListViewOptionListView"
  $ListViewOptionListView.OwnerDraw = $True
  $ListViewOptionListView.ShowGroups = $False
  $ListViewOptionListView.ShowItemToolTips = $PSBoundParameters.ContainsKey("ToolTip")
  $ListViewOptionListView.Size = [System.Drawing.Size]::New(($ListViewOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ([MyConfig]::Font.Height * $Height))
  $ListViewOptionListView.Text = "LAUListViewOptionListView"
  $ListViewOptionListView.View = [System.Windows.Forms.View]::Details
  #endregion $ListViewOptionListView = [System.Windows.Forms.ListView]::New()

  #region ******** Function Start-ListViewOptionListViewDrawColumnHeader ********
  function Start-ListViewOptionListViewDrawColumnHeader
  {
    <#
      .SYNOPSIS
        DrawColumnHeader Event for the ListViewOption ListView Control
      .DESCRIPTION
        DrawColumnHeader Event for the ListViewOption ListView Control
      .PARAMETER Sender
         The ListView Control that fired the DrawColumnHeader Event
      .PARAMETER EventArg
         The Event Arguments for the ListView DrawColumnHeader Event
      .EXAMPLE
         Start-ListViewOptionListViewDrawColumnHeader -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListView]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter DrawColumnHeader Event for `$ListViewOptionListView"

    [MyConfig]::AutoExit = 0

    $EventArg.Graphics.FillRectangle(([System.Drawing.SolidBrush]::New([MyConfig]::Colors.TitleBack)), $EventArg.Bounds)
    $EventArg.Graphics.DrawRectangle(([System.Drawing.Pen]::New([MyConfig]::Colors.TitleFore)), $EventArg.Bounds.X, $EventArg.Bounds.Y, $EventArg.Bounds.Width, ($EventArg.Bounds.Height - 1))
    $EventArg.Graphics.DrawString($EventArg.Header.Text, $Sender.Font, ([System.Drawing.SolidBrush]::New([MyConfig]::Colors.TitleFore)), ($EventArg.Bounds.X + [MyConfig]::FormSpacer), ($EventArg.Bounds.Y + (($EventArg.Bounds.Height - [MyConfig]::Font.Height) / 1)))

    Write-Verbose -Message "Exit DrawColumnHeader Event for `$ListViewOptionListView"
  }
  #endregion ******** Function Start-ListViewOptionListViewDrawColumnHeader ********
  $ListViewOptionListView.add_DrawColumnHeader({Start-ListViewOptionListViewDrawColumnHeader -Sender $This -EventArg $PSItem})

  #region ******** Function Start-ListViewOptionListViewDrawItem ********
  function Start-ListViewOptionListViewDrawItem
  {
    <#
      .SYNOPSIS
        DrawItem Event for the ListViewOption ListView Control
      .DESCRIPTION
        DrawItem Event for the ListViewOption ListView Control
      .PARAMETER Sender
         The ListView Control that fired the DrawItem Event
      .PARAMETER EventArg
         The Event Arguments for the ListView DrawItem Event
      .EXAMPLE
         Start-ListViewOptionListViewDrawItem -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListView]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter DrawItem Event for `$ListViewOptionListView"

    [MyConfig]::AutoExit = 0

    # Return to Default Draw
    $EventArg.DrawDefault = $True

    Write-Verbose -Message "Exit DrawItem Event for `$ListViewOptionListView"
  }
  #endregion ******** Function Start-ListViewOptionListViewDrawItem ********
  $ListViewOptionListView.add_DrawItem({Start-ListViewOptionListViewDrawItem -Sender $This -EventArg $PSItem})

  #region ******** Function Start-ListViewOptionListViewDrawSubItem ********
  function Start-ListViewOptionListViewDrawSubItem
  {
    <#
      .SYNOPSIS
        DrawSubItem Event for the ListViewOption ListView Control
      .DESCRIPTION
        DrawSubItem Event for the ListViewOption ListView Control
      .PARAMETER Sender
         The ListView Control that fired the DrawSubItem Event
      .PARAMETER EventArg
         The Event Arguments for the ListView DrawSubItem Event
      .EXAMPLE
         Start-ListViewOptionListViewDrawSubItem -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListView]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter DrawSubItem Event for `$ListViewOptionListView"

    [MyConfig]::AutoExit = 0

    # Return to Default Draw
    $EventArg.DrawDefault = $True

    Write-Verbose -Message "Exit DrawSubItem Event for `$ListViewOptionListView"
  }
  #endregion ******** Function Start-ListViewOptionListViewDrawSubItem ********
  $ListViewOptionListView.add_DrawSubItem({Start-ListViewOptionListViewDrawSubItem -Sender $This -EventArg $PSItem})

  #region ******** Function Start-ListViewOptionListViewMouseDown ********
  function Start-ListViewOptionListViewMouseDown
  {
    <#
      .SYNOPSIS
        MouseDown Event for the IDP TreeView Control
      .DESCRIPTION
        MouseDown Event for the IDP TreeView Control
      .PARAMETER Sender
         The TreeView Control that fired the MouseDown Event
      .PARAMETER EventArg
         The Event Arguments for the TreeView MouseDown Event
      .EXAMPLE
         Start-ListViewOptionListViewMouseDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListView]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter MouseDown Event for `$ListViewOptionListView"

    [MyConfig]::AutoExit = 0

    If ($EventArg.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
      if ($ListViewOptionListView.Items.Count -gt 0)
      {
        $ListViewOptionContextMenuStrip.Show($ListViewOptionListView, $EventArg.Location)
      }
    }

    Write-Verbose -Message "Exit MouseDown Event for `$ListViewOptionListView"
  }
  #endregion ******** Function Start-ListViewOptionListViewMouseDown ********
  $ListViewOptionListView.add_MouseDown({ Start-ListViewOptionListViewMouseDown -Sender $This -EventArg $PSItem })


  foreach ($PropName in $Property)
  {
    [Void]$ListViewOptionListView.Columns.Add($PropName, -2)
  }
  [Void]$ListViewOptionListView.Columns.Add(" ", ($ListViewOptionForm.Width * 2))

  ForEach ($Item in $Items)
  {
    ($ListViewOptionListView.Items.Add(($ListViewItem = [System.Windows.Forms.ListViewItem]::New("$($Item.($Property[0]))")))).SubItems.AddRange(@($Property[1..99] | ForEach-Object -Process { "$($Item.($PSItem))" }))
    $ListViewItem.Name = "$($Item.($Property[0]))"
    $ListViewItem.Tag = $Item
    $ListViewItem.Tooltiptext = "$($Item.($Tooltip))"
    $ListViewItem.Selected = ($Item -in $Selected)
    $ListViewItem.Checked = ($Multi.IsPresent -and $ListViewItem.Selected)
    $ListViewItem.Font = [MyConfig]::Font.Regular
  }
  $ListViewOptionListView.Tag = @($ListViewOptionListView.Items)

  If ($Filter.IsPresent)
  {
    #region $ListViewOptionFilterLabel = [System.Windows.Forms.Label]::New()
    $ListViewOptionFilterLabel = [System.Windows.Forms.Label]::New()
    $ListViewOptionPanel.Controls.Add($ListViewOptionFilterLabel)
    $ListViewOptionFilterLabel.AutoSize = $True
    $ListViewOptionFilterLabel.BackColor = [MyConfig]::Colors.Back
    $ListViewOptionFilterLabel.Font = [MyConfig]::Font.Regular
    $ListViewOptionFilterLabel.ForeColor = [MyConfig]::Colors.Fore
    $ListViewOptionFilterLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($ListViewOptionListView.Bottom + [MyConfig]::FormSpacer))
    $ListViewOptionFilterLabel.Name = "ListViewOptionFilterLabel"
    $ListViewOptionFilterLabel.Size = [System.Drawing.Size]::New(([MyConfig]::Font.Width * [MyConfig]::LabelWidth), $ListViewOptionFilterLabel.PreferredHeight)
    $ListViewOptionFilterLabel.Text = "Filter List:"
    $ListViewOptionFilterLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    #endregion $ListViewOptionFilterLabel = [System.Windows.Forms.Label]::New()

    #region $ListViewOptionTextBox = [System.Windows.Forms.TextBox]::New()
    $ListViewOptionTextBox = [System.Windows.Forms.TextBox]::New()
    $ListViewOptionPanel.Controls.Add($ListViewOptionTextBox)
    $ListViewOptionTextBox.AutoSize = $False
    $ListViewOptionTextBox.BackColor = [MyConfig]::Colors.TextBack
    $ListViewOptionTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ListViewOptionTextBox.Font = [MyConfig]::Font.Regular
    $ListViewOptionTextBox.ForeColor = [MyConfig]::Colors.TextFore
    $ListViewOptionTextBox.Location = [System.Drawing.Point]::New(($ListViewOptionFilterLabel.Right + [MyConfig]::FormSpacer), $ListViewOptionFilterLabel.Top)
    $ListViewOptionTextBox.MaxLength = 100
    $ListViewOptionTextBox.Name = "ListViewOptionTextBox"
    $ListViewOptionTextBox.Size = [System.Drawing.Size]::New(($ListViewOptionListView.Right - $ListViewOptionTextBox.Left), $ListViewOptionFilterLabel.Height)
    #$ListViewOptionTextBox.TabIndex = 0
    $ListViewOptionTextBox.TabStop = $False
    $ListViewOptionTextBox.Tag = @{ "HintText" = "Enter Text and Press [Enter] to Filter List Items."; "HintEnabled" = $True }
    $ListViewOptionTextBox.Text = ""
    $ListViewOptionTextBox.WordWrap = $False
    #endregion $ListViewOptionTextBox = [System.Windows.Forms.TextBox]::New()

    #region ******** Function Start-ListViewOptionTextBoxGotFocus ********
    Function Start-ListViewOptionTextBoxGotFocus
    {
      <#
        .SYNOPSIS
          GotFocus Event for the ListViewOption TextBox Control
        .DESCRIPTION
          GotFocus Event for the ListViewOption TextBox Control
        .PARAMETER Sender
           The TextBox Control that fired the GotFocus Event
        .PARAMETER EventArg
           The Event Arguments for the TextBox GotFocus Event
        .EXAMPLE
           Start-ListViewOptionTextBoxGotFocus -Sender $Sender -EventArg $EventArg
        .NOTES
          Original Function By ken.sweet
      #>
      [CmdletBinding()]
      Param (
        [parameter(Mandatory = $True)]
        [System.Windows.Forms.TextBox]$Sender,
        [parameter(Mandatory = $True)]
        [Object]$EventArg
      )
      Write-Verbose -Message "Enter GotFocus Event for `$ListViewOptionTextBox"

      [MyConfig]::AutoExit = 0

      # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
      If ($Sender.Tag.HintEnabled)
      {
        $Sender.Text = ""
        $Sender.Font = [MyConfig]::Font.Regular
        $Sender.ForeColor = [MyConfig]::Colors.TextFore
      }

      Write-Verbose -Message "Exit GotFocus Event for `$ListViewOptionTextBox"
    }
    #endregion ******** Function Start-ListViewOptionTextBoxGotFocus ********
    $ListViewOptionTextBox.add_GotFocus({ Start-ListViewOptionTextBoxGotFocus -Sender $This -EventArg $PSItem })

    #region ******** Function Start-ListViewOptionTextBoxKeyDown ********
    Function Start-ListViewOptionTextBoxKeyDown
    {
      <#
        .SYNOPSIS
          KeyDown Event for the ListViewOption TextBox Control
        .DESCRIPTION
          KeyDown Event for the ListViewOption TextBox Control
        .PARAMETER Sender
           The TextBox Control that fired the KeyDown Event
        .PARAMETER EventArg
           The Event Arguments for the TextBox KeyDown Event
        .EXAMPLE
           Start-ListViewOptionTextBoxKeyDown -Sender $Sender -EventArg $EventArg
        .NOTES
          Original Function By ken.sweet
      #>
      [CmdletBinding()]
      Param (
        [parameter(Mandatory = $True)]
        [System.Windows.Forms.TextBox]$Sender,
        [parameter(Mandatory = $True)]
        [Object]$EventArg
      )
      Write-Verbose -Message "Enter KeyDown Event for `$ListViewOptionTextBox"

      [MyConfig]::AutoExit = 0

      If ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Enter)
      {
        # Suppress KeyPress
        $EventArg.SuppressKeyPress = $True

        If ([String]::IsNullOrEmpty($Sender.Text.Trim()))
        {
          $ListViewOptionListView.Items.Clear()
          $ListViewOptionListView.Items.AddRange($ListViewOptionListView.Tag)
        }
        else
        {
          $TmpNewList = @($ListViewOptionListView.Tag | Where-Object -FilterScript { ($PSItem.Text -Match $Sender.Text) -or ($PSItem.SubItems[1].Text -Match $Sender.Text) })
          $ListViewOptionListView.Items.Clear()
          $ListViewOptionListView.Items.AddRange($TmpNewList)
        }
      }

      Write-Verbose -Message "Exit KeyDown Event for `$ListViewOptionTextBox"
    }
    #endregion ******** Function Start-ListViewOptionTextBoxKeyDown ********
    $ListViewOptionTextBox.add_KeyDown({ Start-ListViewOptionTextBoxKeyDown -Sender $This -EventArg $PSItem })

    #region ******** Function Start-ListViewOptionTextBoxLostFocus ********
    Function Start-ListViewOptionTextBoxLostFocus
    {
      <#
        .SYNOPSIS
          LostFocus Event for the ListViewOption TextBox Control
        .DESCRIPTION
          LostFocus Event for the ListViewOption TextBox Control
        .PARAMETER Sender
           The TextBox Control that fired the LostFocus Event
        .PARAMETER EventArg
           The Event Arguments for the TextBox LostFocus Event
        .EXAMPLE
           Start-ListViewOptionTextBoxLostFocus -Sender $Sender -EventArg $EventArg
        .NOTES
          Original Function By ken.sweet
      #>
      [CmdletBinding()]
      Param (
        [parameter(Mandatory = $True)]
        [System.Windows.Forms.TextBox]$Sender,
        [parameter(Mandatory = $True)]
        [Object]$EventArg
      )
      Write-Verbose -Message "Enter LostFocus Event for `$ListViewOptionTextBox"

      [MyConfig]::AutoExit = 0

      # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
      If ([String]::IsNullOrEmpty(($Sender.Text.Trim())))
      {
        $Sender.Text = $Sender.Tag.HintText
        $Sender.Tag.HintEnabled = $True
        $Sender.Font = [MyConfig]::Font.Hint
        $Sender.ForeColor = [MyConfig]::Colors.TextHint

        $ListViewOptionListView.Items.Clear()
        $ListViewOptionListView.Items.AddRange($ListViewOptionListView.Tag)
      }
      Else
      {
        $Sender.Tag.HintEnabled = $False
        $Sender.Font = [MyConfig]::Font.Regular
        $Sender.ForeColor = [MyConfig]::Colors.TextFore

        $TmpNewList = @($ListViewOptionListView.Tag | Where-Object -FilterScript { ($PSItem.Text -Match $ListViewOptionTextBox.Text) -or ($PSItem.SubItems[1].Text -Match $ListViewOptionTextBox.Text) })
        $ListViewOptionListView.Items.Clear()
        $ListViewOptionListView.Items.AddRange($TmpNewList)
      }

      Write-Verbose -Message "Exit LostFocus Event for `$ListViewOptionTextBox"
    }
    #endregion ******** Function Start-ListViewOptionTextBoxLostFocus ********
    $ListViewOptionTextBox.add_LostFocus({ Start-ListViewOptionTextBoxLostFocus -Sender $This -EventArg $PSItem })

    Start-ListViewOptionTextBoxLostFocus -Sender $ListViewOptionTextBox -EventArg "Lost Focus"

    $TempClientSize = [System.Drawing.Size]::New(($ListViewOptionTextBox.Right + [MyConfig]::FormSpacer), ($ListViewOptionTextBox.Bottom + [MyConfig]::FormSpacer))
  }
  Else
  {
    $TempClientSize = [System.Drawing.Size]::New(($ListViewOptionListView.Right + [MyConfig]::FormSpacer), ($ListViewOptionListView.Bottom + [MyConfig]::FormSpacer))
  }

  # ************************************************
  # ListViewOption ContextMenuStrip
  # ************************************************
  #region $ListViewOptionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  $ListViewOptionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  #$ListViewOptionListView.Controls.Add($ListViewOptionContextMenuStrip)
  $ListViewOptionContextMenuStrip.BackColor = [MyConfig]::Colors.Back
  #$ListViewOptionContextMenuStrip.Enabled = $True
  $ListViewOptionContextMenuStrip.Font = [MyConfig]::Font.Regular
  $ListViewOptionContextMenuStrip.ForeColor = [MyConfig]::Colors.Fore
  $ListViewOptionContextMenuStrip.ImageList = $PILSmallImageList
  $ListViewOptionContextMenuStrip.Name = "ListViewOptionContextMenuStrip"
  #endregion $ListViewOptionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()

  #region ******** Function Start-ListViewOptionContextMenuStripOpening ********
  function Start-ListViewOptionContextMenuStripOpening
  {
    <#
      .SYNOPSIS
        Opening Event for the ListViewOption ContextMenuStrip Control
      .DESCRIPTION
        Opening Event for the ListViewOption ContextMenuStrip Control
      .PARAMETER Sender
         The ContextMenuStrip Control that fired the Opening Event
      .PARAMETER EventArg
         The Event Arguments for the ContextMenuStrip Opening Event
      .EXAMPLE
         Start-ListViewOptionContextMenuStripOpening -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ContextMenuStrip]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Opening Event for `$ListViewOptionContextMenuStrip"

    [MyConfig]::AutoExit = 0

    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

    Write-Verbose -Message "Exit Opening Event for `$ListViewOptionContextMenuStrip"
  }
  #endregion ******** Function Start-ListViewOptionContextMenuStripOpening ********
  $ListViewOptionContextMenuStrip.add_Opening({Start-ListViewOptionContextMenuStripOpening -Sender $This -EventArg $PSItem})

  #region ******** Function Start-ListViewOptionContextMenuStripItemClick ********
  function Start-ListViewOptionContextMenuStripItemClick
  {
    <#
      .SYNOPSIS
        Click Event for the ListViewOption ToolStripItem Control
      .DESCRIPTION
        Click Event for the ListViewOption ToolStripItem Control
      .PARAMETER Sender
         The ToolStripItem Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the ToolStripItem Click Event
      .EXAMPLE
         Start-ListViewOptionContextMenuStripItemClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ToolStripItem]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ListViewOptionContextMenuStripItem"

    [MyConfig]::AutoExit = 0

    switch ($Sender.Name)
    {
      "CheckAll"
      {
        $TmpChecked = @($ListViewOptionListView.Items | Where-Object -FilterScript { -not $PSItem.Checked })
        $TmpChecked | ForEach-Object -Process { $PSItem.Checked = $True }
        Break
      }
      "UnCheckAll"
      {
        $TmpChecked = @($ListViewOptionListView.Items | Where-Object -FilterScript { $PSItem.Checked })
        $TmpChecked | ForEach-Object -Process { $PSItem.Checked = $False }
        Break
      }
    }

    Write-Verbose -Message "Exit Click Event for `$ListViewOptionContextMenuStripItem"
  }
  #endregion ******** Function Start-ListViewOptionContextMenuStripItemClick ********

  (New-MenuItem -Menu $ListViewOptionContextMenuStrip -Text "Check All" -Name "CheckAll" -Tag "CheckAll" -DisplayStyle "ImageAndText" -ImageKey "CheckIcon" -PassThru).add_Click({Start-ListViewOptionContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $ListViewOptionContextMenuStrip -Text "Uncheck All" -Name "UnCheckAll" -Tag "UnCheckAll" -DisplayStyle "ImageAndText" -ImageKey "UncheckIcon" -PassThru).add_Click({Start-ListViewOptionContextMenuStripItemClick -Sender $This -EventArg $PSItem})

  #endregion ******** $ListViewOptionPanel Controls ********

  # ************************************************
  # ListViewOptionBtm Panel
  # ************************************************
  #region $ListViewOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $ListViewOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $ListViewOptionForm.Controls.Add($ListViewOptionBtmPanel)
  $ListViewOptionBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ListViewOptionBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $ListViewOptionBtmPanel.Name = "ListViewOptionBtmPanel"
  #endregion $ListViewOptionBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ListViewOptionBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($ListViewOptionBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $ListViewOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ListViewOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ListViewOptionBtmPanel.Controls.Add($ListViewOptionBtmLeftButton)
  $ListViewOptionBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $ListViewOptionBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ListViewOptionBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ListViewOptionBtmLeftButton.Font = [MyConfig]::Font.Bold
  $ListViewOptionBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ListViewOptionBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $ListViewOptionBtmLeftButton.Name = "ListViewOptionBtmLeftButton"
  $ListViewOptionBtmLeftButton.TabIndex = 1
  $ListViewOptionBtmLeftButton.TabStop = $True
  $ListViewOptionBtmLeftButton.Text = $ButtonLeft
  $ListViewOptionBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $ListViewOptionBtmLeftButton.PreferredSize.Height)
  #endregion $ListViewOptionBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ListViewOptionBtmLeftButtonClick ********
  function Start-ListViewOptionBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ListViewOptionBtmLeft Button Control
      .DESCRIPTION
        Click Event for the ListViewOptionBtmLeft Button Control
      .PARAMETER Sender
         The Button Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the Button Click Event
      .EXAMPLE
         Start-ListViewOptionBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ListViewOptionBtmLeftButton"

    [MyConfig]::AutoExit = 0

    if ((($ListViewOptionListView.CheckedItems.Count -gt 0) -and ((-not $Multi.IsPresent) -or $Multi.IsPresent)) -or (-not $Required.IsPresent))
    {
      $ListViewOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    else
    {
      [Void][System.Windows.Forms.MessageBox]::Show($ListViewOptionForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }

    Write-Verbose -Message "Exit Click Event for `$ListViewOptionBtmLeftButton"
  }
  #endregion ******** Function Start-ListViewOptionBtmLeftButtonClick ********
  $ListViewOptionBtmLeftButton.add_Click({ Start-ListViewOptionBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $ListViewOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $ListViewOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $ListViewOptionBtmPanel.Controls.Add($ListViewOptionBtmMidButton)
  $ListViewOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $ListViewOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $ListViewOptionBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ListViewOptionBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ListViewOptionBtmMidButton.Font = [MyConfig]::Font.Bold
  $ListViewOptionBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ListViewOptionBtmMidButton.Location = [System.Drawing.Point]::New(($ListViewOptionBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ListViewOptionBtmMidButton.Name = "ListViewOptionBtmMidButton"
  $ListViewOptionBtmMidButton.TabIndex = 2
  $ListViewOptionBtmMidButton.TabStop = $True
  $ListViewOptionBtmMidButton.Text = $ButtonMid
  $ListViewOptionBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $ListViewOptionBtmMidButton.PreferredSize.Height)
  #endregion $ListViewOptionBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ListViewOptionBtmMidButtonClick ********
  function Start-ListViewOptionBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ListViewOptionBtmMid Button Control
      .DESCRIPTION
        Click Event for the ListViewOptionBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ListViewOptionBtmMidButtonClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By MyUserName)
  #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ListViewOptionBtmMidButton"

    [MyConfig]::AutoExit = 0

    if ([String]::IsNullOrEmpty($Selected))
    {
      $ListViewOptionListView.SelectedItems.Clear()
      $ListViewOptionListView.Items | ForEach-Object -Process { $PSItem.Checked = $False }
    }
    else
    {
      foreach ($Item in $ListViewOptionListView.Items)
      {
        $Item.Selected = ($Item.Tag -in $Selected)
        $Item.Checked = ($Multi.IsPresent -and $Item.Selected)
      }
    }
    $ListViewOptionListView.Refresh()
    $ListViewOptionListView.Select()

    Write-Verbose -Message "Exit Click Event for `$ListViewOptionBtmMidButton"
  }
  #endregion ******** Function Start-ListViewOptionBtmMidButtonClick ********
  $ListViewOptionBtmMidButton.add_Click({ Start-ListViewOptionBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $ListViewOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $ListViewOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $ListViewOptionBtmPanel.Controls.Add($ListViewOptionBtmRightButton)
  $ListViewOptionBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $ListViewOptionBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ListViewOptionBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ListViewOptionBtmRightButton.Font = [MyConfig]::Font.Bold
  $ListViewOptionBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ListViewOptionBtmRightButton.Location = [System.Drawing.Point]::New(($ListViewOptionBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ListViewOptionBtmRightButton.Name = "ListViewOptionBtmRightButton"
  $ListViewOptionBtmRightButton.TabIndex = 3
  $ListViewOptionBtmRightButton.TabStop = $True
  $ListViewOptionBtmRightButton.Text = $ButtonRight
  $ListViewOptionBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $ListViewOptionBtmRightButton.PreferredSize.Height)
  #endregion $ListViewOptionBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ListViewOptionBtmRightButtonClick ********
  function Start-ListViewOptionBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ListViewOptionBtmRight Button Control
      .DESCRIPTION
        Click Event for the ListViewOptionBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ListViewOptionBtmRightButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ListViewOptionBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here

    $ListViewOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$ListViewOptionBtmRightButton"
  }
  #endregion ******** Function Start-ListViewOptionBtmRightButtonClick ********
  $ListViewOptionBtmRightButton.add_Click({ Start-ListViewOptionBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $ListViewOptionBtmPanel.ClientSize = [System.Drawing.Size]::New(($ListViewOptionBtmRightButton.Right + [MyConfig]::FormSpacer), ($ListViewOptionBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ListViewOptionBtmPanel Controls ********

  $ListViewOptionForm.ClientSize = [System.Drawing.Size]::New($ListViewOptionForm.ClientSize.Width, ($TempClientSize.Height + $ListViewOptionBtmPanel.Height))
  $ListViewOptionForm.MinimumSize = $ListViewOptionForm.Size
  $ListViewOptionListView.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom, Right")
  If ($Filter.IsPresent)
  {
    $ListViewOptionFilterLabel.Anchor = [System.Windows.Forms.AnchorStyles]("Left, Bottom")
    $ListViewOptionTextBox.Anchor = [System.Windows.Forms.AnchorStyles]("Left, Bottom, Right")
  }

  #endregion ******** Controls for ListViewOption Form ********

  #endregion ******** End **** ListViewOption **** End ********

  $DialogResult = $ListViewOptionForm.ShowDialog($PILForm)
  if ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
  {
    if ($Multi.IsPresent)
    {
      [ListViewOption]::New($True, $DialogResult, ($ListViewOptionListView.CheckedItems | Select-Object -ExpandProperty "Tag"))
    }
    else
    {
      [ListViewOption]::New($True, $DialogResult, ($ListViewOptionListView.SelectedItems | Select-Object -ExpandProperty "Tag"))
    }
  }
  else
  {
    [ListViewOption]::New($False, $DialogResult, "")
  }

  $ListViewOptionForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-ListViewOption"
}
#endregion function Get-ListViewOption

# ---------------------------
# Show RichTextStatus Function
# ---------------------------
#region Function Write-RichTextBox
Function Write-RichTextBox
{
  <#
    .SYNOPSIS
      Write to RichTextBox
    .DESCRIPTION
      Write to RichTextBox
    .PARAMETER RichTextBox
    .PARAMETER TextFore
    .PARAMETER Font
    .PARAMETER Alignment
    .PARAMETER Text
    .PARAMETER BulletFore
    .PARAMETER NoNewLine
    .EXAMPLE
      Write-RichTextBox -RichTextBox $RichTextBox -Text $Text
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "NewLine")]
  param (
    [System.Windows.Forms.RichTextBox]$RichTextBox = $RichTextStatusRichTextBox,
    [System.Drawing.Color]$TextFore = [MyConfig]::Colors.TextFore,
    [System.Drawing.Font]$Font = [MyConfig]::Font.Regular,
    [System.Windows.Forms.HorizontalAlignment]$Alignment = [System.Windows.Forms.HorizontalAlignment]::Left,
    [String]$Text,
    [parameter(Mandatory = $False, ParameterSetName = "NewLine")]
    [System.Drawing.Color]$BulletFore = [MyConfig]::Colors.TextFore,
    [parameter(Mandatory = $True, ParameterSetName = "NoNewLine")]
    [Switch]$NoNewLine
  )
  $RichTextBox.SelectionLength = 0
  $RichTextBox.SelectionStart = $RichTextBox.TextLength
  $RichTextBox.SelectionAlignment = $Alignment
  $RichTextBox.SelectionFont = $Font
  $RichTextBox.SelectionColor = $TextFore
  $RichTextBox.AppendText($Text)
  if (-not $NoNewLine.IsPresent)
  {
    $RichTextBox.SelectionColor = $BulletFore
    $RichTextBox.AppendText("`r`n")
  }
  $RichTextBox.ScrollToCaret()
  $RichTextBox.Refresh()
  $RichTextBox.Parent.Parent.Activate()
  [System.Windows.Forms.Application]::DoEvents()
}
#endregion Function Write-RichTextBox

#region Function Write-RichTextBoxValue
Function Write-RichTextBoxValue
{
  <#
    .SYNOPSIS
      Write Property Value to RichTextBox
    .DESCRIPTION
      Write Property Value to RichTextBox
    .PARAMETER RichTextBox
    .PARAMETER TextFore
    .PARAMETER ValueFore
    .PARAMETER BulletFore
    .PARAMETER Font
    .PARAMETER Text
    .PARAMETER Value
    .EXAMPLE
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text $Text -Value $Value
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  Param (
    [System.Windows.Forms.RichTextBox]$RichTextBox = $RichTextStatusRichTextBox,
    [System.Drawing.Color]$TextFore = [MyConfig]::Colors.TextFore,
    [System.Drawing.Color]$ValueFore = [MyConfig]::Colors.TextInfo,
    [System.Drawing.Color]$BulletFore = [MyConfig]::Colors.TextFore,
    [System.Drawing.Font]$Font = [MyConfig]::Font.Regular,
    [Parameter(Mandatory = $True)]
    [String]$Text,
    [Parameter(Mandatory = $True)]
    [AllowEmptyString()]
    [AllowNull()]
    [String]$Value
  )
  $RichTextBox.SelectionLength = 0
  $RichTextBox.SelectionStart = $RichTextBox.TextLength
  $RichTextBox.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Left
  $RichTextBox.SelectionFont = $Font
  $RichTextBox.SelectionColor = $TextFore
  $RichTextBox.AppendText("$($Text)")
  $RichTextBox.SelectionColor = $BulletFore
  $RichTextBox.AppendText(": ")
  $RichTextBox.SelectionColor = $ValueFore
  $RichTextBox.AppendText("$($Value)")
  $RichTextBox.SelectionColor = $BulletFore
  $RichTextBox.AppendText("`r`n")
  $RichTextBox.ScrollToCaret()
  $RichTextBox.Refresh()
  $RichTextBox.Parent.Parent.Activate()
  [System.Windows.Forms.Application]::DoEvents()
}
#endregion Function Write-RichTextBoxValue

#region Function Write-RichTextBoxError
Function Write-RichTextBoxError
{
  <#
    .SYNOPSIS
      Write Error Message to RichTextBox
    .DESCRIPTION
      Write Error Message to RichTextBox
    .PARAMETER RichTextBox
    .EXAMPLE
      Write-RichTextBoxError -RichTextBox $RichTextBox
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  param (
    [System.Windows.Forms.RichTextBox]$RichTextBox = $RichTextStatusRichTextBox
  )
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value ($($Error[0].Exception.Message)) -ValueFore ([MyConfig]::Colors.TextFore)
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "CODE" -TextFore ([MyConfig]::Colors.TextBad) -Value (($Error[0].InvocationInfo.Line).Trim()) -ValueFore ([MyConfig]::Colors.TextFore)
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "LINE" -TextFore ([MyConfig]::Colors.TextBad) -Value ($Error[0].InvocationInfo.ScriptLineNumber) -ValueFore ([MyConfig]::Colors.TextFore)
}
#endregion Function Write-RichTextBoxError

#region RichTextStatus Result Class
Class RichTextStatus
{
  [Bool]$Success
  [Object]$DialogResult

  RichTextStatus ([Bool]$Success, [Object]$DialogResult)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
  }
}
#endregion RichTextStatus Result Class

#region function Show-RichTextStatus
function Show-RichTextStatus ()
{
  <#
    .SYNOPSIS
      Shows Show-RichTextStatus
    .DESCRIPTION
      Shows Show-RichTextStatus
    .PARAMETER Title
      Title of the Show-RichTextStatus Dialog Window
    .PARAMETER ScriptBlock
      Script Block to Execure
    .PARAMETER HashTable
      HashTable of Paramerts to Pass to the ScriptBlock
    .PARAMETER Width
      Width of the Show-RichTextStatus Dialog Window
    .PARAMETER Height
      Height of the Show-RichTextStatus Dialog Window
    .PARAMETER ButtonDefault
      The Default Selected Button
    .PARAMETER ButtonLeft
      The DialogResult of the Left Button
    .PARAMETER ButtonMid
      The DialogResult of the Middle Button
    .PARAMETER ButtonRight
      The DialogResult of the Right Button
    .PARAMETER AllowControl
      Enable Pause and Break out of Script Block
    .PARAMETER AutoClose
      Auto Close the Status Message Dialog Window
    .PARAMETER AutoCloseWait
      Number of MilliSeconds to wait Before Auto Closing the Dialog Window
    .EXAMPLE
      $HashTable = @{"ShowHeader" = $True}
      $ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Sample-RichTextStatus -RichTextBox $RichTextBox -HashTable $HashTable }
      $DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title "Initializing $([MyConfig]::ScriptName)" -ButtonMid "OK" -HashTable $HashTable
      if ($DialogResult.Success)
      {
        # Success
      }
      else
      {
        # Failed
      }
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Zero")]
  param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [parameter(Mandatory = $True)]
    [ScriptBlock]$ScriptBlock = { },
    [HashTable]$HashTable = @{ },
    [Int]$Width = 45,
    [Int]$Height = 30,
    [System.Windows.Forms.DialogResult]$ButtonDefault = "OK",
    [parameter(Mandatory = $True, ParameterSetName = "Two")]
    [parameter(Mandatory = $True, ParameterSetName = "Three")]
    [System.Windows.Forms.DialogResult]$ButtonLeft,
    [parameter(Mandatory = $True, ParameterSetName = "One")]
    [parameter(Mandatory = $True, ParameterSetName = "Three")]
    [System.Windows.Forms.DialogResult]$ButtonMid,
    [parameter(Mandatory = $True, ParameterSetName = "Two")]
    [parameter(Mandatory = $True, ParameterSetName = "Three")]
    [System.Windows.Forms.DialogResult]$ButtonRight,
    [Switch]$AllowControl,
    [Switch]$AutoClose,
    [ValidateRange(0, 60000)]
    [int]$AutoCloseWait = 10
  )
  Write-Verbose -Message "Enter Function Show-RichTextStatus"

  #region ******** Begin **** $Show-RichTextStatus **** Begin ********

  # ************************************************
  # $RichTextStatus Form
  # ************************************************
  #region $RichTextStatusForm = [System.Windows.Forms.Form]::New()
  $RichTextStatusForm = [System.Windows.Forms.Form]::New()
  $RichTextStatusForm.BackColor = [MyConfig]::Colors.Back
  $RichTextStatusForm.Font = [MyConfig]::Font.Regular
  $RichTextStatusForm.ForeColor = [MyConfig]::Colors.Fore
  $RichTextStatusForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $RichTextStatusForm.Icon = $PILForm.Icon
  $RichTextStatusForm.KeyPreview = $AllowControl.IsPresent
  $RichTextStatusForm.MaximizeBox = $False
  $RichTextStatusForm.MinimizeBox = $False
  $RichTextStatusForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * $Height))
  $RichTextStatusForm.Name = "RichTextStatusForm"
  $RichTextStatusForm.Owner = $PILForm
  $RichTextStatusForm.ShowInTaskbar = $False
  $RichTextStatusForm.Size = $RichTextStatusForm.MinimumSize
  $RichTextStatusForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $RichTextStatusForm.Tag = @{ "Cancel" = $False; "Pause" = $False; "Finished" = $False }
  $RichTextStatusForm.Text = $Title
  #endregion $RichTextStatusForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-RichTextStatusFormKeyDown ********
  Function Start-RichTextStatusFormKeyDown
  {
  <#
    .SYNOPSIS
      KeyDown Event for the RichTextStatus Form Control
    .DESCRIPTION
      KeyDown Event for the RichTextStatus Form Control
    .PARAMETER Sender
       The Form Control that fired the KeyDown Event
    .PARAMETER EventArg
       The Event Arguments for the Form KeyDown Event
    .EXAMPLE
       Start-RichTextStatusFormKeyDown -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$RichTextStatusForm"

    [MyConfig]::AutoExit = 0

    If ($EventArg.Control -and $EventArg.Alt)
    {
      Switch ($EventArg.KeyCode)
      {
        { $PSItem -in ([System.Windows.Forms.Keys]::Back, [System.Windows.Forms.Keys]::End) }
        {
          $Sender.Tag.Cancel = $True
          Break
        }
      }
    }
    Else
    {
      Switch ($EventArg.KeyCode)
      {
        { $PSItem -eq [System.Windows.Forms.Keys]::Pause }
        {
          $Sender.Tag.Pause = (-not $Sender.Tag.Pause)
          Break
        }
        { $PSItem -in ([System.Windows.Forms.Keys]::Enter, [System.Windows.Forms.Keys]::Space, [System.Windows.Forms.Keys]::Escape) }
        {
          if ($Sender.Tag.Finished)
          {
            $Sender.DialogResult = $ButtonDefault
          }
          Break
        }
      }
    }

    Write-Verbose -Message "Exit KeyDown Event for `$RichTextStatusForm"
  }
  #endregion ******** Function Start-RichTextStatusFormKeyDown ********
  If ($AllowControl.IsPresent)
  {
    $RichTextStatusForm.add_KeyDown({ Start-RichTextStatusFormKeyDown -Sender $This -EventArg $PSItem })
  }

  #region ******** Function Start-RichTextStatusFormShown ********
  function Start-RichTextStatusFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the $RichTextStatus Form Control
      .DESCRIPTION
        Shown Event for the $RichTextStatus Form Control
      .PARAMETER Sender
         The Form Control that fired the Shown Event
      .PARAMETER EventArg
         The Event Arguments for the Form Shown Event
      .EXAMPLE
         Start-RichTextStatusFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$RichTextStatusForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    If ([MyConfig]::Production)
    {
      # Disable Auto Exit Timer
      $PILTimer.Enabled = $False
    }

    if ($PassHashTable)
    {
      $DialogResult = Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $RichTextStatusRichTextBox, $HashTable
    }
    else
    {
      $DialogResult = Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $RichTextStatusRichTextBox
    }

    If ([MyConfig]::Production)
    {
      # Re-enable Auto Exit Timer
      $PILTimer.Enabled = ([MyConfig]::AutoExitMax -gt 0)
    }

    switch ($RichTextStatusButtons)
    {
      1
      {
        $RichTextStatusBtmMidButton.Enabled = $True
        $RichTextStatusBtmMidButton.DialogResult = $DialogResult
        Break
      }
      2
      {
        $RichTextStatusBtmLeftButton.Enabled = $True
        $RichTextStatusBtmRightButton.Enabled = $True
        Break
      }
      3
      {
        $RichTextStatusBtmLeftButton.Enabled = $True
        $RichTextStatusBtmMidButton.Enabled = $True
        $RichTextStatusBtmRightButton.Enabled = $True
        Break
      }
    }

    $Sender.Tag.Finished = $True

    if ((($DialogResult -eq $ButtonDefault) -and $AutoClose.IsPresent) -or ($RichTextStatusButtons -eq 0))
    {
      $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
      while ($StopWatch.Elapsed.TotalMilliseconds -le $AutoCloseWait)
      {
        [System.Threading.Thread]::Sleep(10)
        [System.Windows.Forms.Application]::DoEvents()
      }

      $Sender.DialogResult = $DialogResult
    }

    Write-Verbose -Message "Exit Shown Event for `$RichTextStatusForm"
  }
  #endregion ******** Function Start-RichTextStatusFormShown ********
  $RichTextStatusForm.add_Shown({ Start-RichTextStatusFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for $RichTextStatus Form ********

  # ************************************************
  # $RichTextStatus Panel
  # ************************************************
  #region $RichTextStatusPanel = [System.Windows.Forms.Panel]::New()
  $RichTextStatusPanel = [System.Windows.Forms.Panel]::New()
  $RichTextStatusForm.Controls.Add($RichTextStatusPanel)
  $RichTextStatusPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $RichTextStatusPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $RichTextStatusPanel.Name = "RichTextStatusPanel"
  #endregion $RichTextStatusPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $RichTextStatusPanel Controls ********

  #region $RichTextStatusRichTextBox = [System.Windows.Forms.RichTextBox]::New()
  $RichTextStatusRichTextBox = [System.Windows.Forms.RichTextBox]::New()
  $RichTextStatusPanel.Controls.Add($RichTextStatusRichTextBox)
  $RichTextStatusRichTextBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
  $RichTextStatusRichTextBox.BackColor = [MyConfig]::Colors.TextBack
  $RichTextStatusRichTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $RichTextStatusRichTextBox.DetectUrls = $True
  $RichTextStatusRichTextBox.Font = [MyConfig]::Font.Regular
  $RichTextStatusRichTextBox.ForeColor = [MyConfig]::Colors.TextFore
  $RichTextStatusRichTextBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $RichTextStatusRichTextBox.MaxLength = [Int]::MaxValue
  $RichTextStatusRichTextBox.Multiline = $True
  $RichTextStatusRichTextBox.Name = "RichTextStatusRichTextBox"
  $RichTextStatusRichTextBox.ReadOnly = $True
  $RichTextStatusRichTextBox.Rtf = ""
  $RichTextStatusRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
  $RichTextStatusRichTextBox.Size = [System.Drawing.Size]::New(($RichTextStatusPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($RichTextStatusPanel.ClientSize.Height - ($RichTextStatusRichTextBox.Top + [MyConfig]::FormSpacer)))
  $RichTextStatusRichTextBox.TabStop = $False
  $RichTextStatusRichTextBox.WordWrap = $False
  #endregion $RichTextStatusRichTextBox = [System.Windows.Forms.RichTextBox]::New()

  #region ******** Function Start-RichTextStatusRichTextBoxMouseDown ********
  Function Start-RichTextStatusRichTextBoxMouseDown
  {
  <#
    .SYNOPSIS
      MouseDown Event for the RichTextStatus RichTextBox Control
    .DESCRIPTION
      MouseDown Event for the RichTextStatus RichTextBox Control
    .PARAMETER Sender
       The RichTextBox Control that fired the MouseDown Event
    .PARAMETER EventArg
       The Event Arguments for the RichTextBox MouseDown Event
    .EXAMPLE
       Start-RichTextStatusRichTextBoxMouseDown -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.RichTextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter MouseDown Event for `$RichTextStatusRichTextBox"

    [MyConfig]::AutoExit = 0

    $RichTextStatusRichTextBox.SelectionLength = 0
    $RichTextStatusRichTextBox.SelectionStart = $RichTextStatusRichTextBox.TextLength

    Write-Verbose -Message "Exit MouseDown Event for `$RichTextStatusRichTextBox"
  }
  #endregion ******** Function Start-RichTextStatusRichTextBoxMouseDown ********
  $RichTextStatusRichTextBox.add_MouseDown({ Start-RichTextStatusRichTextBoxMouseDown -Sender $This -EventArg $PSItem })

  #endregion ******** $RichTextStatusPanel Controls ********

  switch ($PSCmdlet.ParameterSetName)
  {
    "Zero"
    {
      $RichTextStatusButtons = 0
      Break
    }
    "One"
    {
      $RichTextStatusButtons = 1
      Break
    }
    "Two"
    {
      $RichTextStatusButtons = 2
      Break
    }
    "Three"
    {
      $RichTextStatusButtons = 3
      Break
    }
  }

  # Evenly Space Buttons - Move Size to after Text
  if ($RichTextStatusButtons -gt 0)
  {
    # ************************************************
    # $RichTextStatusBtm Panel
    # ************************************************
    #region $RichTextStatusBtmPanel = [System.Windows.Forms.Panel]::New()
    $RichTextStatusBtmPanel = [System.Windows.Forms.Panel]::New()
    $RichTextStatusForm.Controls.Add($RichTextStatusBtmPanel)
    $RichTextStatusBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $RichTextStatusBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $RichTextStatusBtmPanel.Name = "RichTextStatusBtmPanel"
    #endregion $RichTextStatusBtmPanel = [System.Windows.Forms.Panel]::New()

    #region ******** $RichTextStatusBtmPanel Controls ********

    $NumButtons = 3
    $TempSpace = [Math]::Floor($RichTextStatusBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
    $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
    $TempMod = $TempSpace % $NumButtons

    #region $RichTextStatusBtmLeftButton = [System.Windows.Forms.Button]::New()
    If (($RichTextStatusButtons -eq 2) -or ($RichTextStatusButtons -eq 3))
    {
      $RichTextStatusBtmLeftButton = [System.Windows.Forms.Button]::New()
      $RichTextStatusBtmPanel.Controls.Add($RichTextStatusBtmLeftButton)
      $RichTextStatusBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
      $RichTextStatusBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
      $RichTextStatusBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
      $RichTextStatusBtmLeftButton.DialogResult = $ButtonLeft
      $RichTextStatusBtmLeftButton.Enabled = $False
      $RichTextStatusBtmLeftButton.Font = [MyConfig]::Font.Bold
      $RichTextStatusBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
      $RichTextStatusBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
      $RichTextStatusBtmLeftButton.Name = "RichTextStatusBtmLeftButton"
      $RichTextStatusBtmLeftButton.TabIndex = 0
      $RichTextStatusBtmLeftButton.TabStop = $True
      $RichTextStatusBtmLeftButton.Text = "&$($ButtonLeft.ToString())"
      $RichTextStatusBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $RichTextStatusBtmLeftButton.PreferredSize.Height)
      if ($ButtonLeft -eq $ButtonDefault)
      {
        $RichTextStatusBtmLeftButton.Select()
      }
    }
    #endregion $RichTextStatusBtmLeftButton = [System.Windows.Forms.Button]::New()

    #region $RichTextStatusBtmMidButton = [System.Windows.Forms.Button]::New()
    If (($RichTextStatusButtons -eq 1) -or ($RichTextStatusButtons -eq 3))
    {
      $RichTextStatusBtmMidButton = [System.Windows.Forms.Button]::New()
      $RichTextStatusBtmPanel.Controls.Add($RichTextStatusBtmMidButton)
      $RichTextStatusBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
      $RichTextStatusBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
      $RichTextStatusBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
      $RichTextStatusBtmMidButton.DialogResult = $ButtonMid
      $RichTextStatusBtmMidButton.Enabled = $False
      $RichTextStatusBtmMidButton.Font = [MyConfig]::Font.Bold
      $RichTextStatusBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
      $RichTextStatusBtmMidButton.Location = [System.Drawing.Point]::New(($TempWidth + ([MyConfig]::FormSpacer * 2)), [MyConfig]::FormSpacer)
      $RichTextStatusBtmMidButton.Name = "RichTextStatusBtmMidButton"
      $RichTextStatusBtmMidButton.TabStop = $True
      $RichTextStatusBtmMidButton.Text = "&$($ButtonMid.ToString())"
      $RichTextStatusBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $RichTextStatusBtmMidButton.PreferredSize.Height)
      if ($ButtonMid -eq $ButtonDefault)
      {
        $RichTextStatusBtmMidButton.Select()
      }
    }
    #endregion $RichTextStatusBtmMidButton = [System.Windows.Forms.Button]::New()

    #region $RichTextStatusBtmRightButton = [System.Windows.Forms.Button]::New()
    If (($RichTextStatusButtons -eq 2) -or ($RichTextStatusButtons -eq 3))
    {
      $RichTextStatusBtmRightButton = [System.Windows.Forms.Button]::New()
      $RichTextStatusBtmPanel.Controls.Add($RichTextStatusBtmRightButton)
      $RichTextStatusBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
      $RichTextStatusBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
      $RichTextStatusBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
      $RichTextStatusBtmRightButton.DialogResult = $ButtonRight
      $RichTextStatusBtmRightButton.Enabled = $False
      $RichTextStatusBtmRightButton.Font = [MyConfig]::Font.Bold
      $RichTextStatusBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
      $RichTextStatusBtmRightButton.Location = [System.Drawing.Point]::New(($RichTextStatusBtmLeftButton.Right + $TempWidth + $TempMod + ([MyConfig]::FormSpacer * 2)), [MyConfig]::FormSpacer)
      $RichTextStatusBtmRightButton.Name = "RichTextStatusBtmRightButton"
      $RichTextStatusBtmRightButton.TabIndex = 1
      $RichTextStatusBtmRightButton.TabStop = $True
      $RichTextStatusBtmRightButton.Text = "&$($ButtonRight.ToString())"
      $RichTextStatusBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $RichTextStatusBtmRightButton.PreferredSize.Height)
      if ($ButtonRight -eq $ButtonDefault)
      {
        $RichTextStatusBtmRightButton.Select()
      }
    }
    #endregion $RichTextStatusBtmRightButton = [System.Windows.Forms.Button]::New()

    $RichTextStatusBtmPanel.ClientSize = [System.Drawing.Size]::New(($RichTextStatusTextBox.Right + [MyConfig]::FormSpacer), (($RichTextStatusBtmPanel.Controls[$RichTextStatusBtmPanel.Controls.Count - 1]).Bottom + [MyConfig]::FormSpacer))

    #endregion ******** $RichTextStatusBtmPanel Controls ********
  }

  #endregion ******** Controls for $RichTextStatus Form ********

  #endregion ******** End **** $Show-RichTextStatus **** End ********

  $PassHashTable = $PSBoundParameters.ContainsKey("HashTable")
  $DialogResult = $RichTextStatusForm.ShowDialog($PILForm)
  [RichTextStatus]::New(($DialogResult -eq $ButtonDefault), $DialogResult)

  $RichTextStatusForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Show-RichTextStatus"
}
#endregion function Show-RichTextStatus

#endregion ******** PIL Common Dialogs ********

#region ******** PIL Custom Dialogs ********

#region function Display-InitiliazePILUtility
Function Display-InitiliazePILUtility()
{
  <#
    .SYNOPSIS
      Display PILUtility Status Sample Function
    .DESCRIPTION
      Display PILUtility Status Sample Function
    .PARAMETER RichTextBox
    .PARAMETER HashTable
    .EXAMPLE
      Display-InitiliazePILUtility -RichTextBox $RichTextBox
    .EXAMPLE
      Display-InitiliazePILUtility -RichTextBox $RichTextBox -HashTable $HashTable
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory = $True)]
    [System.Windows.Forms.RichTextBox]$RichTextBox,
    [HashTable]$HashTable
  )
  Write-Verbose -Message "Enter Function Display-InitiliazePILUtility"
  
  $DisplayResult = [System.Windows.Forms.DialogResult]::OK
  $RichTextBox.Refresh()
  
  $ShowHeader = $HashTable.ShowHeader
  $ConfigFile = $HashTable.ConfigFile
  $ImportFile = $HashTable.ImportFile
  
  $RichTextBox.SelectionIndent = 10
  $RichTextBox.SelectionBullet = $False
  
  # Write KPI Event
  #Write-KPIEvent -Source "Utility" -EntryType "Information" -EventID 0 -Category 0 -Message "Some Unknown KPI Event"
  
  If ($ShowHeader)
  {
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -Text "$($RichTextBox.Parent.Parent.Text)" -TextFore ([MyConfig]::Colors.TextTitle)
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $RichTextBox.Parent.Parent.Text
    
    # Initialize StopWatch
    $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
  }
  
  
  Write-RichTextBox -RichTextBox $RichTextBox
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Validate Runtime Parameters" -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextTitle)
  $RichTextBox.SelectionIndent = 20
  $RichTextBox.SelectionBullet = $True
  
  #region ******** Validating Runtime Parameters ********
  
  # Script / Utility
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Utility" -Value ([MyConfig]::ScriptName) -ValueFore ([MyConfig]::Colors.TextGood)
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Version" -Value ([MyConfig]::ScriptVersion) -ValueFore ([MyConfig]::Colors.TextGood)
  
  # Run From/As Info
  $TmpRunFrom = Get-WmiObject -Query "Select Name, Domain, PartOfDomain From Win32_ComputerSystem"
  If ($TmpRunFrom.PartOfDomain)
  {
    $TmpRunFromText = "$($TmpRunFrom.Name).$($TmpRunFrom.Domain)"
  }
  Else
  {
    $TmpRunFromText = "$($TmpRunFrom.Name)"
  }
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Run From" -Value $TmpRunFromText
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Run As" -Value "$([Environment]::UserDomainName)\$([Environment]::UserName)"
  
  # Microsoft Entra Logon
  #Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Microsoft Entra Logon: " -Value ([MyConfig]::AADLogonInfo.Context.Account.Id)
  
  # Logon Authentication
  If ([MyConfig]::CurrentUser.AuthenticationType -eq "CloudAP")
  {
    $TmpText = "Microsoft Entra"
  }
  Else
  {
    $TmpText = "Active Directory"
  }
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Authentication" -Value "$($TmpText)"
  
  # Verify OS Architecture
  $TempRunOS = Get-WmiObject -Query "Select Caption, Version, OSArchitecture From Win32_OperatingSystem"
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Op Sys" -Value "$($TempRunOS.Caption)"
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Build" -Value "$($TempRunOS.Version)"
  
  # Verify AC Power
  $ChkBattery = (Get-WmiObject -Class Win32_Battery).BatteryStatus
  If ([String]::IsNullOrEmpty($ChkBattery) -or ($ChkBattery -eq 2))
  {
    $TmpText = "Yes"
    $TmpColor = [MyConfig]::Colors.TextGood
  }
  Else
  {
    $TmpText = "No"
    $TmpColor = [MyConfig]::Colors.TextWarn
  }
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "AC Power" -Value "$($TmpText)" -ValueFore $TmpColor
  
  # -------------------------
  # Display Passed Parameters
  # -------------------------
  $CheckParams = $Script:PSBoundParameters
  If ($CheckParams.Keys.Count)
  {
    Write-RichTextBox -RichTextBox $RichTextBox -Text "Runtime Parameters"
    ForEach ($Key In $CheckParams.Keys)
    {
      $RichTextBox.SelectionIndent = 30
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text $Key -Value $($CheckParams[$Key])
    }
  }
  
  #endregion ******** Validating Runtime Parameters ********
  
  $RichTextBox.SelectionIndent = 10
  $RichTextBox.SelectionBullet = $False
  Write-RichTextBox -RichTextBox $RichTextBox
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Begining Initialization" -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextTitle)
  
  $RichTextBox.SelectionIndent = 20
  $RichTextBox.SelectionBullet = $True
  # Get All Users Modules
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Getting All Users Modules"
  Get-ModuleList -Location "All Users" -Path ([MyRuntime]::AUModules)
  # Get Curent User Modules
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Getting Current User Modules"
  Get-ModuleList -Location "Current User" -Path ([MyRuntime]::CUModules)
  
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Modules Discovered" -Value ([MyRuntime]::Modules.Count)
  
  If (-not [String]::IsNullOrEmpty($ConfigFile))
  {
    $HashTable = @{"ShowHeader" = $False; "ConfigFile" = $ConfigFile }
    If ([System.IO.Path]::GetExtension($ConfigFile) -eq ".Json")
    {
      $DialogResult = Load-PILConfigFIleJson -RichTextBox $RichTextBox -HashTable $HashTable
    }
    Else
    {
      $DialogResult = Load-PILConfigFIleXml -RichTextBox $RichTextBox -HashTable $HashTable
    }
    If (-not [String]::IsNullOrEmpty($ImportFile))
    {
      $HashTable = @{"ShowHeader" = $False; "ImportFile" = $ImportFile }
      $DialogResult = Load-PILDataExport -RichTextBox $RichTextBox -HashTable $HashTable
    }
  }
  
  If ($ShowHeader)
  {
    $RichTextBox.SelectionIndent = 10
    $RichTextBox.SelectionBullet = $False
    Write-RichTextBox -RichTextBox $RichTextBox
    
    If ($DisplayResult -eq [System.Windows.Forms.DialogResult]::OK)
    {
      $FinalMsg = "Initialization was Successful"
      $FinalClr = [MyConfig]::Colors.TextGood
    }
    Else
    {
      $FinalMsg = "Initialization Failed"
      $FinalClr = [MyConfig]::Colors.TextBad
    }
    
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -TextFore $FinalClr -Text $FinalMsg
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Alignment "Center" -Text ($StopWatch.Elapsed.ToString())
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $FinalMsg
    $StopWatch.Stop()
  }
  
  $DisplayResult
  $DisplayResult = $Null
  
  Write-Verbose -Message "Exit Function Display-InitiliazePILUtility"
}
#endregion function Display-InitiliazePILUtility

#region function Get-ModuleList
function Get-ModuleList ()
{
  <#
    .SYNOPSIS
      Get List of Instaled Modules
    .DESCRIPTION
      Get List of Instaled Modules
    .PARAMETER Location
      Location of the Modules
    .PARAMETER Path
      Location to Search for Modules
    .PARAMETER Modules
      List to Add Modules to
    .EXAMPLE
      $Modules = [System.Collections.Generic.List[Modules]]::New()
      Get-ModuleList -Modules ([Ref]$Modules) -Location "All Users" -Path "$($ENV:ProgramFiles)\WindowsPowerShell\Modules"
      Get-ModuleList -Modules ([Ref]$Modules) -Location "Current User" -Path "$([Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell\Modules"
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $False)]
    [ValidateSet("All Users", "Current User")]
    [String]$Location = "All Users",
    [parameter(Mandatory = $True)]
    [String]$Path
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"
  
  # Get Installed Modules
  $TmpModList = Get-ChildItem -Path $Path
  foreach ($TmpModItem in $TmpModList)
  {
    # get Module Versions
    $TmpVersions = @(Get-ChildItem -Path $TmpModItem.FullName | Where-Object -FilterScript { $PSItem.Name -match "\d+\.\d+\.\d+" } | Sort-Object -Property Name -Descending | Select-Object -First 1)
    if ($TmpVersions.Count -eq 0)
    {
      if (-not [MyRuntime]::Modules.ContainsKey($TmpModItem.Name))
      {
        # Custom Module
        [MyRuntime]::Modules.Add($TmpModItem.Name, [PILModule]::New($Location, $TmpModItem.Name, "0.0.0"))
      }
    }
    else
    {
      if (-not [MyRuntime]::Modules.ContainsKey($TmpModItem.Name))
      {
        # Installed Module
        foreach ($TmpVersion in $TmpVersions)
        {
          [MyRuntime]::Modules.Add($TmpModItem.Name, [PILModule]::New($Location, $TmpModItem.Name, $TmpVersion.Name))
        }
      }
    }
  }
  
  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Get-ModuleList

#region function Load-PILConfigFIleXml
Function Load-PILConfigFIleXml()
{
  <#
    .SYNOPSIS
      Display Utility Status Sample Function
    .DESCRIPTION
      Display Utility Status Sample Function
    .PARAMETER RichTextBox
    .PARAMETER HashTable
      Passed Paramters HashTable
    .EXAMPLE
      Load-PILConfigFIleXml -RichTextBox $RichTextBox
    .EXAMPLE
      Load-PILConfigFIleXml -RichTextBox $RichTextBox -HashTable $HashTable
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory = $True)]
    [System.Windows.Forms.RichTextBox]$RichTextBox,
    [HashTable]$HashTable
  )
  Write-Verbose -Message "Enter Function Load-PILConfigFIleXml"
  
  $DisplayResult = [System.Windows.Forms.DialogResult]::OK
  $RichTextBox.Refresh()
  
  # Get Passed Values
  $ShowHeader = $HashTable.ShowHeader
  $ConfigFile = $HashTable.ConfigFile
  
  
  $RichTextBox.SelectionIndent = 10
  $RichTextBox.SelectionBullet = $False
  
  If ($ShowHeader)
  {
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -Text "$($RichTextBox.Parent.Parent.Text)" -TextFore ([MyConfig]::Colors.TextTitle)
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $RichTextBox.Parent.Parent.Text
    
    # Initialize StopWatch
    $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
  }
  
  Write-RichTextBox -RichTextBox $RichTextBox
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Processing PIL Configuration File" -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextTitle)
  $RichTextBox.SelectionIndent = 20
  $RichTextBox.SelectionBullet = $True
  
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Configuration File" -Value ([System.IO.Path]::GetFileName($ConfigFile)) -Font ([MyConfig]::Font.Bold)
  $RichTextBox.SelectionIndent = 30
  
  If ([System.IO.File]::Exists($ConfigFile))
  {
    Try
    {
      # Load Configuration
      $TmpConfig = Import-Clixml -LiteralPath $ConfigFile
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Found PIL Config File" -ValueFore ([MyConfig]::Colors.TextFore)
      
      $ChkConfig = @($TmpConfig.PSObject.Properties | Select-Object -ExpandProperty Name | Where-Object -FilterScript { $PSItem -in [MyRuntime]::ConfigProperties })
      If ($ChkCOnfig.Count -eq [MyRuntime]::ConfigProperties.Count)
      {
        # Add / Update PIL Columns
        $RichTextBox.SelectionIndent = 20
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Number of Columns" -Value ($TmpConfig.ColumnNames.Count)
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = $Null
        [MyRuntime]::UpdateTotalColumn($TmpConfig.ColumnNames.Count)
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = "Selected16Icon"
        $RichTextBox.SelectionIndent = 30
        $PILItemListListView.BeginUpdate()
        $PILItemListListView.Columns.Clear()
        $PILItemListListView.Items.Clear()
        
        [MyRuntime]::ThreadConfig.SetColumnNames($TmpConfig.ColumnNames)
        For ($I = 0; $I -lt ([MyRuntime]::CurrentColumns); $I++)
        {
          $TmpColName = [MyRuntime]::ThreadConfig.ColumnNames[$I]
          $PILItemListListView.Columns.Insert($I, $TmpColName, $TmpColName, -2)
        }
        $PILItemListListView.Columns[0].Width = -2
        $PILItemListListView.Columns.Insert([MyRuntime]::CurrentColumns, "Blank", " ", ($PILForm.Width * 4))
        $PILItemListListView.EndUpdate()
        
        # Resize Columns
        $PILItemListListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
        If ($PILItemListListView.Items.Count -gt 0)
        {
          $PILItemListListView.Columns[0].Width = -1
        }
        $PILItemListListView.Columns[([MyRuntime]::CurrentColumns)].Width = ($PILForm.Width * 4)
        $PILItemListListView.EndUpdate()
        
        # Update Thread Script
        $RichTextBox.SelectionIndent = 20
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Runspace Pool Threads" -Value ($TmpConfig.ThreadCount)
        [MyRuntime]::ThreadConfig.UpdateThreadInfo($TmpConfig.ThreadCount, $TmpConfig.ThreadScript)
        
        # Add / Update Common Modules
        $RichTextBox.SelectionIndent = 20
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Common Runspace Pool Modules" -Value ($TmpConfig.Modules.Count)
        
        # Install Modules Message
        If ([MyConfig]::IsLocalAdmin)
        {
          $TmpInallMsg = "the System Module Folder"
          $TmpScope = "AllUsers"
        }
        Else
        {
          $TmpInallMsg = "Your User Profile Module Folder"
          $TmpScope = "CurrentUser"
        }
        
        [MyRuntime]::ThreadConfig.Modules.Clear()
        :LoadMods ForEach ($Key In $TmpConfig.Modules.Keys)
        {
          $Module = $TmpConfig.Modules[$Key]
          $RichTextBox.SelectionIndent = 30
          Write-RichTextBoxValue -RichTextBox $RichTextBox -Text $Module.Name -Value $Module.Version
          $RichTextBox.SelectionIndent = 40
          
          If ([MyRuntime]::Modules.ContainsKey($Module.Name))
          {
            If ([Version]::New([MyRuntime]::Modules[$Module.Name].Version) -lt [Version]::New($Module.Version))
            {
              $Response = Get-UserResponse -Title "Incorrect Module Version" -Message "The Module $($Module.Name) Version $($Module.Version) was not Found would you like to Install it to $($TmpInallMsg)?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
              If ($Response.Success)
              {
                $ChkInstall = Install-MyModule -Name $Module.Name -Version $Module.Version -Scope $TmpScope -Install -NoImport
                If ($ChkInstall.Success)
                {
                  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation was Successful" -ValueFore ([MyConfig]::Colors.TextFore)
                }
                Else
                {
                  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "FAILED" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation Failed" -ValueFore ([MyConfig]::Colors.TextFore)
                  $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
                  Break LoadMods
                }
              }
              Else
              {
                Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "FAILED" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation Failed" -ValueFore ([MyConfig]::Colors.TextFore)
                $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
                Break LoadMods
              }
            }
          }
          Else
          {
            $Response = Get-UserResponse -Title "Module Not Instaled" -Message "The Module $($Module.Name) Version $($Module.Version) was not Found would you like to Install it to $($TmpInallMsg)?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
            If ($Response.Success)
            {
              $ChkInstall = Install-MyModule -Name $Module.Name -Version $Module.Version -Scope $TmpScope -Install -NoImport
              If ($ChkInstall.Success)
              {
                Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation was Successful" -ValueFore ([MyConfig]::Colors.TextFore)
              }
              Else
              {
                Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "FAILED" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation Failed" -ValueFore ([MyConfig]::Colors.TextFore)
                $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
                Break LoadMods
              }
            }
            Else
            {
              Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "FAILED" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation Failed" -ValueFore ([MyConfig]::Colors.TextFore)
              $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
              Break LoadMods
            }
          }
          
          # Add Module to List
          [Void][MyRuntime]::ThreadConfig.Modules.Add($Module.Name, [PILModule]::New($Module.Location, $Module.Name, $Module.Version))
        }
        
        If ($DisplayResult -eq [System.Windows.Forms.DialogResult]::OK)
        {
          # Add / Update Common Functions
          $RichTextBox.SelectionIndent = 20
          Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Common Runspace Pool Functions" -Value ($TmpConfig.Functions.Count)
          [MyRuntime]::ThreadConfig.Functions.Clear()
          $RichTextBox.SelectionIndent = 30
          ForEach ($Key In $TmpConfig.Functions.Keys)
          {
            Write-RichTextBox -RichTextBox $RichTextBox -Text $Key
            [Void][MyRuntime]::ThreadConfig.Functions.Add($Key, [PILFunction]::New($Key, $TmpConfig.Functions[$Key].ScriptBlock))
          }
          
          # Add / Update Common Variables
          $RichTextBox.SelectionIndent = 20
          Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Common Runspace Pool Variables" -Value ($TmpConfig.Variables.Count)
          [MyRuntime]::ThreadConfig.Variables.Clear()
          $RichTextBox.SelectionIndent = 30
          ForEach ($Key In $TmpConfig.Variables.Keys)
          {
            Write-RichTextBox -RichTextBox $RichTextBox -Text $Key
            [Void][MyRuntime]::ThreadConfig.Variables.Add($Key, [PILVariable]::New($Key, $TmpConfig.Variables[$Key].Value))
          }
        }
      }
      Else
      {
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "WARNING" -TextFore ([MyConfig]::Colors.TextWarn) -Value "Invalid PIL Config File" -ValueFore ([MyConfig]::Colors.TextFore)
        $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
      }
    }
    Catch
    {
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "PIL Config File was not Loaded" -ValueFore ([MyConfig]::Colors.TextFore)
      Write-RichTextBoxError -RichTextBox $RichTextBox
      $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
      
      # Reset Configuration
      $HashTable = @{"ShowHeader" = $False; "ConfigObject" = $UnknownConfig; "ConfigName" = "Unknown Configuration"}
      Load-PILConfigFIleJson -RichTextBox $RichTextBox -HashTable $HashTable | Out-Null
    }
  }
  Else
  {
    Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "PIL Config File not Found" -ValueFore ([MyConfig]::Colors.TextFore)
    $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
  }
  
  If ($ShowHeader)
  {
    $RichTextBox.SelectionIndent = 10
    $RichTextBox.SelectionBullet = $False
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Set Final Status Message
    Switch ($DisplayResult)
    {
      "OK"
      {
        $FinalMsg = "Successfully Imported PIL Configuration"
        $FinalClr = [MyConfig]::Colors.TextGood
        Break
      }
      "Cancel"
      {
        $FinalMsg = "Errors Importing PIL Configuration"
        $FinalClr = [MyConfig]::Colors.TextBad
        Break
      }
    }
    
    # Write Final Status Message
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -TextFore $FinalClr -Text $FinalMsg
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Alignment "Center" -Text ($StopWatch.Elapsed.ToString())
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $FinalMsg
    $StopWatch.Stop()
  }
    
  # Return DialogResult
  $DisplayResult
  $DisplayResult = $Null
  
  Write-Verbose -Message "Exit Function Load-PILConfigFIleXml"
}
#endregion function Load-PILConfigFIleXml

#region function Load-PILConfigFIleJson
Function Load-PILConfigFIleJson()
{
  <#
    .SYNOPSIS
      Display Utility Status Sample Function
    .DESCRIPTION
      Display Utility Status Sample Function
    .PARAMETER RichTextBox
    .PARAMETER HashTable
      Passed Paramters HashTable
    .EXAMPLE
      Load-PILConfigFIleJson -RichTextBox $RichTextBox
    .EXAMPLE
      Load-PILConfigFIleJson -RichTextBox $RichTextBox -HashTable $HashTable
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory = $True)]
    [System.Windows.Forms.RichTextBox]$RichTextBox,
    [HashTable]$HashTable
  )
  Write-Verbose -Message "Enter Function Load-PILConfigFIleJson"
  
  $DisplayResult = [System.Windows.Forms.DialogResult]::OK
  $RichTextBox.Refresh()
  
  # Get Passed Values
  $ShowHeader = $HashTable.ShowHeader
  $ConfigFile = $HashTable.ConfigFile
  $ConfigObject = $HashTable.ConfigObject
  $ConfigName = $HashTable.ConfigName
  
  
  $RichTextBox.SelectionIndent = 10
  $RichTextBox.SelectionBullet = $False
  
  If ($ShowHeader)
  {
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -Text "$($RichTextBox.Parent.Parent.Text)" -TextFore ([MyConfig]::Colors.TextTitle)
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $RichTextBox.Parent.Parent.Text
    
    # Initialize StopWatch
    $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
  }
  
  Write-RichTextBox -RichTextBox $RichTextBox
  If ([String]::IsNullOrEmpty($ConfigFile))
  {
    # Sample Config
    Write-RichTextBox -RichTextBox $RichTextBox -Text "Processing PIL Sample Configuration" -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextTitle)
    $RichTextBox.SelectionIndent = 20
    $RichTextBox.SelectionBullet = $True
    Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Sample Configuration" -Value $ConfigName -Font ([MyConfig]::Font.Bold)
    $RichTextBox.SelectionIndent = 30
    
    Try
    {
      $TmpConfig = $ConfigObject | ConvertFrom-Json
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Loading PIL Sample Configonfiguration" -ValueFore ([MyConfig]::Colors.TextFore)
    }
    Catch
    {
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "Loading PIL Sample Configonfiguration" -ValueFore ([MyConfig]::Colors.TextFore)
      $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
      Write-RichTextBoxError -RichTextBox $RichTextBox
      
    }
  }
  Else
  {
    Write-RichTextBox -RichTextBox $RichTextBox -Text "Processing PIL Configuration File" -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextTitle)
    $RichTextBox.SelectionIndent = 20
    $RichTextBox.SelectionBullet = $True
    
    Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Configuration File" -Value ([System.IO.Path]::GetFileName($ConfigFile)) -Font ([MyConfig]::Font.Bold)
    $RichTextBox.SelectionIndent = 30
    
    If ([System.IO.File]::Exists($ConfigFile))
    {
      Try
      {
        # Load Configuration
        $TmpConfig = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Loading PIL Config File" -ValueFore ([MyConfig]::Colors.TextFore)
      }
      Catch
      {
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "Loading PIL Config File" -ValueFore ([MyConfig]::Colors.TextFore)
        $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
        Write-RichTextBoxError -RichTextBox $RichTextBox
      }
    }
    Else
    {
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "WARNING" -TextFore ([MyConfig]::Colors.TextwARN) -Value "PIL Config File Not Found" -ValueFore ([MyConfig]::Colors.TextFore)
      $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
    }
  }
  
  If ($DisplayResult -eq [System.Windows.Forms.DialogResult]::OK)
  {
    Try
    {
      $ChkConfig = @($TmpConfig.PSObject.Properties | Select-Object -ExpandProperty Name | Where-Object -FilterScript { $PSItem -in [MyRuntime]::ConfigProperties })
      If ($ChkCOnfig.Count -eq [MyRuntime]::ConfigProperties.Count)
      {
        # Add / Update PIL Columns
        $RichTextBox.SelectionIndent = 20
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Number of Columns" -Value ($TmpConfig.ColumnNames.Count)
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = $Null
        [MyRuntime]::UpdateTotalColumn($TmpConfig.ColumnNames.Count)
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = "Selected16Icon"
        $RichTextBox.SelectionIndent = 30
        $PILItemListListView.BeginUpdate()
        $PILItemListListView.Columns.Clear()
        $PILItemListListView.Items.Clear()
        [MyRuntime]::ThreadConfig.SetColumnNames($TmpConfig.ColumnNames)
        [MyRuntime]::ThreadConfig.SetColumnNames($TmpConfig.ColumnNames)
        For ($I = 0; $I -lt ([MyRuntime]::CurrentColumns); $I++)
        {
          $TmpColName = [MyRuntime]::ThreadConfig.ColumnNames[$I]
          $PILItemListListView.Columns.Insert($I, $TmpColName, $TmpColName, -2)
        }
        $PILItemListListView.Columns[0].Width = -2
        $PILItemListListView.Columns.Insert([MyRuntime]::CurrentColumns, "Blank", " ", ($PILForm.Width * 4))
        $PILItemListListView.EndUpdate()
        
        # Resize Columns
        $PILItemListListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
        If ($PILItemListListView.Items.Count -gt 0)
        {
          $PILItemListListView.Columns[0].Width = -1
        }
        $PILItemListListView.Columns[([MyRuntime]::CurrentColumns)].Width = ($PILForm.Width * 4)
        $PILItemListListView.EndUpdate()
        
        # Update Thread Script
        $RichTextBox.SelectionIndent = 20
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Runspace Pool Threads" -Value ($TmpConfig.ThreadCount)
        [MyRuntime]::ThreadConfig.UpdateThreadInfo($TmpConfig.ThreadCount, $TmpConfig.ThreadScript)
        
        # Add / Update Common Modules
        $RichTextBox.SelectionIndent = 20
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Common Runspace Pool Modules" -Value ($TmpConfig.Modules.Count)
        
        # Install Modules Message
        If ([MyConfig]::IsLocalAdmin)
        {
          $TmpInallMsg = "the System Module Folder"
          $TmpScope = "AllUsers"
        }
        Else
        {
          $TmpInallMsg = "Your User Profile Module Folder"
          $TmpScope = "CurrentUser"
        }
        
        [MyRuntime]::ThreadConfig.Modules.Clear()
        $FndModules = @($TmpConfig.Modules.PSObject.Properties | Select-Object -Property Name, Value)
        :LoadMods ForEach ($Module In $FndModules)
        {
          $RichTextBox.SelectionIndent = 30
          Write-RichTextBoxValue -RichTextBox $RichTextBox -Text $Module.Name -Value $Module.Value.Version
          $RichTextBox.SelectionIndent = 40
          
          If ([MyRuntime]::Modules.ContainsKey($Module.Name))
          {
            If ([Version]::New([MyRuntime]::Modules[$Module.Name].Version) -lt [Version]::New($Module.Value.Version))
            {
              $Response = Get-UserResponse -Title "Incorrect Module Version" -Message "The Module $($Module.Name) Version $($Module.Value.Version) was not Found would you like to Install it to $($TmpInallMsg)?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
              If ($Response.Success)
              {
                $ChkInstall = Install-MyModule -Name $Module.Name -Version $Module.Value.Version -Scope $TmpScope -Install -NoImport
                If ($ChkInstall.Success)
                {
                  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation was Successful" -ValueFore ([MyConfig]::Colors.TextFore)
                }
                Else
                {
                  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "FAILED" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation Failed" -ValueFore ([MyConfig]::Colors.TextFore)
                  $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
                  Break LoadMods
                }
              }
              Else
              {
                Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "FAILED" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation Failed" -ValueFore ([MyConfig]::Colors.TextFore)
                $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
                Break LoadMods
              }
            }
          }
          Else
          {
            $Response = Get-UserResponse -Title "Module Not Instaled" -Message "The Module $($Module.Name) Version $($Module.Value.Version) was not Found would you like to Install it to $($TmpInallMsg)?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
            If ($Response.Success)
            {
              $ChkInstall = Install-MyModule -Name $Module.Name -Version $Module.Value.Version -Scope $TmpScope -Install -NoImport
              If ($ChkInstall.Success)
              {
                Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation was Successful" -ValueFore ([MyConfig]::Colors.TextFore)
              }
              Else
              {
                Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "FAILED" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation Failed" -ValueFore ([MyConfig]::Colors.TextFore)
                $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
                Break LoadMods
              }
            }
            Else
            {
              Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "FAILED" -TextFore ([MyConfig]::Colors.TextBad) -Value "Module $($Module.Name) Installation Failed" -ValueFore ([MyConfig]::Colors.TextFore)
              $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
              Break LoadMods
            }
          }
          
          # Add Module to List
          [Void][MyRuntime]::ThreadConfig.Modules.Add($Module.Name, [PILModule]::New($Module.Value.Location, $Module.Value.Name, $Module.Value.Version))
        }
        
        If ($DisplayResult -eq [System.Windows.Forms.DialogResult]::OK)
        {
          # Add / Update Common Functions
          $RichTextBox.SelectionIndent = 20
          $FndFunctions = @($TmpConfig.Functions.PSObject.Properties | Select-Object -Property Name, Value)
          Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Common Runspace Pool Functions" -Value ($FndFunctions.Count)
          [MyRuntime]::ThreadConfig.Functions.Clear()
          $RichTextBox.SelectionIndent = 30
          ForEach ($Function In $FndFunctions)
          {
            Write-RichTextBox -RichTextBox $RichTextBox -Text $Function.Name
            [Void][MyRuntime]::ThreadConfig.Functions.Add($Function.Name, [PILFunction]::New($Function.Value.Name, $Function.Value.ScriptBlock))
          }
          
          # Add / Update Common Variables
          $RichTextBox.SelectionIndent = 20
          $FndVariables = @($TmpConfig.Variables.PSObject.Properties | Select-Object -Property Name, Value)
          Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Common Runspace Pool Variables" -Value ($FndVariables.Count)
          [MyRuntime]::ThreadConfig.Variables.Clear()
          $RichTextBox.SelectionIndent = 30
          ForEach ($Variable In $FndVariables)
          {
            Write-RichTextBox -RichTextBox $RichTextBox -Text $Variable.Name
            [Void][MyRuntime]::ThreadConfig.Variables.Add($Variable.Name, [PILVariable]::New($Variable.Value.Name, $Variable.Value.Value))
          }
        }
      }
      Else
      {
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "WARNING" -TextFore ([MyConfig]::Colors.TextWarn) -Value "Invalid PIL Config File" -ValueFore ([MyConfig]::Colors.TextFore)
        $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
      }
    }
    Catch
    {
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "PIL Config File was not Loaded" -ValueFore ([MyConfig]::Colors.TextFore)
      $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
      Write-RichTextBoxError -RichTextBox $RichTextBox
      
      # Reset Configuration
      $HashTable = @{"ShowHeader" = $False; "ConfigObject" = $UnknownConfig; "ConfigName" = "Unknown Configuration"}
      Load-PILConfigFIleJson -RichTextBox $RichTextBox -HashTable $HashTable
        
    }
  }
    
  If ($ShowHeader)
  {
    $RichTextBox.SelectionIndent = 10
    $RichTextBox.SelectionBullet = $False
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Set Final Status Message
    Switch ($DisplayResult)
    {
      "OK"
      {
        $FinalMsg = "Successfully Imported PIL Configuration"
        $FinalClr = [MyConfig]::Colors.TextGood
        Break
      }
      "Cancel"
      {
        $FinalMsg = "Errors Importing PIL Configuration"
        $FinalClr = [MyConfig]::Colors.TextBad
        Break
      }
    }
    
    # Write Final Status Message
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -TextFore $FinalClr -Text $FinalMsg
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Alignment "Center" -Text ($StopWatch.Elapsed.ToString())
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $FinalMsg
    $StopWatch.Stop()
  }
  
  # Return DialogResult
  $DisplayResult
  $DisplayResult = $Null
  
  Write-Verbose -Message "Exit Function Load-PILConfigFIleJson"
}
#endregion function Load-PILConfigFIleJson

#region function Load-PILDataExport
Function Load-PILDataExport()
{
  <#
    .SYNOPSIS
      Display Utility Status Sample Function
    .DESCRIPTION
      Display Utility Status Sample Function
    .PARAMETER RichTextBox
    .PARAMETER HashTable
      Passed Paramters HashTable
    .EXAMPLE
      Load-PILDataExport -RichTextBox $RichTextBox
    .EXAMPLE
      Load-PILDataExport -RichTextBox $RichTextBox -HashTable $HashTable
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory = $True)]
    [System.Windows.Forms.RichTextBox]$RichTextBox,
    [HashTable]$HashTable
  )
  Write-Verbose -Message "Enter Function Load-PILDataExport"
  
  $DisplayResult = [System.Windows.Forms.DialogResult]::OK
  $RichTextBox.Refresh()
  
  # Get Passed Values
  $ShowHeader = $HashTable.ShowHeader
  $ImportFile = $HashTable.ImportFile
  
  $RichTextBox.SelectionIndent = 10
  $RichTextBox.SelectionBullet = $False
  
  If ($ShowHeader)
  {
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -Text "$($RichTextBox.Parent.Parent.Text)" -TextFore ([MyConfig]::Colors.TextTitle)
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $RichTextBox.Parent.Parent.Text
    
    # Initialize StopWatch
    $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
  }
  
  Write-RichTextBox -RichTextBox $RichTextBox
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Loading PIL Data Export" -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextTitle)
  $RichTextBox.SelectionIndent = 20
  $RichTextBox.SelectionBullet = $True
  
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Data Export File File" -Value ([System.IO.Path]::GetFileName($ImportFile)) -Font ([MyConfig]::Font.Bold)
  $RichTextBox.SelectionIndent = 30
  
  If ([System.IO.File]::Exists($ImportFile))
  {
    Try
    {
      # Get Column Names
      $TmpColNames = @($PILItemListListView.Columns | Select-Object -ExpandProperty Text)
      
      # Load Configuration
      $TmpExport = Import-Csv -LiteralPath $ImportFile
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Found PIL Data Export File" -ValueFore ([MyConfig]::Colors.TextFore)
      $TmpImportCols = @($TmpExport[0].PSObject.Properties | Select-Object -ExpandProperty Name)
      
      $RichTextBox.SelectionIndent = 20
      Write-RichTextBox -RichTextBox $RichTextBox -Text "Validateing PIL Export Data Columns"
      $RichTextBox.SelectionIndent = 30
      
      $ChkColumns = [ArrayList]::New(@($TmpColNames | Where-Object -FilterScript { $PSItem -in $TmpImportCols }))
      [Void]$ChkColumns.Add("FakeColumn")
      If ($ChkColumns.Count -eq $TmpColNames.Count)
      {
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Validated Data Export Columns" -ValueFore ([MyConfig]::Colors.TextFore)
        $RichTextBox.SelectionIndent = 20
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Processing Data Export" -Value "Importing $($TmpExport.Count) List Items"
        $RichTextBox.SelectionIndent = 30
        $TmpCurRowCount = $PILItemListListView.Items.Count
        $TmpDataList = $TmpExport | Select-Object -Property $ChkColumns
        $PILItemListListView.BeginUpdate()
        ForEach ($TmpDataItem In $TmpDataList)
        {
          $TmpName = $TmpDataItem."$($ChkColumns[0])"
          If (-not $PILItemListListView.Items.ContainsKey($TmpName))
          {
            $TmpDataItem.FakeColumn = ""
            ($PILItemListListView.Items.Add([System.Windows.Forms.ListViewItem]::New(@($TmpDataItem.PSObject.Properties | Select-Object -ExpandProperty Value), "StatusInfo16Icon", [MyConfig]::Colors.TextFore, [MyConfig]::Colors.TextBack, [MyConfig]::Font.Regular))).Name = $TmpName
          }
        }
        
        $PILItemListListView.EndUpdate()
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Imported $($PILItemListListView.Items.Count - $TmpCurRowCount) List Items" -ValueFore ([MyConfig]::Colors.TextFore)
      }
      Else
      {
        Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "Data Export Columns do Not Match Configured PIL Columns" -ValueFore ([MyConfig]::Colors.TextFore)
        $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
      }
    }
    Catch
    {
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "PIL Data Export File was not Loaded" -ValueFore ([MyConfig]::Colors.TextFore)
      Write-RichTextBoxError -RichTextBox $RichTextBox
      $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
    }
  }
  Else
  {
    Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "PIL Data Export File not Found" -ValueFore ([MyConfig]::Colors.TextFore)
    $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
  }
  
  If ($ShowHeader)
  {
    $RichTextBox.SelectionIndent = 10
    $RichTextBox.SelectionBullet = $False
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Set Final Status Message
    Switch ($DisplayResult)
    {
      "OK"
      {
        $FinalMsg = "Successfully Imported PIL Data Export"
        $FinalClr = [MyConfig]::Colors.TextGood
        Break
      }
      "Cancel"
      {
        $FinalMsg = "Errors Importing PIL Data Export"
        $FinalClr = [MyConfig]::Colors.TextBad
        Break
      }
    }
    
    # Write Final Status Message
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -TextFore $FinalClr -Text $FinalMsg
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Alignment "Center" -Text ($StopWatch.Elapsed.ToString())
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $FinalMsg
    $StopWatch.Stop()
  }
  
  # Return DialogResult
  $DisplayResult
  $DisplayResult = $Null
  
  Write-Verbose -Message "Exit Function Load-PILDataExport"
}
#endregion function Load-PILDataExport

#region NewPILColumn Result Class
Class NewPILColumn
{
  [Bool]$Success
  [Object]$DialogResult
  [UInt16]$Index
  [String]$Name
  
  NewPILColumn ([Object]$DialogResult)
  {
    $This.Success = ($DialogResult -eq "OK")
    $This.DialogResult = $DialogResult
  }
  
  NewPILColumn ([Object]$DialogResult, [UInt16]$Index, [String]$Name)
  {
    $This.Success = ($DialogResult -eq "OK")
    $This.DialogResult = $DialogResult
    $This.Index = $Index
    $This.Name = $Name
  }
}
#endregion NewPILColumn Result Class

#region function Add-NewPILColumn
function Add-NewPILColumn ()
{
  <#
    .SYNOPSIS
      Shows Add-NewPILColumn
    .DESCRIPTION
      Shows Add-NewPILColumn
    .PARAMETER Title
      Title of the Add-NewPILColumn Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER Items
      Items to show in the ComboBox
    .PARAMETER Sorted
      Sort ComboBox
    .PARAMETER SelectText
      The Default Selected Item when no Value is Selected
    .PARAMETER DisplayMember
      Name of the Property to Display in the CheckedListBox
    .PARAMETER ValueMember
      Name of the Property for the Value
    .PARAMETER Width
      Width of Add-NewPILColumn Dialog Window
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $Variables = @(Get-ChildItem -Path "Variable:\")
      $DialogResult = Add-NewPILColumn -Title "Combo Choice Dialog 01" -Message "Show this Sample Message Prompt to the User" -Items $Variables -DisplayMember "Name" -ValueMember "Value" -Selected ($Variables[4])
      If ($DialogResult.Success)
      {
        # Success
      }
      Else
      {
        # Failed
      }
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [String]$Message = "Status Message",
    [parameter(Mandatory = $True)]
    [Object[]]$Items = @(),
    [Switch]$Sorted,
    [String]$SelectText = "Select Value",
    [String]$DisplayMember = "Text",
    [String]$ValueMember = "Value",
    [Object]$Selected,
    [Int]$Width = 35,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel"
  )
  Write-Verbose -Message "Enter Function Add-NewPILColumn"

  #region ******** Begin **** NewPILColumn **** Begin ********

  # ************************************************
  # NewPILColumn Form
  # ************************************************
  #region $NewPILColumnForm = [System.Windows.Forms.Form]::New()
  $NewPILColumnForm = [System.Windows.Forms.Form]::New()
  $NewPILColumnForm.BackColor = [MyConfig]::Colors.Back
  $NewPILColumnForm.Font = [MyConfig]::Font.Regular
  $NewPILColumnForm.ForeColor = [MyConfig]::Colors.Fore
  $NewPILColumnForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $NewPILColumnForm.Icon = $PILForm.Icon
  $NewPILColumnForm.KeyPreview = $True
  $NewPILColumnForm.MaximizeBox = $False
  $NewPILColumnForm.MinimizeBox = $False
  $NewPILColumnForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), 0)
  $NewPILColumnForm.Name = "NewPILColumnForm"
  $NewPILColumnForm.Owner = $PILForm
  $NewPILColumnForm.ShowInTaskbar = $False
  $NewPILColumnForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $NewPILColumnForm.Text = $Title
  #endregion $NewPILColumnForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-NewPILColumnFormKeyDown ********
  function Start-NewPILColumnFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the NewPILColumn Form Control
      .DESCRIPTION
        KeyDown Event for the NewPILColumn Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-NewPILColumnFormKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$NewPILColumnForm"

    [MyConfig]::AutoExit = 0

    if ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $NewPILColumnForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$NewPILColumnForm"
  }
  #endregion ******** Function Start-NewPILColumnFormKeyDown ********
  $NewPILColumnForm.add_KeyDown({ Start-NewPILColumnFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-NewPILColumnFormShown ********
  function Start-NewPILColumnFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the NewPILColumn Form Control
      .DESCRIPTION
        Shown Event for the NewPILColumn Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-NewPILColumnFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$NewPILColumnForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Shown Event for `$NewPILColumnForm"
  }
  #endregion ******** Function Start-NewPILColumnFormShown ********
  $NewPILColumnForm.add_Shown({ Start-NewPILColumnFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for NewPILColumn Form ********

  # ************************************************
  # NewPILColumn Panel
  # ************************************************
  #region $NewPILColumnPanel = [System.Windows.Forms.Panel]::New()
  $NewPILColumnPanel = [System.Windows.Forms.Panel]::New()
  $NewPILColumnForm.Controls.Add($NewPILColumnPanel)
  $NewPILColumnPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $NewPILColumnPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $NewPILColumnPanel.Name = "NewPILColumnPanel"
  #endregion $NewPILColumnPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $NewPILColumnPanel Controls ********
  
  #region $NewPILColumnLabel = [System.Windows.Forms.Label]::New()
  $NewPILColumnLabel = [System.Windows.Forms.Label]::New()
  $NewPILColumnPanel.Controls.Add($NewPILColumnLabel)
  $NewPILColumnLabel.ForeColor = [MyConfig]::Colors.LabelFore
  $NewPILColumnLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
  $NewPILColumnLabel.Name = "NewPILColumnLabel"
  $NewPILColumnLabel.Size = [System.Drawing.Size]::New(($NewPILColumnPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
  $NewPILColumnLabel.Text = $Message
  $NewPILColumnLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  #endregion $NewPILColumnLabel = [System.Windows.Forms.Label]::New()
  
  # Returns the minimum size required to display the text
  $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($NewPILColumnLabel.Text, [MyConfig]::Font.Regular, $NewPILColumnLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
  $NewPILColumnLabel.Size = [System.Drawing.Size]::New(($NewPILColumnPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))
  
  # ************************************************
  # NewPILColumn GroupBox
  # ************************************************
  #region $NewPILColumnGroupBox = [System.Windows.Forms.GroupBox]::New()
  $NewPILColumnGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $NewPILColumnPanel.Controls.Add($NewPILColumnGroupBox)
  $NewPILColumnGroupBox.BackColor = [MyConfig]::Colors.Back
  $NewPILColumnGroupBox.Font = [MyConfig]::Font.Regular
  $NewPILColumnGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $NewPILColumnGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($NewPILColumnLabel.Bottom + ([MyConfig]::FormSpacer * 2)))
  $NewPILColumnGroupBox.Name = "NewPILColumnGroupBox"
  $NewPILColumnGroupBox.Size = [System.Drawing.Size]::New(($NewPILColumnPanel.Width - ([MyConfig]::FormSpacer * 2)), 50)
  #endregion $NewPILColumnGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $NewPILColumnGroupBox Controls ********

  #region $NewPILColumnComboBox = [System.Windows.Forms.ComboBox]::New()
  $NewPILColumnComboBox = [System.Windows.Forms.ComboBox]::New()
  $NewPILColumnGroupBox.Controls.Add($NewPILColumnComboBox)
  $NewPILColumnComboBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
  $NewPILColumnComboBox.AutoSize = $True
  $NewPILColumnComboBox.BackColor = [MyConfig]::Colors.TextBack
  $NewPILColumnComboBox.DisplayMember = $DisplayMember
  $NewPILColumnComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  $NewPILColumnComboBox.Font = [MyConfig]::Font.Regular
  $NewPILColumnComboBox.ForeColor = [MyConfig]::Colors.TextFore
  [void]$NewPILColumnComboBox.Items.Add([PSCustomObject]@{ $DisplayMember = " - $($SelectText) - "; $ValueMember = " - $($SelectText) - "})
  $NewPILColumnComboBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $NewPILColumnComboBox.Name = "NewPILColumnComboBox"
  $NewPILColumnComboBox.SelectedIndex = 0
  $NewPILColumnComboBox.Size = [System.Drawing.Size]::New(($NewPILColumnGroupBox.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), $NewPILColumnComboBox.PreferredHeight)
  $NewPILColumnComboBox.Sorted = $Sorted.IsPresent
  $NewPILColumnComboBox.TabIndex = 0
  $NewPILColumnComboBox.TabStop = $True
  $NewPILColumnComboBox.Tag = $Null
  $NewPILColumnComboBox.ValueMember = $ValueMember
  #endregion $NewPILColumnComboBox = [System.Windows.Forms.ComboBox]::New()

  $NewPILColumnComboBox.Items.AddRange($Items)
  $NewPILColumnComboBox.SelectedIndex = 0
  
  #region $NewPILColumnTextBox = [System.Windows.Forms.TextBox]::New()
  $NewPILColumnTextBox = [System.Windows.Forms.TextBox]::New()
  $NewPILColumnGroupBox.Controls.Add($NewPILColumnTextBox)
  $NewPILColumnTextBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
  $NewPILColumnTextBox.AutoSize = $True
  $NewPILColumnTextBox.BackColor = [MyConfig]::Colors.TextBack
  $NewPILColumnTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
  $NewPILColumnTextBox.Font = [MyConfig]::Font.Hint
  $NewPILColumnTextBox.ForeColor = [MyConfig]::Colors.TextHint
  $NewPILColumnTextBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($NewPILColumnComboBox.Bottom + [MyConfig]::FormSpacer))
  $NewPILColumnTextBox.MaxLength = 50
  $NewPILColumnTextBox.Multiline = $False
  $NewPILColumnTextBox.Name = "NewPILColumnTextBox"
  $NewPILColumnTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::None
  $NewPILColumnTextBox.Text = "Enter New PIL Column Name"
  $NewPILColumnTextBox.Tag = @{ "HintText" = "Enter New PIL Column Name"; "HintEnabled" = $True; "Items" = "" }
  $NewPILColumnTextBox.Size = [System.Drawing.Size]::New($NewPILColumnComboBox.Width, $NewPILColumnTextBox.PreferredHeight)
  $NewPILColumnTextBox.TabIndex = 0
  $NewPILColumnTextBox.TabStop = $True
  $NewPILColumnTextBox.WordWrap = $False
  #endregion $NewPILColumnTextBox = [System.Windows.Forms.TextBox]::New()
  
  #region ******** Function Start-NewPILColumnTextBoxGotFocus ********
  Function Start-NewPILColumnTextBoxGotFocus
  {
  <#
    .SYNOPSIS
      GotFocus Event for the NewPILColumn TextBox Control
    .DESCRIPTION
      GotFocus Event for the NewPILColumn TextBox Control
    .PARAMETER Sender
       The TextBox Control that fired the GotFocus Event
    .PARAMETER EventArg
       The Event Arguments for the TextBox GotFocus Event
    .EXAMPLE
       Start-NewPILColumnTextBoxGotFocus -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter GotFocus Event for `$NewPILColumnTextBox"
    
    [MyConfig]::AutoExit = 0
    
    # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
    If ($Sender.Tag.HintEnabled)
    {
      $Sender.Text = ""
      $Sender.Font = [MyConfig]::Font.Regular
      $Sender.ForeColor = [MyConfig]::Colors.TextFore
    }
    
    Write-Verbose -Message "Exit GotFocus Event for `$NewPILColumnTextBox"
  }
  #endregion ******** Function Start-NewPILColumnTextBoxGotFocus ********
  $NewPILColumnTextBox.add_GotFocus({ Start-NewPILColumnTextBoxGotFocus -Sender $This -EventArg $PSItem })
  
  #region ******** Function Start-NewPILColumnTextBoxKeyDown ********
  function Start-NewPILColumnTextBoxKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the NewPILColumn TextBox Control
      .DESCRIPTION
        KeyDown Event for the NewPILColumn TextBox Control
      .PARAMETER Sender
        The TextBox Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the TextBox KeyDown Event
      .EXAMPLE
        Start-NewPILColumnTextBoxKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$NewPILColumnTextBox"

    [MyConfig]::AutoExit = 0
    
    if ((-not $Sender.Multiline) -and ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Return))
    {
      $NewPILColumnBtmLeftButton.PerformClick()
    }
    
    Write-Verbose -Message "Exit KeyDown Event for `$NewPILColumnTextBox"
  }
  #endregion ******** Function Start-NewPILColumnTextBoxKeyDown ********
  $NewPILColumnTextBox.add_KeyDown({ Start-NewPILColumnTextBoxKeyDown -Sender $This -EventArg $PSItem })
  
  #region ******** Function Start-NewPILColumnTextBoxKeyPress ********
  Function Start-NewPILColumnTextBoxKeyPress
  {
    <#
      .SYNOPSIS
        KeyPress Event for the NewPILColumn TextBox Control
      .DESCRIPTION
        KeyPress Event for the NewPILColumn TextBox Control
      .PARAMETER Sender
         The TextBox Control that fired the KeyPress Event
      .PARAMETER EventArg
         The Event Arguments for the TextBox KeyPress Event
      .EXAMPLE
         Start-NewPILColumnTextBoxKeyPress -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyPress Event for `$NewPILColumnTextBox"
    
    [MyConfig]::AutoExit = 0
    
    # 1 = Ctrl-A, 3 = Ctrl-C, 8 = Backspace, 22 = Ctrl-V, 24 = Ctrl-X
    $EventArg.Handled = (($EventArg.KeyChar -notmatch ".") -and ([Int]($EventArg.KeyChar) -notin (1, 3, 8, 22, 24)))
    
    Write-Verbose -Message "Exit KeyPress Event for `$NewPILColumnTextBox"
  }
  #endregion ******** Function Start-NewPILColumnTextBoxKeyPress ********
  $NewPILColumnTextBox.add_KeyPress({Start-NewPILColumnTextBoxKeyPress -Sender $This -EventArg $PSItem})
  
  #region ******** Function Start-NewPILColumnTextBoxKeyUp ********
  Function Start-NewPILColumnTextBoxKeyUp
  {
  <#
    .SYNOPSIS
      KeyUp Event for the NewPILColumn TextBox Control
    .DESCRIPTION
      KeyUp Event for the NewPILColumn TextBox Control
    .PARAMETER Sender
       The TextBox Control that fired the KeyUp Event
    .PARAMETER EventArg
       The Event Arguments for the TextBox KeyUp Event
    .EXAMPLE
       Start-NewPILColumnTextBoxKeyUp -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyUp Event for `$NewPILColumnTextBox"
    
    [MyConfig]::AutoExit = 0
    
    # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
    $Sender.Tag.HintEnabled = ($Sender.Text.Trim().Length -eq 0)
    
    Write-Verbose -Message "Exit KeyUp Event for `$NewPILColumnTextBox"
  }
  #endregion ******** Function Start-NewPILColumnTextBoxKeyUp ********
  $NewPILColumnTextBox.add_KeyUp({ Start-NewPILColumnTextBoxKeyUp -Sender $This -EventArg $PSItem })
  
  #region ******** Function Start-NewPILColumnTextBoxLostFocus ********
  Function Start-NewPILColumnTextBoxLostFocus
  {
  <#
    .SYNOPSIS
      LostFocus Event for the NewPILColumn TextBox Control
    .DESCRIPTION
      LostFocus Event for the NewPILColumn TextBox Control
    .PARAMETER Sender
       The TextBox Control that fired the LostFocus Event
    .PARAMETER EventArg
       The Event Arguments for the TextBox LostFocus Event
    .EXAMPLE
       Start-NewPILColumnTextBoxLostFocus -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter LostFocus Event for `$NewPILColumnTextBox"
    
    [MyConfig]::AutoExit = 0
    
    # $TextBox.Tag = @{ "HintText" = ""; "HintEnabled" = $True }
    If ([String]::IsNullOrEmpty(($Sender.Text.Trim())))
    {
      $Sender.Text = $Sender.Tag.HintText
      $Sender.Tag.HintEnabled = $True
      $Sender.Font = [MyConfig]::Font.Hint
      $Sender.ForeColor = [MyConfig]::Colors.TextHint
    }
    Else
    {
      $Sender.Tag.HintEnabled = $False
      $Sender.Font = [MyConfig]::Font.Regular
      $Sender.ForeColor = [MyConfig]::Colors.TextFore
    }
    
    Write-Verbose -Message "Exit LostFocus Event for `$NewPILColumnTextBox"
  }
  #endregion ******** Function Start-NewPILColumnTextBoxLostFocus ********
  $NewPILColumnTextBox.add_LostFocus({ Start-NewPILColumnTextBoxLostFocus -Sender $This -EventArg $PSItem })
  
  $NewPILColumnGroupBox.ClientSize = [System.Drawing.Size]::New($NewPILColumnGroupBox.ClientSize.Width, ($NewPILColumnTextBox.Bottom + ([MyConfig]::FormSpacer * 2)))

  #endregion ******** $NewPILColumnGroupBox Controls ********

  $TempClientSize = [System.Drawing.Size]::New(($NewPILColumnGroupBox.Right + [MyConfig]::FormSpacer), ($NewPILColumnGroupBox.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $NewPILColumnPanel Controls ********

  # ************************************************
  # NewPILColumnBtm Panel
  # ************************************************
  #region $NewPILColumnBtmPanel = [System.Windows.Forms.Panel]::New()
  $NewPILColumnBtmPanel = [System.Windows.Forms.Panel]::New()
  $NewPILColumnForm.Controls.Add($NewPILColumnBtmPanel)
  $NewPILColumnBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $NewPILColumnBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $NewPILColumnBtmPanel.Name = "NewPILColumnBtmPanel"
  #endregion $NewPILColumnBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $NewPILColumnBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($NewPILColumnBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $NewPILColumnBtmLeftButton = [System.Windows.Forms.Button]::New()
  $NewPILColumnBtmLeftButton = [System.Windows.Forms.Button]::New()
  $NewPILColumnBtmPanel.Controls.Add($NewPILColumnBtmLeftButton)
  $NewPILColumnBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $NewPILColumnBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $NewPILColumnBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $NewPILColumnBtmLeftButton.Font = [MyConfig]::Font.Bold
  $NewPILColumnBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $NewPILColumnBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $NewPILColumnBtmLeftButton.Name = "NewPILColumnBtmLeftButton"
  $NewPILColumnBtmLeftButton.TabIndex = 1
  $NewPILColumnBtmLeftButton.TabStop = $True
  $NewPILColumnBtmLeftButton.Text = $ButtonLeft
  $NewPILColumnBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $NewPILColumnBtmLeftButton.PreferredSize.Height)
  #endregion $NewPILColumnBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-NewPILColumnBtmLeftButtonClick ********
  function Start-NewPILColumnBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the NewPILColumnBtmLeft Button Control
      .DESCRIPTION
        Click Event for the NewPILColumnBtmLeft Button Control
      .PARAMETER Sender
         The Button Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the Button Click Event
      .EXAMPLE
         Start-NewPILColumnBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$NewPILColumnBtmLeftButton"

    [MyConfig]::AutoExit = 0
    
    If (($NewPILColumnComboBox.SelectedIndex -gt 0) -and (-not $NewPILColumnTextBox.Tag.HinstEnabled) -and ($NewPILColumnTextBox.Text -notin @($NewPILColumnComboBox.Items | Select-Object -ExpandProperty Text)))
    {
      $NewPILColumnForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    else
    {
      [Void][System.Windows.Forms.MessageBox]::Show($NewPILColumnForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }

    Write-Verbose -Message "Exit Click Event for `$NewPILColumnBtmLeftButton"
  }
  #endregion ******** Function Start-NewPILColumnBtmLeftButtonClick ********
  $NewPILColumnBtmLeftButton.add_Click({ Start-NewPILColumnBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $NewPILColumnBtmMidButton = [System.Windows.Forms.Button]::New()
  $NewPILColumnBtmMidButton = [System.Windows.Forms.Button]::New()
  $NewPILColumnBtmPanel.Controls.Add($NewPILColumnBtmMidButton)
  $NewPILColumnBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $NewPILColumnBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $NewPILColumnBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $NewPILColumnBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $NewPILColumnBtmMidButton.Font = [MyConfig]::Font.Bold
  $NewPILColumnBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $NewPILColumnBtmMidButton.Location = [System.Drawing.Point]::New(($NewPILColumnBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $NewPILColumnBtmMidButton.Name = "NewPILColumnBtmMidButton"
  $NewPILColumnBtmMidButton.TabIndex = 2
  $NewPILColumnBtmMidButton.TabStop = $True
  $NewPILColumnBtmMidButton.Text = $ButtonMid
  $NewPILColumnBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $NewPILColumnBtmMidButton.PreferredSize.Height)
  #endregion $NewPILColumnBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-NewPILColumnBtmMidButtonClick ********
  function Start-NewPILColumnBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the NewPILColumnBtmMid Button Control
      .DESCRIPTION
        Click Event for the NewPILColumnBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-NewPILColumnBtmMidButtonClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By MyUserName)
  #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$NewPILColumnBtmMidButton"

    [MyConfig]::AutoExit = 0
    
    $NewPILColumnComboBox.SelectedIndex = 0
    $NewPILColumnTextBox.Text = ""
    
    Write-Verbose -Message "Exit Click Event for `$NewPILColumnBtmMidButton"
  }
  #endregion ******** Function Start-NewPILColumnBtmMidButtonClick ********
  $NewPILColumnBtmMidButton.add_Click({ Start-NewPILColumnBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $NewPILColumnBtmRightButton = [System.Windows.Forms.Button]::New()
  $NewPILColumnBtmRightButton = [System.Windows.Forms.Button]::New()
  $NewPILColumnBtmPanel.Controls.Add($NewPILColumnBtmRightButton)
  $NewPILColumnBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $NewPILColumnBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $NewPILColumnBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $NewPILColumnBtmRightButton.Font = [MyConfig]::Font.Bold
  $NewPILColumnBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $NewPILColumnBtmRightButton.Location = [System.Drawing.Point]::New(($NewPILColumnBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $NewPILColumnBtmRightButton.Name = "NewPILColumnBtmRightButton"
  $NewPILColumnBtmRightButton.TabIndex = 3
  $NewPILColumnBtmRightButton.TabStop = $True
  $NewPILColumnBtmRightButton.Text = $ButtonRight
  $NewPILColumnBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $NewPILColumnBtmRightButton.PreferredSize.Height)
  #endregion $NewPILColumnBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-NewPILColumnBtmRightButtonClick ********
  function Start-NewPILColumnBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the NewPILColumnBtmRight Button Control
      .DESCRIPTION
        Click Event for the NewPILColumnBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-NewPILColumnBtmRightButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By MyUserName)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$NewPILColumnBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here
    $NewPILColumnForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$NewPILColumnBtmRightButton"
  }
  #endregion ******** Function Start-NewPILColumnBtmRightButtonClick ********
  $NewPILColumnBtmRightButton.add_Click({ Start-NewPILColumnBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $NewPILColumnBtmPanel.ClientSize = [System.Drawing.Size]::New(($NewPILColumnBtmRightButton.Right + [MyConfig]::FormSpacer), ($NewPILColumnBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $NewPILColumnBtmPanel Controls ********

  $NewPILColumnForm.ClientSize = [System.Drawing.Size]::New($NewPILColumnForm.ClientSize.Width, ($TempClientSize.Height + $NewPILColumnBtmPanel.Height))

  #endregion ******** Controls for NewPILColumn Form ********

  #endregion ******** End **** NewPILColumn **** End ********
  
  $DialogResult = $NewPILColumnForm.ShowDialog()
  If ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
  {
    [NewPILColumn]::New($DialogResult, $NewPILColumnComboBox.SelectedItem.Value, "$($NewPILColumnTextBox.Text)".Trim())
  }
  Else
  {
    [NewPILColumn]::New($DialogResult)
  }

  $NewPILColumnForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Add-NewPILColumn"
}
#endregion function Add-NewPILColumn

#region function Start-ProcessingItems
Function Start-ProcessingItems()
{
  <#
    .SYNOPSIS
      Display Utility Status Sample Function
    .DESCRIPTION
      Display Utility Status Sample Function
    .PARAMETER RichTextBox
    .PARAMETER HashTable
      Passed Paramters HashTable
    .EXAMPLE
      Start-ProcessingItems -RichTextBox $RichTextBox
    .EXAMPLE
      Start-ProcessingItems -RichTextBox $RichTextBox -HashTable $HashTable
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory = $True)]
    [System.Windows.Forms.RichTextBox]$RichTextBox,
    [HashTable]$HashTable
  )
  Write-Verbose -Message "Enter Function Start-ProcessingItems"
  
  $DisplayResult = [System.Windows.Forms.DialogResult]::OK
  $RichTextBox.Refresh()
  
  # Get Passed Values
  $ShowHeader = $HashTable.ShowHeader
  $ListItems = $HashTable.ListItems
  
  $RichTextBox.SelectionIndent = 10
  $RichTextBox.SelectionBullet = $False
  
  If ($ShowHeader)
  {
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -Text "$($RichTextBox.Parent.Parent.Text)" -TextFore ([MyConfig]::Colors.TextTitle)
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $RichTextBox.Parent.Parent.Text
    
    # Initialize StopWatch
    $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
  }
  
  Write-RichTextBox -RichTextBox $RichTextBox
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Starting Runspace Pool Jobs" -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextTitle)
  $RichTextBox.SelectionIndent = 20
  $RichTextBox.SelectionBullet = $True
  
  $PoolParams = @{
    "Mutex" = [GUID]::NewGuid().Guid
    "HashTable" = @{
      "Pause"     = $False
      "Terminate" = $False;
      "Object"    = $Null
    }
    "MaxJobs" = [MyRuntime]::ThreadConfig.ThreadCount
  }
  
  # Add Common Modules
  $RichTextBox.SelectionIndent = 20
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Adding Common Runspace Pool Modules"
  $RichTextBox.SelectionIndent = 30
  If ([MyRuntime]::ThreadConfig.Modules.Count -gt 0)
  {
    $PoolParams.Add("Modules", @([MyRuntime]::ThreadConfig.Modules.Keys))
  }
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Found $([MyRuntime]::ThreadConfig.Modules.Count) Common Runspace Pool Modules" -ValueFore ([MyConfig]::Colors.TextFore)
  
  # Add Common Functions
  $RichTextBox.SelectionIndent = 20
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Adding Common Runspace Pool Functions"
  $RichTextBox.SelectionIndent = 30
  If ([MyRuntime]::ThreadConfig.Functions.Count -gt 0)
  {
    $TmpFunctions = @{}
    ForEach ($Function In [MyRuntime]::ThreadConfig.Functions.Values)
    {
      $TmpFunctions.Add($Function.Name, $Function.ScriptBlock)
    }
    $PoolParams.Add("Functions", $TmpFunctions)
  }
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Found $([MyRuntime]::ThreadConfig.Functions.Count) Common Runspace Pool Functions" -ValueFore ([MyConfig]::Colors.TextFore)
  
  # Add Status Icons
  $RichTextBox.SelectionIndent = 20
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Adding ListViewItem Status Icons Variables"
  $TmpVariables = @{
    "GoodIcon"  = "StatusGood16Icon"
    "BadIcon"   = "StatusBad16Icon"
    "InfoIcon"  = "StatusInfo16Icon"
    "CheckIcon" = "CheckIcon"
    "ErrorIcon" = "UncheckIcon"
    "UpIcon"    = "Up16Icon"
    "DownIcon"  = "Down16Icon"
  }
  
  $ChkPrompt = [Ordered]@{}
  [MyRuntime]::ThreadConfig.Variables.Values | Where-Object -FilterScript { $PSItem.Value -eq "*" } | ForEach-Object -Process { $ChkPrompt.Add($PSItem.Name, "")}
  If ($ChkPrompt.Count)
  {
    $DialogResult = Get-MultiTextBoxInput -Title "Prompt Variables" -Message "Enter the Runtime Values for the Indicated Variables." -OrderedItems $ChkPrompt -AllRequired -ValidChars "."
    If ($DialogResult.Success)
    {
      $TmpNames = @($DialogResult.OrderedItems.Keys)
      $Response = Get-UserResponse -Title "Save Values?" -Message "Do you want to Save These Values for this Session?" -ButtonLeft Yes -ButtonRight No -ButtonDefault No -Icon ([System.Drawing.SystemIcons]::Question)
      If (-not $Response.Success)
      {
        ForEach ($Key In $DialogResult.OrderedItems.Keys)
        {
          [MyRuntime]::ThreadConfig.Variables[$Key].Value = $DialogResult.OrderedItems[$Key]
        }
      }
      ForEach ($Key In $DialogResult.OrderedItems.Keys)
      {
        $TmpVariables.Add($Key, $DialogResult.OrderedItems[$Key])
      }
    }
  }
  
  # Add Common Variables
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Adding Custom Runspace Pool Variables"
  $RichTextBox.SelectionIndent = 30
  If ([MyRuntime]::ThreadConfig.Variables.Count -gt 0)
  {
    ForEach ($Variable In @([MyRuntime]::ThreadConfig.Variables.Values | Where-Object -FilterScript { $PSItem.Name -notin $TmpNames }))
    {
      $TmpVariables.Add($Variable.Name, $Variable.Value)
    }
  }
  
  $PoolParams.Add("Variables", $TmpVariables)
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Found $([MyRuntime]::ThreadConfig.Variables.Count) Runspace Pool Variables" -ValueFore ([MyConfig]::Colors.TextFore)
  
  # Starting Runspace Pool
  $RichTextBox.SelectionIndent = 20
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Starting Runspace Pool"
  Clear-Host
  $RichTextBox.SelectionIndent = 30
  $ChkPool = Start-MyRSPool @PoolParams -PassThru
    
  If ($ChkPool.State -eq "Opened")
  {
    Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "SUCCESS" -TextFore ([MyConfig]::Colors.TextGood) -Value "Runspace Pool ID: $($ChkPool.InstanceID)" -ValueFore ([MyConfig]::Colors.TextFore)
    $RichTextBox.SelectionIndent = 20
    Write-RichTextBox -RichTextBox $RichTextBox -Text "Starting $($ListItems.Count) Runspace Pool Jobs"
    
    ForEach ($ListItem In $ListItems)
    {
      $ListItem | Start-MyRSJob -InputParam "ListViewItem" -ScriptBlock ([ScriptBlock]::Create([MyRuntime]::ThreadConfig.ThreadScript))
    }
    
    $PILLeftProgressBar.Value = 0
    $PILRightProgressBar.Value = 0
    $PILLeftProgressBar.Maximum = $ListItems.Count
    $PILRightProgressBar.Maximum = $ListItems.Count

    # Disable Auto Close
    $PILTimer.Enabled = $False
  }
  Else
  {
    Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "Failed to Start Runspace Pool" -ValueFore ([MyConfig]::Colors.TextFore)
    $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
  }
  
  If ($ShowHeader)
  {
    $RichTextBox.SelectionIndent = 10
    $RichTextBox.SelectionBullet = $False
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Set Final Status Message
    Switch ($DisplayResult)
    {
      "OK"
      {
        $FinalMsg = "Successfully Started Runspace Pool Jobs"
        $FinalClr = [MyConfig]::Colors.TextGood
        Break
      }
      "Cancel"
      {
        $FinalMsg = "Errors Starting Runspace Pool Jobs"
        $FinalClr = [MyConfig]::Colors.TextBad
        Break
      }
    }
    
    # Write Final Status Message
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Font ([MyConfig]::Font.Title) -Alignment "Center" -TextFore $FinalClr -Text $FinalMsg
    Write-RichTextBox -RichTextBox $RichTextBox
    Write-RichTextBox -RichTextBox $RichTextBox -Alignment "Center" -Text ($StopWatch.Elapsed.ToString())
    Write-RichTextBox -RichTextBox $RichTextBox
    
    # Update Status Message
    $PILBtmStatusStrip.Items["Status"].Text = $FinalMsg
    $StopWatch.Stop()
  }
  
  # Return DialogResult
  $DisplayResult
  $DisplayResult = $Null
  
  Write-Verbose -Message "Exit Function Start-ProcessingItems"
}
#endregion function Start-ProcessingItems

#region function Monitor-RunspacePoolThreads
Function Monitor-RunspacePoolThreads ()
{
  <#
    .SYNOPSIS
      Function to do something specific
    .DESCRIPTION
      Function to do something specific
    .PARAMETER Value
      Value Command Line Parameter
    .EXAMPLE
      Monitor-RunspacePoolThreads -Value "String"
    .NOTES
      Original Function By %YourName%
      
      %Date% - Initial Release
  #>
  [CmdletBinding()]
  Param (
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"
  
  [MyConfig]::AutoExit = 0
  
  $WaitScript = {
    $JobsDone = @(Get-MyRSJob -State Completed, Failed)
    While ($JobsDone.Count -gt $PILLeftProgressBar.Value)
    {
      $PILLeftProgressBar.Increment(1)
      $PILRightProgressBar.Increment(1)
    }
    $PILLeftProgressBar.Refresh()
    $PILRightProgressBar.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    [System.Threading.Thread]::Sleep(100)
  }
  
  Try
  {
    # Wait for TYhreads to Exit  
    Get-MyRSJob | Wait-MyRSJob -Wait 0 -SciptBlock $WaitScript
    
    Get-MyRSJob | Receive-MyRSJob -AutoRemove -Force | Out-Null
    Close-MyRSPool
  }
  Catch
  {
    
  }
  
  While ($PILLeftProgressBar.Maximum -gt $PILLeftProgressBar.Value)
  {
    $PILLeftProgressBar.Increment(1)
    $PILRightProgressBar.Increment(1)
    $PILLeftProgressBar.Refresh()
    $PILRightProgressBar.Refresh()
  }
  
  # Set Processing ToolStrip Menu Items
  $PILPlayProcButton.Enabled = $True
  $PILPlayPauseButton.Enabled = $True
  $PILPlayStopButton.Enabled = $True
  $PILPlayBarPanel.Visible = $False
  
  # Re-Enable Main Menu Items
  $PILTopMenuStrip.Items["AddItems"].Enabled = $True
  $PILTopMenuStrip.Items["Configure"].Enabled = $True
  $PILTopMenuStrip.Items["ProcessItems"].Enabled = $True
  $PILTopMenuStrip.Items["ListData"].Enabled = $True
  
  # Re-Enable Right Click Menu
  $PILItemListContextMenuStrip.Enabled = $True
  
  # Enable ListView Sort
  $PILItemListListView.ListViewItemSorter.Enable = $True

  # Enable Auto Close
  $PILTimer.Enabled = ([MyConfig]::AutoExitMax -gt 0)
  
  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Monitor-RunspacePoolThreads

#region ThreadConfiguration Result Class
Class ThreadConfiguration
{
  [Bool]$Success
  [Object]$DialogResult

  ThreadConfiguration ([Bool]$Success, [Object]$DialogResult)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
  }

}
#endregion ThreadConfiguration Result Class

#region function Update-ThreadConfiguration
function Update-ThreadConfiguration ()
{
  <#
    .SYNOPSIS
      Shows ThreadConfiguration
    .DESCRIPTION
      Shows ThreadConfiguration
    .PARAMETER Title
      Title of the ThreadConfiguration Window
    .PARAMETER Width
      Width of the Statts ThreadConfiguration Window
    .PARAMETER Height
      Height of the Status ThreadConfiguration Window
    .PARAMETER ButtonLeft
      The DialogResult of the Left Button
    .PARAMETER ButtonMid
      The DialogResult of the Middle Button
    .PARAMETER ButtonRight
      The DialogResult of the Right Button
    .EXAMPLE
      $Return = ThreadConfiguration -Title $Title
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [String]$Title = "PIL Configuration - $([MyRuntime]::ConfigName)",
    [Int]$Width = 70,
    [Int]$Height = 33,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel"
  )
  Write-Verbose -Message "Enter Function Update-ThreadConfiguration"

  #region ******** Begin **** ThreadConfiguration **** Begin ********

  # ************************************************
  # ThreadConfiguration Form
  # ************************************************
  #region $ThreadConfigurationForm = [System.Windows.Forms.Form]::New()
  $ThreadConfigurationForm = [System.Windows.Forms.Form]::New()
  $ThreadConfigurationForm.BackColor = [MyConfig]::Colors.Back
  $ThreadConfigurationForm.Font = [MyConfig]::Font.Regular
  $ThreadConfigurationForm.ForeColor = [MyConfig]::Colors.Fore
  $ThreadConfigurationForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $ThreadConfigurationForm.Icon = $PILForm.Icon
  $ThreadConfigurationForm.KeyPreview = $True
  $ThreadConfigurationForm.MaximizeBox = $False
  $ThreadConfigurationForm.MinimizeBox = $False
  $ThreadConfigurationForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * $Height))
  $ThreadConfigurationForm.Name = "ThreadConfigurationForm"
  $ThreadConfigurationForm.Owner = $PILForm
  $ThreadConfigurationForm.ShowInTaskbar = $False
  $ThreadConfigurationForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $ThreadConfigurationForm.Text = $Title
  #endregion $ThreadConfigurationForm = [System.Windows.Forms.Form]::New()
  
  #region ******** Function Start-ThreadConfigurationFormKeyDown ********
  function Start-ThreadConfigurationFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the ThreadConfiguration Form Control
      .DESCRIPTION
        KeyDown Event for the ThreadConfiguration Form Control
      .PARAMETER Sender
         The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
         The Event Arguments for the Form KeyDown Event
      .EXAMPLE
         Start-ThreadConfigurationFormKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$ThreadConfigurationForm"

    [MyConfig]::AutoExit = 0

    if ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $ThreadConfigurationForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$ThreadConfigurationForm"
  }
  #endregion ******** Function Start-ThreadConfigurationFormKeyDown ********
  $ThreadConfigurationForm.add_KeyDown({ Start-ThreadConfigurationFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ThreadConfigurationFormLoad ********
  function Start-ThreadConfigurationFormLoad
  {
    <#
      .SYNOPSIS
        Load Event for the ThreadConfiguration Form Control
      .DESCRIPTION
        Load Event for the ThreadConfiguration Form Control
      .PARAMETER Sender
        The Form Control that fired the Load Event
      .PARAMETER EventArg
        The Event Arguments for the Form Load Event
      .EXAMPLE
        Start-ThreadConfigurationFormLoad -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Load Event for `$ThreadConfigurationForm"

    [MyConfig]::AutoExit = 0

    Write-Verbose -Message "Exit Load Event for `$ThreadConfigurationForm"
  }
  #endregion ******** Function Start-ThreadConfigurationFormLoad ********
  $ThreadConfigurationForm.add_Load({ Start-ThreadConfigurationFormLoad -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ThreadConfigurationFormMove ********
  function Start-ThreadConfigurationFormMove
  {
    <#
      .SYNOPSIS
        Move Event for the ThreadConfiguration Form Control
      .DESCRIPTION
        Move Event for the ThreadConfiguration Form Control
      .PARAMETER Sender
        The Form Control that fired the Move Event
      .PARAMETER EventArg
        The Event Arguments for the Form Move Event
      .EXAMPLE
        Start-ThreadConfigurationFormMove -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Move Event for `$ThreadConfigurationForm"

    [MyConfig]::AutoExit = 0

    Write-Verbose -Message "Exit Move Event for `$ThreadConfigurationForm"
  }
  #endregion ******** Function Start-ThreadConfigurationFormMove ********
  $ThreadConfigurationForm.add_Move({ Start-ThreadConfigurationFormMove -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ThreadConfigurationFormShown ********
  function Start-ThreadConfigurationFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the ThreadConfiguration Form Control
      .DESCRIPTION
        Shown Event for the ThreadConfiguration Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-ThreadConfigurationFormShown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Shown Event for `$ThreadConfigurationForm"

    [MyConfig]::AutoExit = 0
    
    $Sender.Refresh()
    Start-ThreadConfigurationBtmMidButtonClick -Sender $ThreadConfigurationBtmMidButton -EventArg "Reset"
    $ThreadConfigurationBtmLeftButton.Select()

    Write-Verbose -Message "Exit Shown Event for `$ThreadConfigurationForm"
  }
  #endregion ******** Function Start-ThreadConfigurationFormShown ********
  $ThreadConfigurationForm.add_Shown({ Start-ThreadConfigurationFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for ThreadConfiguration Form ********

  # ************************************************
  # ThreadConfiguration Panel - Fill
  # ************************************************
  #region $ThreadConfigurationPanel = [System.Windows.Forms.Panel]::New()
  $ThreadConfigurationPanel = [System.Windows.Forms.Panel]::New()
  $ThreadConfigurationForm.Controls.Add($ThreadConfigurationPanel)
  #$ThreadConfigurationPanel.BackColor = [MyConfig]::Colors.Back
  $ThreadConfigurationPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ThreadConfigurationPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $ThreadConfigurationPanel.Name = "ThreadConfigurationPanel"
  $ThreadConfigurationPanel.Padding = [System.Windows.Forms.Padding]::New([MyConfig]::FormSpacer, 0, [MyConfig]::FormSpacer, 0)
  $ThreadConfigurationPanel.Text = "ThreadConfigurationPanel"
  #endregion $ThreadConfigurationPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ThreadConfigurationPanel Controls ********
  
  $TmpValue = (($ThreadConfigurationPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 5)) / 3)
  $TmpWidth = [Math]::Floor($TmpValue)
  $TmpMod = ($TmpValue % 2)
  
  # ************************************************
  # PILTCFunctions GroupBox - Fill
  # ************************************************
  #region $PILTCFunctionsGroupBox = [System.Windows.Forms.GroupBox]::New()
  $PILTCFunctionsGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $ThreadConfigurationPanel.Controls.Add($PILTCFunctionsGroupBox)
  $PILTCFunctionsGroupBox.BackColor = [MyConfig]::Colors.Back
  $PILTCFunctionsGroupBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $PILTCFunctionsGroupBox.Font = [MyConfig]::Font.Regular
  $PILTCFunctionsGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $PILTCFunctionsGroupBox.Height = 100
  $PILTCFunctionsGroupBox.Margin = [System.Windows.Forms.Padding]::New(50, 3, 10, 50)
  $PILTCFunctionsGroupBox.Name = "PILTCFunctionsGroupBox"
  $PILTCFunctionsGroupBox.Text = "Common Functions"
  #$PILTCFunctionsGroupBox.Width = 200
  #endregion $PILTCFunctionsGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $PILTCFunctionsGroupBox Controls ********

  #region $PILTCFunctionsListBox = [System.Windows.Forms.ListBox]::New()
  $PILTCFunctionsListBox = [System.Windows.Forms.ListBox]::New()
  $PILTCFunctionsGroupBox.Controls.Add($PILTCFunctionsListBox)
  $PILTCFunctionsListBox.BackColor = [MyConfig]::Colors.TextBack
  #$PILTCFunctionsListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
  $PILTCFunctionsListBox.DisplayMember = "Name"
  $PILTCFunctionsListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $PILTCFunctionsListBox.Font = [MyConfig]::Font.Regular
  $PILTCFunctionsListBox.ForeColor = [MyConfig]::Colors.TextFore
  $PILTCFunctionsListBox.IntegralHeight = $False
  $PILTCFunctionsListBox.ItemHeight = [MyConfig]::Font.Height
  $PILTCFunctionsListBox.Name = "PILTCFunctionsListBox"
  $PILTCFunctionsListBox.Sorted = $True
  #$PILTCFunctionsListBox.TabIndex = 0
  #$PILTCFunctionsListBox.TabStop = $True
  #$PILTCFunctionsListBox.Tag = [System.Object]::New()
  $PILTCFunctionsListBox.ValueMember = "ScriptBlock"
  #endregion $PILTCFunctionsListBox = [System.Windows.Forms.ListBox]::New()
    
  #region ******** Function Start-PILTCFunctionsListBoxMouseDown ********
  function Start-PILTCFunctionsListBoxMouseDown
  {
    <#
      .SYNOPSIS
        MouseDown Event for the PILTCFunctions ListBox Control
      .DESCRIPTION
        MouseDown Event for the PILTCFunctions ListBox Control
      .PARAMETER Sender
         The TCFunctions Control that fired the MouseDown Event
      .PARAMETER EventArg
         The Event Arguments for the TCFunctions MouseDown Event
      .EXAMPLE
         Start-PILTCFunctionsListBoxMouseDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter MouseDown Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0
    
    If ($EventArg.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
      $TempIndex = $Sender.IndexFromPoint($EventArg.location)
      If ($TempIndex -gt -1)
      {
        $Sender.SelectedIndex = $TempIndex
        $PILTCFunctionsContextMenuStrip.Items["Remove"].Enabled = $True
        $PILTCFunctionsContextMenuStrip.Items["Copy"].Enabled = $True
      }
      Else
      {
        $PILTCFunctionsContextMenuStrip.Items["Remove"].Enabled = $False
        $PILTCFunctionsContextMenuStrip.Items["Copy"].Enabled = $False
      }
      $PILTCFunctionsContextMenuStrip.Items["CopyAll"].Enabled = ($Sender.Items.Count -gt 0)
      $PILTCFunctionsContextMenuStrip.Items["Clear"].Enabled = ($Sender.Items.Count -gt 0)
      $PILTCFunctionsContextMenuStrip.Show($Sender, $EventArg.Location)
    }
    
    Write-Verbose -Message "Exit MouseDown Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCFunctionsListBoxMouseDown ********
  $PILTCFunctionsListBox.add_MouseDown({Start-PILTCFunctionsListBoxMouseDown -Sender $This -EventArg $PSItem})

  #region ******** Function Start-PILTCFunctionsListBoxSelectedIndexChanged ********
  function Start-PILTCFunctionsListBoxSelectedIndexChanged
  {
    <#
      .SYNOPSIS
        SelectedIndexChanged Event for the PILTCFunctions ListBox Control
      .DESCRIPTION
        SelectedIndexChanged Event for the PILTCFunctions ListBox Control
      .PARAMETER Sender
         The TCFunctions Control that fired the SelectedIndexChanged Event
      .PARAMETER EventArg
         The Event Arguments for the TCFunctions SelectedIndexChanged Event
      .EXAMPLE
         Start-PILTCFunctionsListBoxSelectedIndexChanged -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter SelectedIndexChanged Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0

    Write-Verbose -Message "Exit SelectedIndexChanged Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCFunctionsListBoxSelectedIndexChanged ********
  $PILTCFunctionsListBox.add_SelectedIndexChanged({Start-PILTCFunctionsListBoxSelectedIndexChanged -Sender $This -EventArg $PSItem})

  # ************************************************
  # PILTCFunctions ContextMenuStrip
  # ************************************************
  #region $PILTCFunctionsContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  $PILTCFunctionsContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  #$PILTCFunctionsListBox.ContextMenuStrip = $PILTCFunctionsContextMenuStrip
  $PILTCFunctionsContextMenuStrip.BackColor = [MyConfig]::Colors.Back
  $PILTCFunctionsContextMenuStrip.Font = [MyConfig]::Font.Regular
  $PILTCFunctionsContextMenuStrip.ForeColor = [MyConfig]::Colors.Fore
  $PILTCFunctionsContextMenuStrip.ImageList = $PILSmallImageList
  $PILTCFunctionsContextMenuStrip.Name = "PILTCFunctionsContextMenuStrip"
  $PILTCFunctionsContextMenuStrip.ShowImageMargin = $True
  $PILTCFunctionsContextMenuStrip.ShowItemToolTips = $True
  #$PILTCFunctionsContextMenuStrip.TabIndex = 0
  #$PILTCFunctionsContextMenuStrip.TabStop = $False
  #$PILTCFunctionsContextMenuStrip.Tag = [System.Object]::New()
  #endregion $PILTCFunctionsContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()

  #region ******** Function Start-PILTCFunctionsContextMenuStripOpening ********
  function Start-PILTCFunctionsContextMenuStripOpening
  {
    <#
      .SYNOPSIS
        Opening Event for the PILTCFunctions ContextMenuStrip Control
      .DESCRIPTION
        Opening Event for the PILTCFunctions ContextMenuStrip Control
      .PARAMETER Sender
         The TCFunctions Control that fired the Opening Event
      .PARAMETER EventArg
         The Event Arguments for the TCFunctions Opening Event
      .EXAMPLE
         Start-PILTCFunctionsContextMenuStripOpening -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ContextMenuStrip]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Opening Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0

    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

    Write-Verbose -Message "Exit Opening Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCFunctionsContextMenuStripOpening ********
  $PILTCFunctionsContextMenuStrip.add_Opening({Start-PILTCFunctionsContextMenuStripOpening -Sender $This -EventArg $PSItem})
    
  #region ******** Function Start-PILTCFunctionsContextMenuStripItemClick ********
  Function Start-PILTCFunctionsContextMenuStripItemClick
  {
    <#
      .SYNOPSIS
        ItemClicked Event for the PILTCFunctions ToolStripItem Control
      .DESCRIPTION
        ItemClicked Event for the PILTCFunctions ToolStripItem Control
      .PARAMETER Sender
         The TCFunctions Control that fired the ItemClicked Event
      .PARAMETER EventArg
         The Event Arguments for the TCFunctions ItemClicked Event
      .EXAMPLE
         Start-PILTCFunctionsContextMenuStripItemClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ToolStripItem]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter ItemClicked Event for $($MyInvocation.MyCommand)"
    
    [MyConfig]::AutoExit = 0
    
    # Play Sound
    ##[System.Console]::Beep(2000, 10)
    
    Switch ($Sender.Name)
    {
      "Add"
      {
        $PILOpenFileDialog.FileName = ""
        $PILOpenFileDialog.Filter = "PowerShell Scripts|*.PS1|All Files (*.*)|*.*"
        $PILOpenFileDialog.FilterIndex = 1
        $PILOpenFileDialog.Multiselect = $False
        $PILOpenFileDialog.Title = "Load PIL Thread Script"
        $PILOpenFileDialog.Tag = $Null
        $Response = $PILOpenFileDialog.ShowDialog()
        If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
        {
          $TmpFunctions = Get-Content -Path $PILOpenFileDialog.FileName -Raw
          If ($TmpFunctions.Length -gt 0)
          {
            $AST = [System.Management.Automation.Language.Parser]::ParseInput($TmpFunctions, [ref]$Null, [ref]$Null)
            $Functions = @($AST.FindAll({ Param ($Node) (($Node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and (-not ($node.Parent -is [System.Management.Automation.Language.FunctionMemberAst]))) }, $True))
            If ($Functions.Count -gt 0)
            {
              ForEach ($Function In $Functions)
              {
                [Void]$PILTCFunctionsListBox.Items.Add([PILFunction]::New($Function.Name, $Function.Body.GetScriptBlock()))
              }
            }
            # Save File Path
            $PILOpenFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILOpenFileDialog.FileName)
          }
          
        }
        Break
      }
      "Remove"
      {
        $PILTCFunctionsListBox.Items.RemoveAt($PILTCFunctionsListBox.SelectedIndex)
        Break
      }
      "Copy"
      {
        $StringBuilder = [System.Text.StringBuilder]::New("`r`n#region **** Function $($PILTCFunctionsListBox.SelectedItem.Name) ****`r`n")
        [Void]$StringBuilder.AppendLine("function $($PILTCFunctionsListBox.SelectedItem.Name) ()")
        [Void]$StringBuilder.Append("{")
        [Void]$StringBuilder.Append(($PILTCFunctionsListBox.SelectedItem.ScriptBlock.ToString()))
        [Void]$StringBuilder.AppendLine("}")
        [Void]$StringBuilder.AppendLine("#endregion **** Function $($PILTCFunctionsListBox.SelectedItem.Name) ****`r`n")
        [System.Windows.Forms.Clipboard]::SetText($StringBuilder.ToString())
        Break
      }
      "CopyAll"
      {
        $StringBuilder = [System.Text.StringBuilder]::New()
        ForEach ($Item In $PILTCFunctionsListBox.Items)
        {
          [Void]$StringBuilder.AppendLine("#region **** Function $($Item.Name) ****")
          [Void]$StringBuilder.AppendLine("function $($Item.Name) ()")
          [Void]$StringBuilder.Append("{")
          [Void]$StringBuilder.Append(($Item.ScriptBlock.ToString()))
          [Void]$StringBuilder.AppendLine("}")
          [Void]$StringBuilder.AppendLine("#endregion **** Function $($Item.Name) ****`r`n")
        }
        [System.Windows.Forms.Clipboard]::SetText($StringBuilder.ToString())
        Break
      }
      "Clear"
      {
        $PILTCFunctionsListBox.Items.Clear()
        Break
      }
    }
    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"
    
    Write-Verbose -Message "Exit ItemClicked Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCFunctionsContextMenuStripItemClick ********
  
  (New-MenuItem -Menu $PILTCFunctionsContextMenuStrip -Text "Add Functions" -Name "Add" -Tag "Add" -DisplayStyle "ImageAndText" -ImageKey "Add16Icon" -PassThru).add_Click({Start-PILTCFunctionsContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $PILTCFunctionsContextMenuStrip -Text "Remove Function" -Name "Remove" -Tag "Remove" -DisplayStyle "ImageAndText" -ImageKey "Delete16Icon" -PassThru).add_Click({Start-PILTCFunctionsContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  New-MenuSeparator -Menu $PILTCFunctionsContextMenuStrip
  (New-MenuItem -Menu $PILTCFunctionsContextMenuStrip -Text "Copy Function" -Name "Copy" -Tag "Copy" -DisplayStyle "ImageAndText" -ImageKey "Copy16Icon" -PassThru).add_Click({Start-PILTCFunctionsContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $PILTCFunctionsContextMenuStrip -Text "Copy All Functions" -Name "CopyAll" -Tag "CopyAll" -DisplayStyle "ImageAndText" -ImageKey "Copy16Icon" -PassThru).add_Click({Start-PILTCFunctionsContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  New-MenuSeparator -Menu $PILTCFunctionsContextMenuStrip
  (New-MenuItem -Menu $PILTCFunctionsContextMenuStrip -Text "Delete All" -Name "Clear" -Tag "Clear" -DisplayStyle "ImageAndText" -ImageKey "Trash16Icon" -PassThru).add_Click({ Start-PILTCFunctionsContextMenuStripItemClick -Sender $This -EventArg $PSItem})

  $PILTCFunctionsGroupBox.ClientSize = [System.Drawing.Size]::New($PILTCFunctionsGroupBox.ClientSize.Width, ([MyConfig]::Font.Height * 10))

  #endregion ******** $PILTCFunctionsGroupBox Controls ********

  # ************************************************
  # PILTCVariables GroupBox - Right
  # ************************************************
  #region $PILTCVariablesGroupBox = [System.Windows.Forms.GroupBox]::New()
  $PILTCVariablesGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $ThreadConfigurationPanel.Controls.Add($PILTCVariablesGroupBox)
  $PILTCVariablesGroupBox.BackColor = [MyConfig]::Colors.Back
  $PILTCVariablesGroupBox.Dock = [System.Windows.Forms.DockStyle]::Right
  $PILTCVariablesGroupBox.Font = [MyConfig]::Font.Regular
  $PILTCVariablesGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $PILTCVariablesGroupBox.Name = "PILTCVariablesGroupBox"
  $PILTCVariablesGroupBox.Text = "Common Variables"
  $PILTCVariablesGroupBox.Width = $TmpWidth
  #endregion $PILTCVariablesGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $PILTCVariablesGroupBox Controls ********

  #region $PILTCVariablesListBox = [System.Windows.Forms.ListBox]::New()
  $PILTCVariablesListBox = [System.Windows.Forms.ListBox]::New()
  $PILTCVariablesGroupBox.Controls.Add($PILTCVariablesListBox)
  $PILTCVariablesListBox.BackColor = [MyConfig]::Colors.TextBack
  #$PILTCVariablesListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
  $PILTCVariablesListBox.DisplayMember = "Name"
  $PILTCVariablesListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $PILTCVariablesListBox.Font = [MyConfig]::Font.Regular
  $PILTCVariablesListBox.ForeColor = [MyConfig]::Colors.TextFore
  $PILTCVariablesListBox.IntegralHeight = $False
  $PILTCVariablesListBox.ItemHeight = [MyConfig]::Font.Height
  $PILTCVariablesListBox.Name = "PILTCVariablesListBox"
  $PILTCVariablesListBox.Sorted = $True
  #$PILTCVariablesListBox.TabIndex = 0
  #$PILTCVariablesListBox.TabStop = $True
  #$PILTCVariablesListBox.Tag = [System.Object]::New()
  $PILTCVariablesListBox.ValueMember = "Value"
  #endregion $PILTCVariablesListBox = [System.Windows.Forms.ListBox]::New()
  
  #region ******** Function Start-PILTCVariablesListBoxDoubleClick ********
  function Start-PILTCVariablesListBoxDoubleClick
  {
    <#
      .SYNOPSIS
        DoubleClick Event for the PILTCVariables ListBox Control
      .DESCRIPTION
        DoubleClick Event for the PILTCVariables ListBox Control
      .PARAMETER Sender
         The TCVariables Control that fired the DoubleClick Event
      .PARAMETER EventArg
         The Event Arguments for the TCVariables DoubleClick Event
      .EXAMPLE
         Start-PILTCVariablesListBoxDoubleClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter DoubleClick Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0
    
    $TempIndex = $Sender.IndexFromPoint($EventArg.location)
    If ($TempIndex -gt -1)
    {
      $Sender.SelectedIndex = $TempIndex
      $OrderedItems = [Ordered]@{ "Variable Name" = $PILTCVariablesListBox.SelectedItem.Name; "Variable Value" = $PILTCVariablesListBox.SelectedItem.Value }
      $DialogResult = Get-MultiTextBoxInput -Title "Edit Variable" -Message "Update Common Variable Name and Value" -OrderedItems $OrderedItems -AllRequired -ValidChars "."
      If ($DialogResult.Success)
      {
        $PILTCVariablesListBox.Items.RemoveAt($PILTCVariablesListBox.SelectedIndex)
        [Void]$PILTCVariablesListBox.Items.Add([PILVariable]::New($DialogResult.OrderedItems["Variable Name"], $DialogResult.OrderedItems["Variable Value"]))
      }
    }
    
    Write-Verbose -Message "Exit DoubleClick Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCVariablesListBoxDoubleClick ********
  $PILTCVariablesListBox.add_DoubleClick({Start-PILTCVariablesListBoxDoubleClick -Sender $This -EventArg $PSItem})

  #region ******** Function Start-PILTCVariablesListBoxMouseDown ********
  function Start-PILTCVariablesListBoxMouseDown
  {
    <#
      .SYNOPSIS
        MouseDown Event for the PILTCVariables ListBox Control
      .DESCRIPTION
        MouseDown Event for the PILTCVariables ListBox Control
      .PARAMETER Sender
         The TCVariables Control that fired the MouseDown Event
      .PARAMETER EventArg
         The Event Arguments for the TCVariables MouseDown Event
      .EXAMPLE
         Start-PILTCVariablesListBoxMouseDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter MouseDown Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0
    
    If ($EventArg.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
      $TempIndex = $Sender.IndexFromPoint($EventArg.location)
      If ($TempIndex -gt -1)
      {
        $Sender.SelectedIndex = $TempIndex
        $PILTCVariablesContextMenuStrip.Items["Remove"].Enabled = $True
        $PILTCVariablesContextMenuStrip.Items["Edit"].Enabled = $True
      }
      Else
      {
        $PILTCVariablesContextMenuStrip.Items["Remove"].Enabled = $False
        $PILTCVariablesContextMenuStrip.Items["Edit"].Enabled = $False
      }
      $PILTCVariablesContextMenuStrip.Items["Clear"].Enabled = ($Sender.Items.Count -gt 0)
      $PILTCVariablesContextMenuStrip.Show($Sender, $EventArg.Location)
    }
    
    Write-Verbose -Message "Exit MouseDown Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCVariablesListBoxMouseDown ********
  $PILTCVariablesListBox.add_MouseDown({Start-PILTCVariablesListBoxMouseDown -Sender $This -EventArg $PSItem})

  #region ******** Function Start-PILTCVariablesListBoxSelectedIndexChanged ********
  function Start-PILTCVariablesListBoxSelectedIndexChanged
  {
    <#
      .SYNOPSIS
        SelectedIndexChanged Event for the PILTCVariables ListBox Control
      .DESCRIPTION
        SelectedIndexChanged Event for the PILTCVariables ListBox Control
      .PARAMETER Sender
         The TCVariables Control that fired the SelectedIndexChanged Event
      .PARAMETER EventArg
         The Event Arguments for the TCVariables SelectedIndexChanged Event
      .EXAMPLE
         Start-PILTCVariablesListBoxSelectedIndexChanged -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter SelectedIndexChanged Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0


    Write-Verbose -Message "Exit SelectedIndexChanged Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCVariablesListBoxSelectedIndexChanged ********
  $PILTCVariablesListBox.add_SelectedIndexChanged({Start-PILTCVariablesListBoxSelectedIndexChanged -Sender $This -EventArg $PSItem})

  # ************************************************
  # PILTCVariables ContextMenuStrip
  # ************************************************
  #region $PILTCVariablesContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  $PILTCVariablesContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  #$PILTCVariablesListBox.ContextMenuStrip = $PILTCVariablesContextMenuStrip
  $PILTCVariablesContextMenuStrip.BackColor = [MyConfig]::Colors.Back
  $PILTCVariablesContextMenuStrip.Font = [MyConfig]::Font.Regular
  $PILTCVariablesContextMenuStrip.ForeColor = [MyConfig]::Colors.Fore
  $PILTCVariablesContextMenuStrip.ImageList = $PILSmallImageList
  $PILTCVariablesContextMenuStrip.Name = "PILTCVariablesContextMenuStrip"
  $PILTCVariablesContextMenuStrip.ShowImageMargin = $True
  $PILTCVariablesContextMenuStrip.ShowItemToolTips = $True
  #$PILTCVariablesContextMenuStrip.TabIndex = 0
  #$PILTCVariablesContextMenuStrip.TabStop = $False
  #$PILTCVariablesContextMenuStrip.Tag = [System.Object]::New()
  #endregion $PILTCVariablesContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  
  #region ******** Function Start-PILTCVariablesContextMenuStripOpening ********
  function Start-PILTCVariablesContextMenuStripOpening
  {
    <#
      .SYNOPSIS
        Opening Event for the PILTCVariables ContextMenuStrip Control
      .DESCRIPTION
        Opening Event for the PILTCVariables ContextMenuStrip Control
      .PARAMETER Sender
         The TCVariables Control that fired the Opening Event
      .PARAMETER EventArg
         The Event Arguments for the TCVariables Opening Event
      .EXAMPLE
         Start-PILTCVariablesContextMenuStripOpening -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ContextMenuStrip]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Opening Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0

    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

    Write-Verbose -Message "Exit Opening Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCVariablesContextMenuStripOpening ********
  $PILTCVariablesContextMenuStrip.add_Opening({Start-PILTCVariablesContextMenuStripOpening -Sender $This -EventArg $PSItem})
  
  #region ******** Function Start-PILTCVariablesContextMenuStripItemClick ********
  Function Start-PILTCVariablesContextMenuStripItemClick
  {
    <#
      .SYNOPSIS
        ItemClicked Event for the PILTCVariables ToolStripItem Control
      .DESCRIPTION
        ItemClicked Event for the PILTCVariables ToolStripItem Control
      .PARAMETER Sender
         The TCVariables Control that fired the ItemClicked Event
      .PARAMETER EventArg
         The Event Arguments for the TCVariables ItemClicked Event
      .EXAMPLE
         Start-PILTCVariablesContextMenuStripItemClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ToolStripItem]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter ItemClicked Event for $($MyInvocation.MyCommand)"
    
    [MyConfig]::AutoExit = 0
    
    # Play Sound
    ##[System.Console]::Beep(2000, 10)
    
    Switch ($Sender.Name)
    {
      "Add"
      {
        $OrderedItems = [Ordered]@{ "Variable Name"= ""; "Variable Value" = "" }
        $DialogResult = Get-MultiTextBoxInput -Title "Add Variable" -Message "Add New Common Variable Name and Value" -OrderedItems $OrderedItems -AllRequired -ValidChars "."
        If ($DialogResult.Success)
        {
          [Void]$PILTCVariablesListBox.Items.Add([PILVariable]::New($DialogResult.OrderedItems["Variable Name"], $DialogResult.OrderedItems["Variable Value"]))
        }
        Break
      }
      "Edit"
      {
        $OrderedItems = [Ordered]@{ "Variable Name" = $PILTCVariablesListBox.SelectedItem.Name; "Variable Value" = $PILTCVariablesListBox.SelectedItem.Value }
        $DialogResult = Get-MultiTextBoxInput -Title "Edit Variable" -Message "Update Common Variable Name and Value" -OrderedItems $OrderedItems -AllRequired -ValidChars "."
        If ($DialogResult.Success)
        {
          $PILTCVariablesListBox.Items.RemoveAt($PILTCVariablesListBox.SelectedIndex)
          [Void]$PILTCVariablesListBox.Items.Add([PILVariable]::New($DialogResult.OrderedItems["Variable Name"], $DialogResult.OrderedItems["Variable Value"]))
        }
        Break
      }
      "Remove"
      {
        $PILTCVariablesListBox.Items.RemoveAt($PILTCVariablesListBox.SelectedIndex)
        Break
      }
      "Clear"
      {
        $PILTCVariablesListBox.Items.Clear()
        Break
      }
    }
    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"
    
    Write-Verbose -Message "Exit ItemClicked Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCVariablesContextMenuStripItemClick ********
  
  (New-MenuItem -Menu $PILTCVariablesContextMenuStrip -Text "Add Variable" -Name "Add" -Tag "Add" -DisplayStyle "ImageAndText" -ImageKey "Add16Icon" -PassThru).add_Click({Start-PILTCVariablesContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $PILTCVariablesContextMenuStrip -Text "Edit Variable" -Name "Edit" -Tag "Edit" -DisplayStyle "ImageAndText" -ImageKey "Edit16Icon" -PassThru).add_Click({Start-PILTCVariablesContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $PILTCVariablesContextMenuStrip -Text "Remove Variable" -Name "Remove" -Tag "Remove" -DisplayStyle "ImageAndText" -ImageKey "Delete16Icon" -PassThru).add_Click({Start-PILTCVariablesContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  New-MenuSeparator -Menu $PILTCVariablesContextMenuStrip
  (New-MenuItem -Menu $PILTCVariablesContextMenuStrip -Text "Delete All" -Name "Clear" -Tag "Clear" -DisplayStyle "ImageAndText" -ImageKey "Trash16Icon" -PassThru).add_Click({Start-PILTCVariablesContextMenuStripItemClick -Sender $This -EventArg $PSItem})

  #endregion ******** $PILTCVariablesGroupBox Controls ********
  
  # ************************************************
  # PILTCModules GroupBox - Left
  # ************************************************
  #region $PILTCModulesGroupBox = [System.Windows.Forms.GroupBox]::New()
  $PILTCModulesGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $ThreadConfigurationPanel.Controls.Add($PILTCModulesGroupBox)
  $PILTCModulesGroupBox.BackColor = [MyConfig]::Colors.Back
  $PILTCModulesGroupBox.Dock = [System.Windows.Forms.DockStyle]::Left
  $PILTCModulesGroupBox.Font = [MyConfig]::Font.Regular
  $PILTCModulesGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $PILTCModulesGroupBox.Name = "PILTCModulesGroupBox"
  $PILTCModulesGroupBox.Text = "Common Modules"
  $PILTCModulesGroupBox.Width = $TmpWidth
  #endregion $PILTCModulesGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $PILTCModulesGroupBox Controls ********

  #region $PILTCModulesListBox = [System.Windows.Forms.ListBox]::New()
  $PILTCModulesListBox = [System.Windows.Forms.ListBox]::New()
  $PILTCModulesGroupBox.Controls.Add($PILTCModulesListBox)
  $PILTCModulesListBox.BackColor = [MyConfig]::Colors.TextBack
  #$PILTCModulesListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
  $PILTCModulesListBox.DisplayMember = "Name"
  $PILTCModulesListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $PILTCModulesListBox.Font = [MyConfig]::Font.Regular
  $PILTCModulesListBox.ForeColor = [MyConfig]::Colors.TextFore
  $PILTCModulesListBox.IntegralHeight = $False
  $PILTCModulesListBox.ItemHeight = [MyConfig]::Font.Height
  $PILTCModulesListBox.Name = "PILTCModulesListBox"
  $PILTCModulesListBox.Sorted = $False
  #$PILTCModulesListBox.TabIndex = 0
  #$PILTCModulesListBox.TabStop = $True
  #$PILTCModulesListBox.Tag = [System.Object]::New()
  $PILTCModulesListBox.ValueMember = "Version"
  #endregion $PILTCModulesListBox = [System.Windows.Forms.ListBox]::New()
  
  # Add Current Modules
  If ([MyRuntime]::ThreadConfig.Modules.Count -gt 0)
  {
    #$PILTCModulesListBox.Items.AddRange(@([MyRuntime]::ThreadConfig.Modules.Values))
  }
  
  #region ******** Function Start-PILTCModulesListBoxMouseDown ********
  function Start-PILTCModulesListBoxMouseDown
  {
    <#
      .SYNOPSIS
        MouseDown Event for the PILTCModules ListBox Control
      .DESCRIPTION
        MouseDown Event for the PILTCModules ListBox Control
      .PARAMETER Sender
         The TCModules Control that fired the MouseDown Event
      .PARAMETER EventArg
         The Event Arguments for the TCModules MouseDown Event
      .EXAMPLE
         Start-PILTCModulesListBoxMouseDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter MouseDown Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0
    
    If ($EventArg.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
      $TempIndex = $Sender.IndexFromPoint($EventArg.location)
      If ($TempIndex -gt -1)
      {
        $Sender.SelectedIndex = $TempIndex
        $PILTCModulesContextMenuStrip.Items["Remove"].Enabled = $True
        $PILTCModulesContextMenuStrip.Items["Up"].Enabled = ($TempIndex -gt 0)
        $PILTCModulesContextMenuStrip.Items["Down"].Enabled = ($TempIndex -lt ($Sender.Items.Count - 1))
      }
      Else
      {
        $PILTCModulesContextMenuStrip.Items["Remove"].Enabled = $False
        $PILTCModulesContextMenuStrip.Items["Up"].Enabled = $False
        $PILTCModulesContextMenuStrip.Items["Down"].Enabled = $False
      }
      $PILTCModulesContextMenuStrip.Items["Clear"].Enabled = ($Sender.Items.Count -gt 0)
      $PILTCModulesContextMenuStrip.Show($Sender, $EventArg.Location)
    }

    Write-Verbose -Message "Exit MouseDown Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCModulesListBoxMouseDown ********
  $PILTCModulesListBox.add_MouseDown({Start-PILTCModulesListBoxMouseDown -Sender $This -EventArg $PSItem})

  #region ******** Function Start-PILTCModulesListBoxSelectedIndexChanged ********
  function Start-PILTCModulesListBoxSelectedIndexChanged
  {
    <#
      .SYNOPSIS
        SelectedIndexChanged Event for the PILTCModules ListBox Control
      .DESCRIPTION
        SelectedIndexChanged Event for the PILTCModules ListBox Control
      .PARAMETER Sender
         The TCModules Control that fired the SelectedIndexChanged Event
      .PARAMETER EventArg
         The Event Arguments for the TCModules SelectedIndexChanged Event
      .EXAMPLE
         Start-PILTCModulesListBoxSelectedIndexChanged -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter SelectedIndexChanged Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0


    Write-Verbose -Message "Exit SelectedIndexChanged Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCModulesListBoxSelectedIndexChanged ********
  $PILTCModulesListBox.add_SelectedIndexChanged({Start-PILTCModulesListBoxSelectedIndexChanged -Sender $This -EventArg $PSItem})

  # ************************************************
  # PILTCModules ContextMenuStrip
  # ************************************************
  #region $PILTCModulesContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  $PILTCModulesContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  #$PILTCModulesListBox.ContextMenuStrip = $PILTCModulesContextMenuStrip
  $PILTCModulesContextMenuStrip.BackColor = [MyConfig]::Colors.Back
  $PILTCModulesContextMenuStrip.Font = [MyConfig]::Font.Regular
  $PILTCModulesContextMenuStrip.ForeColor = [MyConfig]::Colors.Fore
  $PILTCModulesContextMenuStrip.ImageList = $PILSmallImageList
  $PILTCModulesContextMenuStrip.Name = "PILTCModulesContextMenuStrip"
  $PILTCModulesContextMenuStrip.ShowImageMargin = $True
  $PILTCModulesContextMenuStrip.ShowItemToolTips = $True
  #$PILTCModulesContextMenuStrip.TabIndex = 0
  #$PILTCModulesContextMenuStrip.TabStop = $False
  #$PILTCModulesContextMenuStrip.Tag = [System.Object]::New()
  #endregion $PILTCModulesContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()

 
  #region ******** Function Start-PILTCModulesContextMenuStripOpening ********
  function Start-PILTCModulesContextMenuStripOpening
  {
    <#
      .SYNOPSIS
        Opening Event for the PILTCModules ContextMenuStrip Control
      .DESCRIPTION
        Opening Event for the PILTCModules ContextMenuStrip Control
      .PARAMETER Sender
         The TCModules Control that fired the Opening Event
      .PARAMETER EventArg
         The Event Arguments for the TCModules Opening Event
      .EXAMPLE
         Start-PILTCModulesContextMenuStripOpening -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ContextMenuStrip]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Opening Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0

    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

    Write-Verbose -Message "Exit Opening Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCModulesContextMenuStripOpening ********
  $PILTCModulesContextMenuStrip.add_Opening({Start-PILTCModulesContextMenuStripOpening -Sender $This -EventArg $PSItem})

  #region ******** Function Start-PILTCModulesContextMenuStripItemClick ********
  function Start-PILTCModulesContextMenuStripItemClick
  {
    <#
      .SYNOPSIS
        Click Event for the PILTCModules ContextMenuStripItem Control
      .DESCRIPTION
        Click Event for the PILTCModules ContextMenuStripItem Control
      .PARAMETER Sender
         The TCModules Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the TCModules Click Event
      .EXAMPLE
         Start-PILTCModulesContextMenuStripItemClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ToolStripItem]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0
    
    # Play Sound
    ##[System.Console]::Beep(2000, 10)
    
    Switch ($Sender.Name)
    {
      "Add"
      {
        $TmpCurMods = @($PILTCModulesListBox.Items | Select-Object -ExpandProperty Name)
        $TmpNewMods = @([MyRuntime]::Modules.Values | Where-Object { $PSItem.Name -notin $TmpCurMods } | Sort-Object -Property Location, Name)
        If ($TmpNewMods.Count -eq 0)
        {
          $Response = Get-UserResponse -Title "No More Modules" -Message "No New Modules are Avaible for to Add to the PIL Thread Configuration." -ButtonMid OK -ButtonDefault OK -Icon ([System.Drawing.SystemIcons]::Information)
        }
        Else
        {
          $DialogResult = Get-ListViewOption -Title "Select Modules" -Message "Select The Modules to Add to the PIL Thread Configuration." -Items $TmpNewMods -Property "Location", "Name", "Version" -Resize -Multi
          If ($DialogResult.Success)
          {
            $PILTCModulesListBox.Items.AddRange($DialogResult.item)
          }
        }
        Break
      }
      "Remove"
      {
        $PILTCModulesListBox.Items.RemoveAt($PILTCModulesListBox.SelectedIndex)
        Break
      }
      "Clear"
      {
        $PILTCModulesListBox.Items.Clear()
        Break
      }
      "Up"
      {
        $TmpItem = $PILTCModulesListBox.SelectedItem
        $TmpIndex = $PILTCModulesListBox.SelectedIndex
        $PILTCModulesListBox.Items.RemoveAt($PILTCModulesListBox.SelectedIndex)
        $PILTCModulesListBox.Items.Insert(($TmpIndex - 1), $TmpItem)
        Break
      }
      "Down"
      {
        $TmpItem = $PILTCModulesListBox.SelectedItem
        $TmpIndex = $PILTCModulesListBox.SelectedIndex
        $PILTCModulesListBox.Items.RemoveAt($PILTCModulesListBox.SelectedIndex)
        $PILTCModulesListBox.Items.Insert(($TmpIndex + 1), $TmpItem)
        Break
      }
    }
    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"
    

    Write-Verbose -Message "Exit Click Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCModulesContextMenuStripItemClick ********

  (New-MenuItem -Menu $PILTCModulesContextMenuStrip -Text "Add Module" -Name "Add" -Tag "Add" -DisplayStyle "ImageAndText" -ImageKey "Add16Icon" -PassThru).add_Click({Start-PILTCModulesContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $PILTCModulesContextMenuStrip -Text "Remove Module" -Name "Remove" -Tag "Remove" -DisplayStyle "ImageAndText" -ImageKey "Delete16Icon" -PassThru).add_Click({Start-PILTCModulesContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  New-MenuSeparator -Menu $PILTCModulesContextMenuStrip
  (New-MenuItem -Menu $PILTCModulesContextMenuStrip -Text "Delete All" -Name "Clear" -Tag "Clear" -DisplayStyle "ImageAndText" -ImageKey "Trash16Icon" -PassThru).add_Click({ Start-PILTCModulesContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  New-MenuSeparator -Menu $PILTCModulesContextMenuStrip
  (New-MenuItem -Menu $PILTCModulesContextMenuStrip -Text "Move Up" -Name "Up" -Tag "Up" -DisplayStyle "ImageAndText" -ImageKey "Up16Icon" -PassThru).add_Click({Start-PILTCModulesContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $PILTCModulesContextMenuStrip -Text "Move Down" -Name "Down" -Tag "Down" -DisplayStyle "ImageAndText" -ImageKey "Down16Icon" -PassThru).add_Click({Start-PILTCModulesContextMenuStripItemClick -Sender $This -EventArg $PSItem})

  #endregion ******** $PILTCModulesGroupBox Controls ********
  
  # ************************************************
  # PILTCScript GroupBox - Bottom
  # ************************************************
  #region $PILTCScriptGroupBox = [System.Windows.Forms.GroupBox]::New()
  $PILTCScriptGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $ThreadConfigurationPanel.Controls.Add($PILTCScriptGroupBox)
  $PILTCScriptGroupBox.BackColor = [MyConfig]::Colors.Back
  $PILTCScriptGroupBox.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $PILTCScriptGroupBox.Font = [MyConfig]::Font.Regular
  $PILTCScriptGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  #$PILTCScriptGroupBox.Height = 100
  #$PILTCScriptGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $PILTCScriptGroupBox.Name = "PILTCScriptGroupBox"
  $PILTCScriptGroupBox.Text = "Thread Script"
  $PILTCScriptGroupBox.Size = [System.Drawing.Size]::New($TmpWidth, $TmpWidth)
  #$PILTCScriptGroupBox.Width = 200
  #endregion $PILTCScriptGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $PILTCScriptGroupBox Controls ********

  #region $PILTCScriptTextBox = [System.Windows.Forms.TextBox]::New()
  $PILTCScriptTextBox = [System.Windows.Forms.TextBox]::New()
  $PILTCScriptGroupBox.Controls.Add($PILTCScriptTextBox)
  $PILTCScriptTextBox.BackColor = [MyConfig]::Colors.TextBack
  $PILTCScriptTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
  $PILTCScriptTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $PILTCScriptTextBox.Font = [MyConfig]::Font.Regular
  $PILTCScriptTextBox.Font = [MyConfig]::Font.Regular
  $PILTCScriptTextBox.ForeColor = [MyConfig]::Colors.TextFore
  $PILTCScriptTextBox.Multiline = $True
  $PILTCScriptTextBox.Name = "PILTCScriptTextBox"
  $PILTCScriptTextBox.ReadOnly = $True
  $PILTCScriptTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
  $PILTCScriptTextBox.ShortcutsEnabled = $False
  #$PILTCScriptTextBox.TabIndex = 0
  #$PILTCScriptTextBox.TabStop = $True
  #$PILTCScriptTextBox.Tag = @{ "HintText" = "Double Click to Load Thread Script."; "HintEnabled" = $True }
  $PILTCScriptTextBox.Text = $Null
  $PILTCScriptTextBox.WordWrap = $False
  #endregion $PILTCScriptTextBox = [System.Windows.Forms.TextBox]::New()

  #region ******** Function Start-PILTCScriptTextBoxMouseDown ********
  function Start-PILTCScriptTextBoxMouseDown
  {
    <#
      .SYNOPSIS
        MouseDown Event for the PILTCScript TextBox Control
      .DESCRIPTION
        MouseDown Event for the PILTCScript TextBox Control
      .PARAMETER Sender
         The TCScript Control that fired the MouseDown Event
      .PARAMETER EventArg
         The Event Arguments for the TCScript MouseDown Event
      .EXAMPLE
         Start-PILTCScriptTextBoxMouseDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TextBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter MouseDown Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0
    
    If ($EventArg.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
      $PILTCScriptContextMenuStrip.Items["Copy"].Enabled = ($PILTCScriptTextBox.Text.Length -gt 0)
      $PILTCScriptContextMenuStrip.Items["Clear"].Enabled = ($PILTCScriptTextBox.Text.Length -gt 0)
      $PILTCScriptContextMenuStrip.Show($Sender, $EventArg.Location)
    }
    
    Write-Verbose -Message "Exit MouseDown Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCScriptTextBoxMouseDown ********
  $PILTCScriptTextBox.add_MouseDown({Start-PILTCScriptTextBoxMouseDown -Sender $This -EventArg $PSItem})
  
  # ************************************************
  # PILTCScript ContextMenuStrip
  # ************************************************
  #region $PILTCScriptContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  $PILTCScriptContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  #$PILTCScriptListBox.ContextMenuStrip = $PILTCScriptContextMenuStrip
  $PILTCScriptContextMenuStrip.BackColor = [MyConfig]::Colors.Back
  $PILTCScriptContextMenuStrip.Font = [MyConfig]::Font.Regular
  $PILTCScriptContextMenuStrip.ForeColor = [MyConfig]::Colors.Fore
  $PILTCScriptContextMenuStrip.ImageList = $PILSmallImageList
  $PILTCScriptContextMenuStrip.Name = "PILTCScriptContextMenuStrip"
  $PILTCScriptContextMenuStrip.ShowImageMargin = $True
  $PILTCScriptContextMenuStrip.ShowItemToolTips = $True
  #$PILTCScriptContextMenuStrip.TabIndex = 0
  #$PILTCScriptContextMenuStrip.TabStop = $False
  #$PILTCScriptContextMenuStrip.Tag = [System.Object]::New()
  #endregion $PILTCScriptContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()

  #region ******** Function Start-PILTCScriptContextMenuStripOpening ********
  function Start-PILTCScriptContextMenuStripOpening
  {
    <#
      .SYNOPSIS
        Opening Event for the PILTCScript ContextMenuStrip Control
      .DESCRIPTION
        Opening Event for the PILTCScript ContextMenuStrip Control
      .PARAMETER Sender
         The TCScript Control that fired the Opening Event
      .PARAMETER EventArg
         The Event Arguments for the TCScript Opening Event
      .EXAMPLE
         Start-PILTCScriptContextMenuStripOpening -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ContextMenuStrip]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Opening Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0

    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

    Write-Verbose -Message "Exit Opening Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCScriptContextMenuStripOpening ********
  $PILTCScriptContextMenuStrip.add_Opening({Start-PILTCScriptContextMenuStripOpening -Sender $This -EventArg $PSItem})

  #region ******** Function Start-PILTCScriptContextMenuStripItemClick ********
  function Start-PILTCScriptContextMenuStripItemClick
  {
    <#
      .SYNOPSIS
        Click Event for the PILTCScript ContextMenuStripItem Control
      .DESCRIPTION
        Click Event for the PILTCScript ContextMenuStripItem Control
      .PARAMETER Sender
         The TCScript Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the TCScript Click Event
      .EXAMPLE
         Start-PILTCScriptContextMenuStripItemClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ToolStripItem]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0
    
    # Play Sound
    #[System.Console]::Beep(2000, 10)
    
    Switch ($Sender.Name)
    {
      "Add"
      {
        $PILOpenFileDialog.FileName = ""
        $PILOpenFileDialog.Filter = "PowerShell Scripts|*.PS1|All Files (*.*)|*.*"
        $PILOpenFileDialog.FilterIndex = 1
        $PILOpenFileDialog.Multiselect = $False
        $PILOpenFileDialog.Title = "Load PIL Thread Script"
        $PILOpenFileDialog.Tag = $Null
        $Response = $PILOpenFileDialog.ShowDialog()
        If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
        {
          $PILTCScriptTextBox.Text = Get-Content -Path $PILOpenFileDialog.FileName -Raw
          $PILTCScriptTextBox.AppendText("`r`n`r`n")
          $PILTCScriptTextBox.Focus()
          $PILTCScriptTextBox.SelectionStart = 0
          $PILTCScriptTextBox.SelectionLength = 0
          $PILTCScriptTextBox.ScrollToCaret()
          $PILOpenFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILOpenFileDialog.FileName)
        }
        Break
      }
      "Copy"
      {
        $PILTCScriptTextBox.SelectAll()
        $PILTCScriptTextBox.Copy()
        $PILTCScriptTextBox.DeselectAll()
        Break
      }
      "Clear"
      {
        $PILTCScriptTextBox.Text = $Null
        Break
      }
    }
    
    Write-Verbose -Message "Exit Click Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCScriptContextMenuStripItemClick ********

  (New-MenuItem -Menu $PILTCScriptContextMenuStrip -Text "Load Script" -Name "Add" -Tag "Add" -DisplayStyle "ImageAndText" -ImageKey "Add16Icon" -PassThru).add_Click({Start-PILTCScriptContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $PILTCScriptContextMenuStrip -Text "Copy Script" -Name "Copy" -Tag "Copy" -DisplayStyle "ImageAndText" -ImageKey "Copy16Icon" -PassThru).add_Click({Start-PILTCScriptContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $PILTCScriptContextMenuStrip -Text "Delete Script" -Name "Clear" -Tag "Clear" -DisplayStyle "ImageAndText" -ImageKey "Trash16Icon" -PassThru).add_Click({Start-PILTCScriptContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  
  $PILTCScriptGroupBox.ClientSize = [System.Drawing.Size]::New(($PILTCScriptGroupBox.ClientSize.Width), (([MyConfig]::Font.Height * 10) + [MyConfig]::FormSpacer))

  #endregion ******** $PILTCScriptGroupBox Controls ********
  
  # ************************************************
  # PILTCThreads GroupBox - Bottom
  # ************************************************
  #region $PILTCThreadsGroupBox = [System.Windows.Forms.GroupBox]::New()
  $PILTCThreadsGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $ThreadConfigurationPanel.Controls.Add($PILTCThreadsGroupBox)
  $PILTCThreadsGroupBox.BackColor = [MyConfig]::Colors.Back
  $PILTCThreadsGroupBox.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $PILTCThreadsGroupBox.Font = [MyConfig]::Font.Regular
  $PILTCThreadsGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  #$PILTCThreadsGroupBox.Height = 100
  #$PILTCThreadsGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $PILTCThreadsGroupBox.Name = "PILTCThreadsGroupBox"
  $PILTCThreadsGroupBox.Text = "Maximun Number of Processing Threads"
  $PILTCThreadsGroupBox.Size = [System.Drawing.Size]::New($TmpWidth, $TmpWidth)
  #$PILTCThreadsGroupBox.Width = 200
  #endregion $PILTCThreadsGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $PILTCThreadsGroupBox Controls ********

  #region $PILTCThreadsTrackBar = [System.Windows.Forms.TrackBar]::New()
  $PILTCThreadsTrackBar = [System.Windows.Forms.TrackBar]::New()
  $PILTCThreadsGroupBox.Controls.Add($PILTCThreadsTrackBar)
  $PILTCThreadsTrackBar.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $PILTCThreadsTrackBar.AutoSize = $False
  $PILTCThreadsTrackBar.BackColor = [MyConfig]::Colors.Back
  $PILTCThreadsTrackBar.Dock = [System.Windows.Forms.DockStyle]::Fill
  $PILTCThreadsTrackBar.Font = [MyConfig]::Font.Regular
  $PILTCThreadsTrackBar.ForeColor = [MyConfig]::Colors.Fore
  $PILTCThreadsTrackBar.Height = (3 * [MyConfig]::Font.Height)
  $PILTCThreadsTrackBar.LargeChange = 2
  $PILTCThreadsTrackBar.Maximum = 16
  $PILTCThreadsTrackBar.MinimumSize = [System.Drawing.Size]::New(0, $PILTCThreadsTrackBar.PreferredSize.Height)
  $PILTCThreadsTrackBar.Minimum = 1
  $PILTCThreadsTrackBar.Name = "PILTCThreadsTrackBar"
  $PILTCThreadsTrackBar.Orientation = [System.Windows.Forms.Orientation]::Horizontal
  $PILTCThreadsTrackBar.SmallChange = 1
  #$PILTCThreadsTrackBar.TabIndex = 0
  #$PILTCThreadsTrackBar.TabStop = $True
  #$PILTCThreadsTrackBar.Tag = [System.Object]::New()
  $PILTCThreadsTrackBar.TickFrequency = 1
  $PILTCThreadsTrackBar.TickStyle = [System.Windows.Forms.TickStyle]::Both
  $PILTCThreadsTrackBar.Value = [MyRuntime]::ThreadConfig.ThreadCount
  #endregion $PILTCThreadsTrackBar = [System.Windows.Forms.TrackBar]::New()
  $PILToolTip.SetToolTip($PILTCThreadsTrackBar, $PILTCThreadsTrackBar.Value)
  
  #region ******** Function Start-PILTCThreadsTrackBarValueChanged ********
  function Start-PILTCThreadsTrackBarValueChanged
  {
    <#
      .SYNOPSIS
        ValueChanged Event for the PILTCModules TrackBar Control
      .DESCRIPTION
        ValueChanged Event for the PILTCModules TrackBar Control
      .PARAMETER Sender
         The TCModules Control that fired the ValueChanged Event
      .PARAMETER EventArg
         The Event Arguments for the TCModules ValueChanged Event
      .EXAMPLE
         Start-PILTCThreadsTrackBarValueChanged -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.TrackBar]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter ValueChanged Event for $($MyInvocation.MyCommand)"

    [MyConfig]::AutoExit = 0
    
    # Play Sound
    #[System.Console]::Beep(2000, 10)
    
    $PILToolTip.SetToolTip($PILTCThreadsTrackBar, $PILTCThreadsTrackBar.Value)

    Write-Verbose -Message "Exit ValueChanged Event for $($MyInvocation.MyCommand)"
  }
  #endregion ******** Function Start-PILTCThreadsTrackBarValueChanged ********
  $PILTCThreadsTrackBar.add_ValueChanged({Start-PILTCThreadsTrackBarValueChanged -Sender $This -EventArg $PSItem})
  
  $PILTCThreadsGroupBox.ClientSize = [System.Drawing.Size]::New(($PILTCThreadsGroupBox.ClientSize.Width), ($PILTCThreadsTrackBar.PreferredSize.Height + $PILTCThreadsTrackBar.Top + [MyConfig]::FormSpacer))
  
  #endregion ******** $PILTCThreadsGroupBox Controls ********

  #endregion ******** $ThreadConfigurationPanel Controls ********

  # ************************************************
  # ThreadConfigurationBtm Panel - Bottom
  # ************************************************
  #region $ThreadConfigurationBtmPanel = [System.Windows.Forms.Panel]::New()
  $ThreadConfigurationBtmPanel = [System.Windows.Forms.Panel]::New()
  $ThreadConfigurationForm.Controls.Add($ThreadConfigurationBtmPanel)
  #$ThreadConfigurationBtmPanel.BackColor = [MyConfig]::Colors.Back
  $ThreadConfigurationBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ThreadConfigurationBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $ThreadConfigurationBtmPanel.Name = "ThreadConfigurationBtmPanel"
  $ThreadConfigurationBtmPanel.Text = "ThreadConfigurationBtmPanel"
  #endregion $ThreadConfigurationBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ThreadConfigurationBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($ThreadConfigurationBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $ThreadConfigurationBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ThreadConfigurationBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ThreadConfigurationBtmPanel.Controls.Add($ThreadConfigurationBtmLeftButton)
  $ThreadConfigurationBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $ThreadConfigurationBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ThreadConfigurationBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ThreadConfigurationBtmLeftButton.Font = [MyConfig]::Font.Bold
  $ThreadConfigurationBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ThreadConfigurationBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $ThreadConfigurationBtmLeftButton.Name = "ThreadConfigurationBtmLeftButton"
  #$ThreadConfigurationBtmLeftButton.TabIndex = 0
  #$ThreadConfigurationBtmLeftButton.TabStop = $True
  #$ThreadConfigurationBtmLeftButton.Tag = [System.Object]::New()
  $ThreadConfigurationBtmLeftButton.Text = $ButtonLeft
  $ThreadConfigurationBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $ThreadConfigurationBtmLeftButton.PreferredSize.Height)
  #endregion $ThreadConfigurationBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ThreadConfigurationBtmLeftButtonClick ********
  function Start-ThreadConfigurationBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ThreadConfigurationBtmLeft Button Control
      .DESCRIPTION
        Click Event for the ThreadConfigurationBtmLeft Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ThreadConfigurationBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ThreadConfigurationBtmLeftButton"

    [MyConfig]::AutoExit = 0
    
    If ([String]::IsNullOrEmpty($PILTCScriptTextBox.Text))
    {
      $Result = Get-UserResponse -Title "Missing or Invalid PIL Configuration" -Icon ([System.Drawing.SystemIcons]::Error) -Message "No Thread Script was Selected."
    }
    Else
    {
      [MyRuntime]::ThreadConfig.UpdateThreadInfo($PILTCThreadsTrackBar.Value, $PILTCScriptTextBox.Text)
      
      [MyRuntime]::ThreadConfig.Modules.Clear()
      If ($PILTCModulesListBox.Items.Count -gt 0)
      {
        $PILTCModulesListBox.Items | Select-Object -Property * -Unique | ForEach-Object -Process { [Void][MyRuntime]::ThreadConfig.Modules.Add($PSItem.Name, $PSItem) }
      }
      
      [MyRuntime]::ThreadConfig.Functions.Clear()
      If ($PILTCFunctionsListBox.Items.Count -gt 0)
      {
        $PILTCFunctionsListBox.Items | Select-Object -Property * -Unique | ForEach-Object -Process { [Void][MyRuntime]::ThreadConfig.Functions.Add($PSItem.Name, $PSItem) }
      }
      
      [MyRuntime]::ThreadConfig.Variables.Clear()
      If ($PILTCVariablesListBox.Items.Count -gt 0)
      {
        $PILTCVariablesListBox.Items | Select-Object -Property * -Unique | ForEach-Object -Process { [Void][MyRuntime]::ThreadConfig.Variables.Add($PSItem.Name, $PSItem) }
      }
      
      $ThreadConfigurationForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }

    Write-Verbose -Message "Exit Click Event for `$ThreadConfigurationBtmLeftButton"
  }
  #endregion ******** Function Start-ThreadConfigurationBtmLeftButtonClick ********
  $ThreadConfigurationBtmLeftButton.add_Click({ Start-ThreadConfigurationBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $ThreadConfigurationBtmMidButton = [System.Windows.Forms.Button]::New()
  $ThreadConfigurationBtmMidButton = [System.Windows.Forms.Button]::New()
  $ThreadConfigurationBtmPanel.Controls.Add($ThreadConfigurationBtmMidButton)
  #$ThreadConfigurationBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $ThreadConfigurationBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $ThreadConfigurationBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ThreadConfigurationBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ThreadConfigurationBtmMidButton.Font = [MyConfig]::Font.Bold
  $ThreadConfigurationBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ThreadConfigurationBtmMidButton.Location = [System.Drawing.Point]::New(($ThreadConfigurationBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ThreadConfigurationBtmMidButton.Name = "ThreadConfigurationBtmMidButton"
  #$ThreadConfigurationBtmMidButton.TabIndex = 0
  #$ThreadConfigurationBtmMidButton.TabStop = $True
  #$ThreadConfigurationBtmMidButton.Tag = [System.Object]::New()
  $ThreadConfigurationBtmMidButton.Text = $ButtonMid
  $ThreadConfigurationBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $ThreadConfigurationBtmMidButton.PreferredSize.Height)
  #endregion $ThreadConfigurationBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ThreadConfigurationBtmMidButtonClick ********
  function Start-ThreadConfigurationBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ThreadConfigurationBtmMid Button Control
      .DESCRIPTION
        Click Event for the ThreadConfigurationBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ThreadConfigurationBtmMidButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ThreadConfigurationBtmMidButton"

    [MyConfig]::AutoExit = 0
    
    # Add Current Modules
    $PILTCModulesListBox.Items.Clear()
    If ([MyRuntime]::ThreadConfig.Modules.Count -gt 0)
    {
      $PILTCModulesListBox.Items.AddRange(@([MyRuntime]::ThreadConfig.Modules.Values))
    }
    
    # Add Current Functions
    $PILTCFunctionsListBox.Items.Clear()
    If ([MyRuntime]::ThreadConfig.Functions.Count -gt 0)
    {
      $PILTCFunctionsListBox.Items.AddRange(@([MyRuntime]::ThreadConfig.Functions.Values))
    }
    
    # Add Current Variables
    $PILTCVariablesListBox.Items.Clear()
    If ([MyRuntime]::ThreadConfig.Variables.Count -gt 0)
    {
      $PILTCVariablesListBox.Items.AddRange(@([MyRuntime]::ThreadConfig.Variables.Values))
    }
    
    # Thread Config
    $PILTCScriptTextBox.Text = [MyRuntime]::ThreadConfig.ThreadScript
    $PILTCThreadsTrackBar.Value = [MyRuntime]::ThreadConfig.ThreadCount
    
    Write-Verbose -Message "Exit Click Event for `$ThreadConfigurationBtmMidButton"
  }
  #endregion ******** Function Start-ThreadConfigurationBtmMidButtonClick ********
  $ThreadConfigurationBtmMidButton.add_Click({ Start-ThreadConfigurationBtmMidButtonClick -Sender $This -EventArg $PSItem })
  
  #region $ThreadConfigurationBtmRightButton = [System.Windows.Forms.Button]::New()
  $ThreadConfigurationBtmRightButton = [System.Windows.Forms.Button]::New()
  $ThreadConfigurationBtmPanel.Controls.Add($ThreadConfigurationBtmRightButton)
  $ThreadConfigurationBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $ThreadConfigurationBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ThreadConfigurationBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ThreadConfigurationBtmRightButton.Font = [MyConfig]::Font.Bold
  $ThreadConfigurationBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ThreadConfigurationBtmRightButton.Location = [System.Drawing.Point]::New(($ThreadConfigurationBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ThreadConfigurationBtmRightButton.Name = "ThreadConfigurationBtmRightButton"
  #$ThreadConfigurationBtmRightButton.TabIndex = 0
  #$ThreadConfigurationBtmRightButton.TabStop = $True
  #$ThreadConfigurationBtmRightButton.Tag = [System.Object]::New()
  $ThreadConfigurationBtmRightButton.Text = $ButtonRight
  $ThreadConfigurationBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $ThreadConfigurationBtmRightButton.PreferredSize.Height)
  #endregion $ThreadConfigurationBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ThreadConfigurationBtmRightButtonClick ********
  function Start-ThreadConfigurationBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ThreadConfigurationBtmRight Button Control
      .DESCRIPTION
        Click Event for the ThreadConfigurationBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ThreadConfigurationBtmRightButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By kensw)
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ThreadConfigurationBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here

    $ThreadConfigurationForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$ThreadConfigurationBtmRightButton"
  }
  #endregion ******** Function Start-ThreadConfigurationBtmRightButtonClick ********
  $ThreadConfigurationBtmRightButton.add_Click({ Start-ThreadConfigurationBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $ThreadConfigurationBtmPanel.ClientSize = [System.Drawing.Size]::New(($ThreadConfigurationBtmRightButton.Right + [MyConfig]::FormSpacer), ($ThreadConfigurationBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ThreadConfigurationBtmPanel Controls ********

  #endregion ******** Controls for ThreadConfiguration Form ********

  #endregion ******** End **** ThreadConfiguration **** End ********
  
  # Display Config Form
  $DialogResult = $ThreadConfigurationForm.ShowDialog($PILForm)
  
  # Return Succes / Cancel Status
  [ThreadConfiguration]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult)
  
  $ThreadConfigurationForm.Dispose()
  
  Write-Verbose -Message "Exit Function Update-ThreadConfiguration"
}
#endregion function Update-ThreadConfiguration

#region function Reset-PILConfiguration
Function Reset-PILConfiguration ()
{
  <#
    .SYNOPSIS
      Function to do something specific
    .DESCRIPTION
      Function to do something specific
    .PARAMETER Value
      Value Command Line Parameter
    .EXAMPLE
      Reset-PILConfiguration -Value "String"
    .NOTES
      Original Function By %YourName%
      
      %Date% - Initial Release
  #>
  [CmdletBinding()]
  Param (
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"
  
  [MyConfig]::AutoExit = 0
  
  [MyRuntime]::UpdateTotalColumn([MyRuntime]::StartColumns)
  $PILItemListListView.BeginUpdate()
  $PILItemListListView.Columns.Clear()
  $PILItemListListView.Items.Clear()
  For ($I = 0; $I -lt ([MyRuntime]::CurrentColumns); $I++)
  {
    $TmpColName = [MyRuntime]::ThreadConfig.ColumnNames[$I]
    $PILItemListListView.Columns.Insert($I, $TmpColName, $TmpColName, -2)
  }
  $PILItemListListView.Columns[0].Width = -2
  $PILItemListListView.Columns.Insert([MyRuntime]::CurrentColumns, "Blank", " ", ($PILForm.Width * 4))
  $PILItemListListView.EndUpdate()
  [MyRuntime]::ConfigName = "Unknown Configuration"
  
  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Reset-PILConfiguration

#endregion ******** PIL Custom Dialogs ********

#region ******** Begin **** PIL **** Begin ********

$PILFormComponents = [System.ComponentModel.Container]::New()

#region $PILOpenFileDialog = [System.Windows.Forms.OpenFileDialog]::New()
$PILOpenFileDialog = [System.Windows.Forms.OpenFileDialog]::New()
$PILOpenFileDialog.InitialDirectory = [MyConfig]::ScriptRoot
$PILOpenFileDialog.Multiselect = $False
$PILOpenFileDialog.ShowHelp = $False
$PILOpenFileDialog.ValidateNames = $True
#endregion $PILOpenFileDialog = [System.Windows.Forms.OpenFileDialog]::New()

#region $PILSaveFileDialog = [System.Windows.Forms.SaveFileDialog]::New()
$PILSaveFileDialog = [System.Windows.Forms.SaveFileDialog]::New()
$PILSaveFileDialog.AddExtension = $True
$PILSaveFileDialog.CheckFileExists = $False
$PILSaveFileDialog.CreatePrompt = $False
$PILSaveFileDialog.InitialDirectory = [MyConfig]::ScriptRoot
$PILSaveFileDialog.OverwritePrompt = $True
$PILSaveFileDialog.ShowHelp = $False
$PILSaveFileDialog.ValidateNames = $True
#endregion $PILSaveFileDialog = [System.Windows.Forms.SaveFileDialog]::New()

#region $PILToolTip = [System.Windows.Forms.ToolTip]::New()
$PILToolTip = [System.Windows.Forms.ToolTip]::New($PILFormComponents)
#$PILToolTip.Active = $True
#$PILToolTip.AutomaticDelay = 500
#$PILToolTip.AutoPopDelay = 5000
$PILToolTip.BackColor = [MyConfig]::Colors.Back
$PILToolTip.ForeColor = [MyConfig]::Colors.Fore
#$PILToolTip.InitialDelay = 500
#$PILToolTip.IsBalloon = $False
#$PILToolTip.OwnerDraw = $False
#$PILToolTip.ReshowDelay = 100
#$PILToolTip.ShowAlways = $False
#$PILToolTip.StripAmpersands = $False
#$PILToolTip.Tag = [System.Object]::New()
#$PILToolTip.ToolTipIcon = [System.Windows.Forms.ToolTipIcon]::None
#$PILToolTip.ToolTipTitle = "$([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
#$PILToolTip.UseAnimation = $True
#$PILToolTip.UseFading = $True
#endregion $PILToolTip = [System.Windows.Forms.ToolTip]::New()


# ************************************************
# PILSmall ImageList
# ************************************************
#region $PILSmallImageList = [System.Windows.Forms.ImageList]::New()
$PILSmallImageList = [System.Windows.Forms.ImageList]::New($PILFormComponents)
$PILSmallImageList.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
$PILSmallImageList.ImageSize = [System.Drawing.Size]::New(16, 16)
#endregion $PILSmallImageList = [System.Windows.Forms.ImageList]::New()

#region ******** PIL Small Image Icons ********

#region ******** $PILFormIcon ********
# Icons for Forms are 16x16
$PILFormIcon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJdlMgGlbDlhwXpHu6NrOIUAAAAAAAAAAJpmMxyscD2Hp206dqBqN0qXZTIEAAAAAAAA
AAAAAAAAAAAAAJlmMxmuc0KztXVBqrx4Rc2eaDVhAAAAAAAAAAClbDlepGw5daJrOFuycj+Uv3lGyp5oNT4AAAAAAAAAAJlmNBWweky8qnNDbJJiLhupcUDAqHI/g6BsOFyXZTJppGs4r5toNWCkc0Y1pWw6nbBy
P6qcaDYwAAAAAAAAAACXYy8SrHpMpqRzRmK1hmCKvYtj0bOOd4lsYl3MJz5Vt55oNo6kbDautoFV37qJY62bajl1pnVICQAAAAAAAAAAAAAAAKl4TH23i2Tg3K+XX7qUd8s0Y4vfDkFyrjtunN5+XD2ki2A6joFg
QJLImXWHzqOHgah3S5Ggbj4fAAAAAK9/VSW7jGiLyZ5+nLyMZK2piWvlqcPX/6rD2P+Srsj/MmCH/xg+ZcsMMlkfkWQ3cs+hgoXVqpF7oXBAmKBuPg2hb0Bw2ayQedyynFWVemfF2puD/ba2uf9XtfH/U6vn/z+i
5/9nmsT/Gkh4cFNHOgWweUae3LOdY8WYeZahbj5jsoBVlryekmhcc47DaJm7+OnHs/+QqLf/VbLv/4/R/f9JqOj/esf5/7LN4/83VXZrf1kyYc+ggIrcsp1spHNFm7eGXaG9oJVhYHSMtTeY2vx8veT/XbXu/1q4
9/9fuPP/T6/v/9ju/v+35P//lbTN9ktNS3vBjmaW4rilXLGCWaGyg1us5bmibt+vk19XaHjVVZzR/3i55/8tgsD0DFGMxzeGwfz8////veH7/0aw4f9BaXjFsX9YoeS5pWW3iGGkqHdKhtuwm2zmt51mlHpoqhM8
ZHAAM2tUAC9kKQApXQMSRXhsqLrL9Mjm+f9FneH/PWWC5baFX6Dht6Jkrn1UnptoNlXPo4hz4bikXs2fgIakeVREAAAAAAAAAAAAAAAAAAAAAAU2aS8kWId4OmSWlFBSVZTGlG992K6WXZ9sO36WYy8VsoNbkd+1
oWXetJ5juolhhLyIXSsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACZbUN82aySdrmKZZ+XYy8wAAAAAJhlMTrBk3GL4LaiYdyxnGO5imWBrHpNPK55SAQAAAAAAAAAAAAAAACyf1FCvo5omb6Rb4iZZTNFAAAAAAAA
AAAAAAAAmWYzO7yNaYPbsZt04bekY8yghHe6jGZoqXhLQZ9qNT2cZzBYpHJCqKl5TXGXZDAjAAAAAAAAAAAAAAAAAAAAAAAAAACWYy8ZpHNFVruNaY3JnH51xJd3Z7KDXJKjckOFmWYyPJhkMQ4AAAAAAAAAAAAA
AAAAAAAAwwesQYMDrEEAA6xBAAOsQYABrEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEHgKxBA/CsQYDhrEHAA6xB4A+sQQ==
"@
#endregion ******** $PILFormIcon ********
$PILSmallImageList.Images.Add("PILFormIcon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($PILFormIcon))))

#region ******** $ExitIcon ********
$ExitIcon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB1ToVAgWar/F0B5/wAAAFcAAABNAAAAIQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAABxQnCIfWKjvI1+x/xY8cv8AAABDAAAAPgAAADoAAAAmAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAcUp/TIV6x/ylktP8VOW3/AAAALgAAACkAAAAjAAAAHgAAABgAAAAFAAAAAAAA
AAAAAAAAAAAAAAAAAAAjYLP/HlSg/yVkt/8ybrz/FDhq/wAAADgAAAAuAAAAJAAAABgAAAAJI2Cz/wAAAAAAAAAAAAAAAAAAAAAAAAAAJWS3/yFZpv8par3/O3fC/xM1Zf8AAAA4AAAALgAAACQAAAAYAAAACQ6F
RP8AcwD/AAAAAAAAAAAAAAAAAAAAAGRziP8kXqr/LG/C/0Bup/85SmL/AAAAOAAAAC4AAAAkAAAAGACUAKQAmQD/AHMA/wAAAAAAAAAAAAAAAAAAAAAqar7/J2Ov/y90yP9adpb/Tl93/wAAADgAAAAuAAAAJACE
AG8AmQD/a8lw/wBzAP8AcwD/AHMA/wBzAP8AcwD/LG3B/ylmtP8yec3/VJHV/xQ2Z/8AAAA4AAAALgB7AHcAmQD/V8Bb/0q8T/9Yw17/X8Zm/2bKbv910H3/AHMA/y1ww/8rarf/NH3Q/1yZ2/8UNmf/AAAAOAAA
AC4AmQD/fs6A/0K4Rv82tDv/PrhE/0a8TP9QwVf/Zspu/wBzAP8vcsb/LWy6/zeA0/9koOD/FDZn/wAAADgAAAAuAHsAdwCZAP+Az4L/ccp0/2zIcP93zXz/ftCC/2jJbv8AcwD/MHTI/y5uvP84gtb/YaDh/xQ2
Z/8AAAA4AAAALgAAACQAhABvAJkA/4TQhv8AmQD/AJkA/wCZAP8AmQD/AJkA/2Z2i/8vcL7/P4jZ/4m/7v8vYpa/AAAAOAAAAC4AAAAkAAAAGACUAKQAmQD/AJkA/wAAAAAAAAAAAAAAAAAAAAAxdsr/M3TA/3ey
6P9xnsO/AAMFRgAAADgAAAAuAAAAJAAAABgAAAAJEoxM/wCZAP8AAAAAAAAAAAAAAAAAAAAAMnfL/0yMz/9FgsDeAAMFRgADBUYAAAA4AAAALgAAACQAAAAYAAAACTJ3y/8AAAAAAAAAAAAAAAAAAAAAAAAAADJ3
y/8yd8v/MnfL/zJ3y/8yd8v/MnfL/zJ3y/8yd8v/MnfL/zJ3y/8yd8v/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAwP+sQYA/rEGAH6xBAB+sQQAPrEEAD6xBAACsQQAArEEAAKxBAACsQQAArEEAD6xBAA+sQQAfrEEAH6xB//+sQQ==
"@
#endregion ******** $ExitIcon ********
$PILSmallImageList.Images.Add("ExitIcon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($ExitIcon))))

#region ******** $HelpIcon ********
$HelpIcon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC6XSwQvF8sgMFiLd/IZzH/0nI7/9Z7Q//UeD7/0W8yz9VyM2AAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAC2Wiswulwr38lgKP/WaCv/2nE1/998P//jiUr/6Jlb/+ypbv/kl1r/23s7z9x3NRAAAAAAAAAAAAAAAACyVykwuVkr78xbKf/SYSX/1mkq/+CPXv/w4t7/8NrM/+uWUv/vnFL/9bRu//W9
f//ffj3P23Y1EAAAAAAAAAAAs1cqz8heMf/OWiL/0mEk/9ZpKv/lrIz/8Ojo//Do6P/spnD/75pO//SlVv/5uG//8bZ3/9h0NK8AAAAArlMoYLxeM//LWyj/zlke/9FgJP/VZyn/23c7/+rFsf/qvJ3/6IxE/+yV
S//xnVH/9KNV//S2df/fjVD/03AzMKxSKK/Ia0L/ylMb/81XHf/QXiL/1GUn/+Sri//x6ur/8enp/+WGQP/pjkb/7JRK/+6YTf/unFP/56Bm/9BuMo+yWzP/0XhP/8tUHf/MVRz/z1sg/9NiJf/gmXL/8uvr//Lq
6v/mm2f/5YY//+eLQ//pjkX/6Y5F/+eZXv/MazG/tWA5/9J4Tv/QZTL/zVgh/85YHv/RXiP/1GUn/+7UyP/z7Oz/79bK/+GERP/igTz/44M+/+ODPv/iilD/yWkwv7VhO//Vflb/0GY0/9FnNP/RYy3/z10j/9Jg
JP/YdkD/8uXi//Pt7f/sx7P/3Xc1/955Nv/eeTb/3ntD/8VmL7+qUyzf2o5s/9BlNP/RZjT/0mg0/9NrNf/UajL/1Ggu/9+Saf/17+//9O7u/91/R//acDD/2nAw/9hwOP/CYy6vpUwmj9GKav/TckX/0GY0/9Fn
NP/puaL/7su7/96PZv/rwq3/9vLy//Xw8P/ei1z/2nU7/9p1O//PbTX/v2AtYKRMJSC0Yj7/3ZNy/9BlNP/WeU3//v39//z7+//7+Pj/+fb2//j19f/z4tv/1m83/9dwOP/Vbzf/v2Au77xeLBAAAAAApEslgMd9
XP/bjGn/0Gg4/9mCWv/wzr///fz8//z6+v/uy7r/2oVZ/9RrNv/UbDn/wmEv/7hbK1AAAAAAAAAAAAAAAACjSyWfw3dV/96Xd//UeE7/0Gg4/9BmNP/RZjT/0Wk4/9JuQP/QbUP/vV4x/7RZKmAAAAAAAAAAAAAA
AAAAAAAAAAAAAKNLJXCtWDLvxntZ/86DYv/XjGv/1Ihl/8h1T/+9ZTz/slcr37FWKVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAApEwlEKVMJmCmTiaPqE8mv6lQJ7+rUSeArFIoUAAAAAAAAAAAAAAAAAAA
AAAAAAAA4A+sQcADrEGAAaxBgAGsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBgAGsQcADrEHgB6xB8B+sQQ==
"@
#endregion ******** $HelpIcon ********
$PILSmallImageList.Images.Add("HelpIcon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($HelpIcon))))

#region ******** $CheckIcon ********
$CheckIcon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADKyN4A5tT7vP7hF30K6SDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAACasKp9Eu0r/qe6y/5Dimf8+uETPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABqmHJ84tT3/oOip/6DurP+19b//ctN6/z24Q4AAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2gDp9Gu0v/oOap/47lmv+E5pH/ne+p/6/zuf9KvlD/O7dBMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKaAp87tT//mOCh/4Tdj/963of/qOyy/5nspf+s87f/l+af/zi1
Pc8AAAAAAAAAAAAAAAAAAAAAAAAAAACYAI80sTj/kNuZ/3nWhf9/2ov/n+ao/3nUgP+p7bP/keue/7LzvP9s0HT/N7U8gAAAAAAAAAAAAAAAAAAAAAAAlwD/idaS/3DPe/9104D/l+Cg/0K5R/8ToxX/gtiJ/53q
qP+S6p//p+2x/0O7SP81tDowAAAAAAAAAAAAAAAAAJYA/37Qhv+F1Y//htaO/y+vMv8HnAefEKESQCWrKP+j6Kz/h+SU/6HrrP+M35T/MrI33zSzORAAAAAAAAAAAACVAI8hpST/RLZJ/xGgE/8AmQCPAAAAAAAA
AAAPoBCfWMRe/5zlpv963oj/pequ/2LJaf8xsjaPAAAAAAAAAAAAAAAAAJUAMACXAGAAmAAwAAAAAAAAAAAAAAAADJ8NEA6gD++J2pL/h96T/4DdjP+Y4qH/O7ZA/zCxNEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAJngpgL68y/5fgoP9r1Hj/jN2W/33Uhf8srzDfLrAyEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAidCb9nyW7/idqT/1rLaP+U3p3/V8Jd/yuuL48AAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEmwUwD6AQ/4XVjf9mzHL/Zsxy/4fXkP8mrCr/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOaA4A/tUP/iNWR/2/Nef+J1pH/JKwq/wAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAmQAQApoC30y6Uv9mxW3/TLpR/xqmHL8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACZACACmgKfCJ0Ivw6g
D58ToxUQ+H+sQfB/rEHgP6xBwB+sQYAfrEEAD6xBAAesQQADrEEGA6xBjgGsQf8ArEH/gKxB/4CsQf/ArEH/wKxB/+CsQQ==
"@
#endregion ******** $CheckIcon ********
$PILSmallImageList.Images.Add("CheckIcon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($CheckIcon))))

#region ******** $UncheckIcon ********
$UncheckIcon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA8Ps48SErOvGhqtEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB4eqRAqKsmvLy/TjwAA
AAAAAAAAAAAAAA0Nsp8GBrf/CQm6/xQUtM8bG60QAAAAAAAAAAAAAAAAAAAAAB0dqhAnJ8bPODju/0BA9/81NdqfAAAAADc3v48HB7X/Bga3/wkJuv8MDL3/FRW1zxsbrBAAAAAAAAAAAB0dqhAjI8LPLy/k/zY2
7P87O/L/PT30/y4u0I9ERMSvVFTO/wkJt/8ICLn/Cwu8/w4OwP8WFrfPHBysEBwcqxAfH77PJiba/yws4P8xMeb/NTXr/zY27P8pKcevGRmvEEVFxM9VVc//Cwu5/woKu/8NDb//ERHD/xgYuM8bG7vPHh7R/yMj
1v8nJ9v/Kyvg/y4u4/8mJsXPHh6pEAAAAAAZGa8QRUXEz1VV0P8MDLv/DAy+/w8Pwf8TE8X/FxfJ/xsbzv8fH9L/IiLW/yUl2f8iIsHPHR2qEAAAAAAAAAAAAAAAABkZrhBGRsTPVlbQ/w4OvP8ODr//ERHD/xQU
xv8XF8r/GhrN/x4e0f8eHr7PHR2qEAAAAAAAAAAAAAAAAAAAAAAAAAAAGhquEEZGxc9NTc//DAy9/w4OwP8REcP/FBTG/xYWyf8bG7rPHByrEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkZrhAlJbjPLi7E/x0d
wf8ODr7/Dg7A/xAQwv8TE8X/GBi4zxwcrBAAAAAAAAAAAAAAAAAAAAAAAAAAABkZrxAmJrnPNDTE/zIyxf8wMMX/Ly/G/ygoxf8gIMT/Hx/F/x8fxv8jI7rPGxusEAAAAAAAAAAAAAAAABgYsBAoKLnPOTnE/zY2
xP80NMT/MjLF/zAwxf9oaNf/MTHH/y4ux/8uLsj/LS3J/yMjuc8bG60QAAAAABcXsRArK7rPPj7F/zs7xf85OcT/NjbE/zQ0xP8lJbjPT0/Hz3Bw2f8yMsb/Ly/G/y4uxv8uLsf/JCS5zxoarRBTU8qvVFTM/0FB
xv8+PsX/OzvF/zk5xP8mJrnPGRmuEBoarhBQUMfPcXHY/zMzxf8wMMX/MDDF/zAwxf8kJLivY2PQj5iY5v9TU8z/QUHG/z4+xf8oKLnPGRmvEAAAAAAAAAAAGhquEFBQx89yctj/NTXF/zIyxP8yMsT/Jye6jwAA
AABpadOfmJjm/1RUzP8rK7rPGBiwEAAAAAAAAAAAAAAAAAAAAAAZGa8QUFDIz3R02P84OMT/KSm7nwAAAAAAAAAAAAAAAGNj0I9TU8qvFxewEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkZrxBQUMivSkrFjwAA
AAAAAAAAx+OsQYPBrEEBgKxBAACsQQAArEGAAaxBwAOsQeAHrEHgB6xBwAOsQYABrEEAAKxBAACsQQGArEGDwaxBx+OsQQ==
"@
#endregion ******** $UncheckIcon ********
$PILSmallImageList.Images.Add("UnCheckIcon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($UncheckIcon))))

#region ******** $FavoriteIcon ********
$FavoriteIcon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAmk9lALofTfDp3xcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACZ52WApdde/LHHTcAAA
AAAAAAAAAAAAAAAAAAAJpfeAHLb6/xSn9P8RmO6vFJTrEAAAAAAAAAAAIILfECJ/3a8ukuP/PK7t/yxx058AAAAAAAAAAAAAAAAAAAAACaT2QBWu+P8lxv//F673/xWT6t8YjudAHYfiMB+E4N8npe7/PNL//zKT
4/8rctRQAAAAAAAAAAAAAAAAAAAAAAAAAAALofTvKMf//x3B//8UtPv/F5Pq/xiO5/8Wrvj/IcT//yzJ//8qe9n/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC6D0ryG7+/8ixP//GL7//w+2/f8Gsf3/B7X//xK7
//8drvX/KHbXrwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAyf82Aisfb/NMz//x7C//8Vvf//DLj//wW0//8Es///GZPp/yd32GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALoPSfLbT1/2zi//9E0v//HcH//xW9
//8Puf//Crb//x2M5f8odtefK3LUEAAAAAAAAAAAAAAAAAWp+jAIpvfPO8H5/2/j/v905f//cuP//2be//9Bz///I8P//xS8//8Ru///H5To/yxx088ubdEwAAAAAAGv/nANr/vvWdj9/3vs//956f//d+f//3Xm
//905P//cuP//3Li//9i2///UdT//0XQ//87tvT/NHPS/zJnzXAavf//fu3//4Pw//9/7v//fOv//3rq//956P//d+b//3bl//915P//deT//3Xj//904///dOP//3Dc/P8/ftb/Aa/+3wSr+/8Ipvj/C6H0/yOr
9P8srPL/aNr7/3vp//966P//btr7/zeZ5v81juD/KXXX/yxx0/8wbND/MmfN7wAAAAAAAAAAAAAAAAAAAAAAAAAAE5bscEO69P9+7P//fer//0iw7f8hgd5wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAABOW7BAinu3/g+7//4Ht//8rkeT/IYHeIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFZLqr2TS+P9y3Pr/HoThzwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAABaR6WBDtPD/Tbfv/x2F4mAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWkekQIJTo7yOQ5e8dhuIQAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAx+OsQcGDrEHAA6xB4AesQeAHrEHgB6xB4AOsQYABrEEAAKxBAACsQQAArEH4H6xB+B+sQfw/rEH8P6xB/D+sQQ==
"@
#endregion ******** $FavoriteIcon ********
$PILSmallImageList.Images.Add("FavoriteIcon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($FavoriteIcon))))

#region ******** $AddItems16Icon ********
$AddItems16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJfAz/D44Y/w+OGP8Pjhj/EI8Z/xCP
Gv8KfA3/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD48Z/4fknP+C5Jj/g+SZ/4Plm/9645X/EZAe/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABCPHP+J5qD/J9JU/yjT
V/8o1Fr/fOWb/xKRIf8AM6n/ADOp/wAzqf8AM6n/ADOp/wAzqf8AM6n/ADOp/wAAAAARkR//i+im/ynWXv8q12D/K9hj/37oof8SkiT/BlriMAZZ4v8GWOH/Blfh/wZW4P8FVd//BVXf/wVU3jAAAAAAEpIi/4zq
q/8s2Wf/Ldpq/y7bbf9/6qf/E5Mn/wAAAAAHX+aAB1/l/wde5f8HXeT/B1zk/wZb44AAAAAAAAAAABOTJP+N7LD/L91x/zDedP8w33b/geyt/xOVKv8AAAAAAAAAAAhl6tAIZOn/CGTp/whj6NAAAAAAAAAAAAAA
AAASjyL/Ks1l/yrOZ/8rz2n/K9Br/yvQa/8TjyX/AAAAAAAAAAAJbO4wCWvu/wlq7f8JauwwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAApy8oAKcfGAAAAAAAAA
AACgjW1loo9vsqCNbbKYhWWyjXpaZQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACrmHhltKGBsrelhLK0oYGyq5h4sp2KarKNelplAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAu6iIssi1lLLNupmyyLWUsruoiLKrmHiymIVlsgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMi1lLLZyKWy4tGustnIpbLItZSytKGBsqCNbbIAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAADNupmy4tGusvfmw7Li0a6yzbqZsrelhLKij2+yAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAyLWUZdnIpbLi0a6y2cilssi1lLK0oYGyoI1tZQAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADItZRlzbqZssi1lLK7qIiyq5h4ZQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA/4CsQf+ArEH/gKxBAICsQQCArEGBgKxBw4CsQcP/rEHmD6xB/AesQfwHrEH8B6xB/AesQfwHrEH+D6xB//+sQQ==
"@
#endregion ******** $AddItems16Icon ********
$PILSmallImageList.Images.Add("AddItems16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($AddItems16Icon))))

#region ******** $Add16Icon ********
$Add16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMJjLr/EZS//x2cw/8pqMP/NbDH/0G4yvwAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADBYy7/9L2G/+6cT//ulkH/8ZxH/9BuMv8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwWIu//S9h//ok0//6Y9B//KeSf/PbjL/AAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMBiLv/0vYf/5o9O/+iLQf/ynkj/zm0x/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADAYS3/9LyG/+aOTf/niT//8ZxG/81s
Mf8AAAAAAAAAAAAAAAAAAAAAAAAAALBWKb+zWCr/tlor/7lcK/+8Xiz/vmAt//O8hv/ki0v/5YY9/++YQ//LajD/zWwx/85tMf/PbjL/0G4y/9BuMt+vVSn/8LmK/+aTU//jh0D/44I2/+OCMv/pk0n/4H46/+OB
Of/tlD//75lD//GcRv/ynkj/8p5J//GdR//NbDH/rlQo//C5iv/dg0z/2HM5/9p2Ov/beDr/3Hc2/9t0M//eeDX/44E5/+WGPf/niUD/6IxB/+qPQv/ulkH/ymow/61TKP/wuYv/34xa/9p+S//cgUz/3oRO/+CH
T//dfUL/3n5A/+KGRP/mkVT/55JU/+iUVf/qmFb/76FX/8dnMP+sUij/77mL//C5i//xuor/8buL//G8i//yvYv/4IdQ/95+P//rnVf/9MCN//XBjf/1wY3/9cGN//XBjf/EZS//qlEnv61SKP+vVCj/sVYp/7NY
Kv+2Wir/8byL/96ETv/cej3/5o5E/75gLf/AYS3/wGIu/8FiLv/BYy7/wmMupgAAAAAAAAAAAAAAAAAAAAAAAAAAs1gq//G7i//cgU3/2nY7/+WLQv+8Xiz/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAALFWKf/xuov/239L/9hzOv/jiUL/uVwr/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACvVCj/8LqL/9+MW//dg03/5pRT/7ZaK/8AAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAArVIo//C6i//wuov/8LqL//C6i/+zWCr/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKpRJ7+sUij/rVMo/65UKP+vVSn/sFYpvwAAAAAAAAAAAAAAAAAA
AAAAAAAA+B+sQfgfrEH4H6xB+B+sQfgfrEEAAKxBAACsQQAArEEAAKxBAACsQQAArEH4H6xB+B+sQfgfrEH4H6xB+B+sQQ==
"@
#endregion ******** $Add16Icon ********
$PILSmallImageList.Images.Add("Add16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Add16Icon))))

#region ******** $Copy16Icon ********
$Copy16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAsbGx/4+Pj/+Hh4f/h4eH/4eHh/+Hh4f/h4eH/4eHh/+Hh4f/h4eH/4eH
h/8AAAAAAAAAAAAAAAAAAAAAAAAAALa2tv////////////////////////////////////////////////+Hh4f/AAAAAAAAAACxsbH/j4+P/4eHh/+2trb///////v7+//7+/v/+/v7//v7+//7+/v/+/v7//z8
/P//////h4eH/wAAAAAAAAAAtra2////////////tra2///////bzb//282//9vNv//bzb//282//9vNv//bzb///////4eHh/8AAAAAAAAAALa2tv//////+/v7/7a2tv///////f39//39/f/9/f3//f39//39
/f/9/f3//f39//////+Hh4f/AAAAAAAAAAC2trb//////9vNv/+2trb//////9vNv//bzb//282//9vNv//bzb//282//9vNv///////h4eH/wAAAAAAAAAAtra2///////9/f3/tra2////////////////////
/////////////////////////////4eHh/8AAAAAAAAAALa2tv//////282//7a2tv//////282//9vNv//bzb//282//9vNv//bzb//282///////+Hh4f/AAAAAAAAAAC2trb///////////+2trb/////////
////////////////////////////////////////iIiI/wAAAAAAAAAAtra2///////bzb//tra2////////////////////////////yMjI/8jIyP/IyMj/yMjI/6ioqP8AAAAAAAAAALa2tv///////////7a2
tv///////////////////////////7a2tv////////////r6+v+2trb/AAAAAAAAAAC2trb///////////+2trb///////////////////////////+2trb///////r6+v/ExMT/tra2nwAAAAAAAAAAtra2////
////////tra2////////////////////////////tra2//r6+v/ExMT/tra2nwAAAAAAAAAAAAAAALa2tv///////////7a2tv/IyMj/yMjI/8jIyP/IyMj/yMjI/7a2tv+/v7//tra2nwAAAAAAAAAAAAAAAAAA
AAC2trb///////////////////////////+2trb/+vr6/8TExP+2trafAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAtra2/8jIyP/IyMj/yMjI/8jIyP/IyMj/tra2/7+/v/+2trafAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA8AGsQfABrEGAAaxBgAGsQYABrEGAAaxBgAGsQYABrEGAAaxBgAGsQYABrEGAAaxBgAOsQYAHrEGAH6xBgD+sQQ==
"@
#endregion ******** $Copy16Icon ********
$PILSmallImageList.Images.Add("Copy16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Copy16Icon))))

#region ******** $import16Icon ********
$import16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACkpKQgoqKiWqGhoW+hoaF/oaGhhaGhoZOhoaGVoaGhlaGhoZWhoaGVoaGhk6GhoYWioqKAo6Ojd6Sk
pF+kpKQin5+fCLS0tJi3t7fZt7e337e3t+a3t7fot7e377e3t++3t7fvt7e37764vOi5uLjos7Ozuqurqzm0tLQsp6enDQAAAADR0dHZw8PD/8PDw//Dw8P/w8PD/8PDw//Dw8P/xsTF/9LEzf9zuYz/vMG9/9fP
1OQAAAAAAAAAAAAAAAAAAAAA1NTU1sXFxf/FxcX/xcXF/8XFxf/FxcX/z8bK/77Ewf9DtG//IbBa/266jf92u5HyNr1sgjbIcYwql1UbAAAAANvb29bMzMz/zMzM/8zMzP/MzMz/2s3T/4vEpP8gt2T/Kblr/y66
bv8ouGn/J7hp/yy5bP8uvW//I5NWSAAAAADg4ODW0dHR/9HR0f/R0dH/2tHV/2LJlv8QvWf/KMB0/yrAdf8qwHX/KsB1/yrAdf8qwHX/KsR3/x+XXEYAAAAA5ubm1tbW1v/W1tb/1tbW/9bW1v/e1tr/ddWo/xXJ
eP8fx33/I8h+/xvHe/8ax3r/Hsd8/x/Lf/8bnmNIAAAAAOrq6tba2tr/2tra/9ra2v/a2tr/2tra/+fa4P+63M3/KNaN/wzQgf922q//fduy80/ppZlV9a6jMbh/IgAAAADu7u7W3t7e/97e3v/e3t7/3t7e/97e
3v/e3t7/493g/+ff4v9Q36f/0t7a//fq8OMAAAAAAAAAAAAAAAAAAAAA8vLy1uLi4v/i4uL/4uLi/+Li4v/i4uL/4uLi/+Li4v/j4uL/7eLm/+Ti4v/v7+/kAAAAAAAAAAAAAAAAAAAAAPX19dbl5eX/5eXl/+Xl
5f/l5eX/5eXl/+Xl5f/l5eX/5eXl/+Xl5f/l5eX/8/Pz5AAAAAAAAAAAAAAAAAAAAAD5+fnW6Ojo/+jo6P/o6Oj/6Ojo/+jo6P/o6Oj/6Ojo/+jo6P/o6Oj/6Ojo//b29uQAAAAAAAAAAAAAAAAAAAAA+/v71urq
6v/q6ur/6urq/+rq6v/q6ur/6urq/+rq6v/q6ur/8/Pz//X19f/////pAAAAAAAAAAAAAAAAAAAAAP39/dbs7Oz/7Ozs/+zs7P/s7Oz/7Ozs/+zs7P/s7Oz/9fX1/4CAgP9ra2v/f39/ZwAAAAAAAAAAAAAAAAAA
AAD+/v7Z7e3t/+3t7f/t7e3/7e3t/+3t7f/t7e3/7e3t//n5+f9eXl7/XV1dZQAAAAAAAAAAAAAAAAAAAAAAAAAA/v7+k+7u7s/u7u7M7u7uzO7u7szu7u7M7u7uzO7u7sz5+fnRUlJSYgAAAAAAAAAAAAAAAAAA
AAAAAAAAAACsQQAArEGAB6xBgACsQYAArEGAAKxBgACsQYAArEGAB6xBgAesQYAHrEGAB6xBgAesQYAHrEGAD6xBgB+sQQ==
"@
#endregion ******** $import16Icon ********
$PILSmallImageList.Images.Add("import16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($import16Icon))))

#region ******** $LoadData16Icon ********
$LoadData16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABgAAACQAAAAgAAAADQAAAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAwEAC0ckE7mNTC33iUos9HA8I+JVLRrGOh0Qox8OBn4OBgJYAQAAMwAAABoAAAAJAAAAAQAAAAAAAAAAAAAAACYRB0uaWDn/ypB0/9W5qf/KpZL/w5V+/72Gav+3eFf/q2ZD/6FdO/2HTS/wazwk2k4q
GLwxGQ2XDQUCSAAAAAAwFglWoV8+/9qumP/98Oj//O/m//zu5P/87OP//Ovh//nm2//w18f/68m1/+a8o//hrpD/2ZZw/6VhPvgVCQMsMRYKVqhlQ//ctKD//vj1//738//+9vL//fXw//307v/98+z//fLr//3w
6f/87+f//O7l//vn2v/XjmT/OBoMWDEXClawbEn/3ril//7+/v///v3//v38//79+//+/Pr//vv5//76+P/++fb//vj0//738//99vH/1o9m/zscDVgxFwpWuHRP/82nlP9un2v/kLCJ/5OzkP+8w7j/ycvI/9TU
1P/y8vL//v7+///+/v/+/v3//v38/9ONZf88HQ1YMhcLVsF8Vv+shnP/Xb5Z/3zQdP9f217/bs9p/3DTbf9rzWf/scus/9LQzv/U1NP/1NTU/+7u7v/Pi2P/PB0OWDIYC1bJg1z/4Luo/6Wwo/+Sn4v/cphu/3WX
b/9glFv/XJ1Y/6/Fqv/o5eP/8/Ly//r6+v/T09P/yohg/zwdDlgzGAtW0oti/+K9qf///////////////////////v7+//X19f/f39//zs7O/8TExP/AwMD/8PDw/8aEXv89Hg5YMxkLVtqSaP/ivar/////////
///////////////////////////////////////////////////BgVv/PR4PWDMZDFbhmW7/2K+Z/+3t7f/z8/P/+vr6//7+/v//////////////////////////////////////vH5Z/z0eD1g1GgxW559y/8d/
WP/BlX7/v52J/76ikf++qJv/wrKo/8W6tP/HwsD/ysrK/9HR0f/W1tb/3Nvb/7p8WP89Hw9YQR8OReujdv/Jf1b/zIJZ/9GHXf/VjGL/2pFm/96Wav/im27/5J1v/6tjP/+sa0v/rHBT/3Znov+ygYX/QyERVzMY
CgOVVzWUuHVO27dzTe68eFH8wn1W/8iDW//PimD/1Y9l/9aRZv+3cEv/s21I/712UP9sXbH/hGKI+k0lEisAAAAAAAAAAAAAAAAAAAAAQyAPA0giDxNQJRAlVigRN14sE0lqMxhbdTodbXg8H4B5PSCQdjoeimcx
GDUAAAAAg/+sQQAHrEEAAaxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxB8AGsQQ==
"@
#endregion ******** $LoadData16Icon ********
$PILSmallImageList.Images.Add("LoadData16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($LoadData16Icon))))

#region ******** $Process16Icon ********
$Process16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAASAAAAFAAAABQAAAAUAAAAFAAAABQAAAAUAAAAFAAAABQAAAASAAAAAAAA
AAAAAAAAAAAAAAAAACkOJQDUH1MA/h9SAP8fTwD/HkwA/x5KAP8dRwD/HUQA/xxBAP8cPwD/GzsA/gsZANQAAAApAAAAAAAAAAAKHwLEOaEE/zqXAP84jQD/N4kA/zeFAP82ggD/NX4A/zV7AP80dwD/M3QA/zNx
AP8xawD/CRQAxAAAAAAAAAAAFkUH8DelDP80dwD/NHUA/zR1AP80dQD/Uoor/1WMMP80dQD/NHUA/zR1AP80dQD/M3IA/xQtAPAAAAAAAAAAABVHCvE1qBL/NoEA/zuECf9VkzL/NoEA/5O7dv+Zv33/NoEA/1aU
M/9Ahw//NoEA/zR3AP8VLwDxAAAAAAAAAAAUSQ3xM6wa/ziMAP9Dkw7/0eXG/26pUP+TwXX/lsJ4/2yoTf/R5cb/RZQR/ziMAP81fQD/FTEA8QAAAAAAAAAAE0oR8TGwIf86lwD/OJUA/0ScEP+IwWX/Q50M/0Od
DP+BvVz/Q5sQ/zeUAP85lwD/NoIA/xYzAPEAAAAAAAAAABJMFPEutSn/X7E2/9Tk0v/V5NL/dbtU/zqiA/86ogP/d7tV/9vo2f/b6Nn/X7E2/zeIAP8WNgDxAAAAAAAAAAARTRjxLLky/zarFf9Fsif/Qq8k/0yp
Mv80qBP/NKgT/0SmKf8/riD/P68g/zWrFP83jQH/FjgA8QAAAAAAAAAAEE8b8Sm9O/8vsyb/L7Al/5rLk/+l3qL/gMt6/33LeP+z47H/iMSB/y+yJf8vsyb/OJIC/xc6APEAAAAAAAAAAA9QHvEmwUT/Krw5/z/B
TP+R3Zv/K7w6/5Lcmv+N25X/Lr08/6Diqf84v0X/Krw5/ziXA/8XPADxAAAAAAAAAAAOUiLxJMVN/yTES/8kxEv/JMRL/yTES/9p14X/YtV//yTES/8kxEv/JMRL/yTESv85nAP/Fz4A8QAAAAAAAAAADk0f8CHJ
Vf8ix1H/I8ZO/yXESv8mwkX/J79A/ym9PP8quzf/LLkz/y23Lv8zrhz/O58A/xY/BPAAAAAAAAAAAAgXAsQkrD//IchV/yPFT/8lwkj/J79B/ym8O/8ruTT/LbYt/y+zJv8ysB//M6wZ/zKYEv8JFwLEAAAAAAAA
AAAAAAApDBoA1BtAAP4dRgD/HksA/x5PAP8fUwD/IFcA/x9UAP8eTwD/HUoA/xxDAP4MGwDUAAAAKQAAAAAAAAAAAAAAAAAAAAAAAAARAAAAFAAAABQAAAAUAAAAFAAAABQAAAAUAAAAFAAAABQAAAARAAAAAAAA
AAAAAAAA4AesQYABrEGAAaxBgAGsQYABrEGAAaxBgAGsQYABrEGAAaxBgAGsQYABrEGAAaxBgAGsQYABrEGAAaxB4AesQQ==
"@
#endregion ******** $Process16Icon ********
$PILSmallImageList.Images.Add("Process16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Process16Icon))))

#region ******** $Config16Icon ********
$Config16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ9lBQWlbwZqpW4ISgAAAACdYgMJnGEDGQAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ5kCgSiaQwsw54x/76VK/iseRJ+sIAUx7qRH/ymcQpJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKx4GHrAmD31uY0v1uLMh//Zvm3/0bFe/9Gy
WP/WuWL/qnYQcKNqDyCjaw8HAAAAAAAAAAAAAAAAAAAAAAAAAAC2iSaw2b5x/9O0Xv/t3Kz/5M+Y/9a5a//x4rj/6tii/82rVP/LqE7/sYEhmgAAAAAAAAAAAAAAAAAAAABtZEggpn8tvOvZpv/05sT/s4k4lbSF
OB60hjcWs4QxXNS2defhyof/4MmH/7eLJ7UAAAAAAAAAAFljVg5aZVwcOH6z2kma0P+HpZ3/TJ7d/6eCK7ekcBQYAAAAAAAAAACpdhWB8eK5/9/Igv+vfSCUAAAAAAAAAAA9dZKUC5Ln/0yv8/9cvf//L6n//0Wu
//95o6D/lZJa+rWHIdSzhhvhv5km//Dht//Rslf/togppQAAAAAAAAAAWH6Up6HX//98rczkZoaUfWeLmI9cptbyI6f//0Og0//Mq0n/za1M/+POjP/fx5D+7d2u/7KDI4IAAAAAAAAAADeArbeS0f//JH7O6VBs
eF1ba2c1PniZsW3F//9VmLb1wKRmxPPmxf/Zvnb/s4QzJ6p2IT2xgC0KAAAAAAAAAAB4iIp6oayw0m229f8nl///Dpj//zCr//+O0///Roy01gAAAAC5jT6SvJJDpKp2HgUAAAAAAAAAAAAAAAAAAAAAAAAAAHt0
YgqUtc39mrPC4pe50OyYxOP/anh0X2l7eUMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAenRhKX56aRN+gXYwdX12b3RvXAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA//+sQf//rEH8T6xB+AesQfABrEHwAaxB4AGsQYBhrEGAAaxBgAGsQYABrEGAR6xBwH+sQeD/rEH//6xB//+sQQ==
"@
#endregion ******** $Config16Icon ********
$PILSmallImageList.Images.Add("Config16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Config16Icon))))

#region ******** $Column16Icon ********
$Column16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AAAAAAYAABk4AgAePgEAHT0BAB09AQAdPQEAHT0AAB09AAAdPQAAHD0AABw9AAAcPQAA
HD0AAB0+AAAbPAAAAgwaGSY5hXzT+oZ92P+FfNb+hXzV/oR71f6DetX+gnnV/oF31f5/dtT+fnXU/n100/59dNP+fXTU/n921/0pJ0BaOjo5Rf//////////////////////////////////////////////////
////////////////////////X19daTc3N0P/////2dnZ//Pz8//Dw8P/r6+v/7u7u//r6+v/2NjY/7W1tf/Gxsb/5OTk/8nJyf/09PT//////1paWmc3NzdD/////9nZ2f+Dg4P/q6ur/76+vv+rq6v/iIiI/3x8
fP/Ozs7/jIyM/7Gxsf/CwsL/5eXl//////9aWlpnNzc3Q//////u7u7/mJiY/6ampv/MzMz/3t7e/6enp/+MjIz/yMjI/6ioqP/CwsL/srKy/729vf//////WlpaZzk5N0P///////////r78///////////////
/////////////////v////////////////////7//////15eW2cnJTNDsavq/7Cp6P+wqun/rafn/6+p6P+xq+j/sKro/6+p6P+rpef/pp/l/6Kb5P+imuP/oJjj/6GZ6P84NVFnFBEvQ1tO1P9ZTM//WEvO/1RG
zP9qXtP/enDY/3xy2f9/ddr/bWLU/09By/9AMcf/OyzF/zgoxf84J87/Eg1GZxYTMERkWNb/YVXR/2pe0/+yrOn/kYnf/5CI3/+Cedr/cmfV/3lv1/+Eetv/gHba/0M0x/9ENcj/RDTR/xYRSGgVEiw+Z1vc/2Za
2/9pXdv/g3ni/3ht3/+AdeH/al3b/2BT2P9eUNf/bF/b/2da2v9IOdH/RzjR/0Y22P8WEERiAgIDBBcULzsZFjVCGBU0QRUSM0EWEjNBExAyQRQRM0EUETNBExAyQRAMMkEPDDFBEg8yQREOMkIQDDA/AgIHC///
/wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP///wD///8A////AP//
/wD///8A//+sQf//rEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEH//6xB//+sQQ==
"@
#endregion ******** $Column16Icon ********
$PILSmallImageList.Images.Add("Column16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Column16Icon))))

#region ******** $Threads16Icon ********
$Threads16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGMkLQVnKjI3dDk1hW8zLKBfHipuWxwsBAAAAAAxV4A+M1+I2TNchKcyVHwbAAAAAAAA
AAAAAAAAAAAAAAAAAACCSS83rYRd59K/qf/Mt5n/ilY9/1sXJjUAAAAAMFmDcDiNvf83ibr/MlV9WwAAAAAAAAAAAAAAAAAAAABwNC0YeD4vBp1sNimpfkyo1cWt/6R8YP5aFyU2AAAAADBchpM8m8z/PIe09DJT
eyIAAAAAAAAAAAAAAACgbjcEj1s9xF0cKJ9gICgHj1oxU8u0kv+hdVD+YhwgnEQyUQkwZ5O4RaPS/z1znLwxUHgBAAAAAAAAAAAAAAAAmmUuCMuxkt62loL/ekEy2rCOdevLtJb/1sGp/6d6W/9ZIyq2NW+e6FOr
1v83XoZxAAAAAAAAAAAAAAAAAAAAAJpmLwGyjGSr5d/U/9zPuf/Eq4r/uZZx//f29v/v6uH/pXVU/11NXP9SjrfxLliBIwAAAAAAAAAAAAAAAAAAAAAAAAAAnGgxK6yEWLOwimConWs2QptnMle5l3Ps6ubf/+nj
1/+ldFP/VDRD0j49XwEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAn2UnNbGEVdLi1cb/4trJ/5lnSftdFyJ9AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AABXWFhIkH1k/d/Ouf/azbb/i1ZA/FgWJ08AAAAAAAAAAAAAAAAAAAAAaGVhBmlnZAEAAAAAAAAAAAAAAAAAAAAAJlyOjWeeuv+hg2P02sq1/866nf90OSvkWxomgVsZJX1aGSoub2tlQJKOitGFgX3AbmplTWxn
YA8AAAAAM1J3FkCBre5xqc7vXVZWPqt6SsTLs5f/zree/9C8rP/ApI//dTs10Y6LhsDv7u3/x8TC/8C/vPyDf3rffnhxnTpefZxprtb/U3mdiQAAAACaZkei39C+/8itjsS8m3O44tjG/6B4ZNeopaHm/////8jG
xP/6+Pj/xsXC/6iinv+AjJD/bI+o6EZZcBGQYTEb1r+m/8asmP9mJRpgAAAAALeWauKmfViPhIB7l6+tqu2no5/61tPS/7Ctqv+hnZr/4d3a/6ypp/N1cm6dgmtQZLWVcuXNuJ7/eUE692MkKyeaaDZFn202IWxo
YgFoZF0QbmtlYJaTjv/Cv73/ubWy//Hu7v//////397c/35+e7OHZD0lqHlEbKFzRY59RDAuAAAAAAAAAAAAAAAAAAAAAAAAAAB3c25OioeBwJuXk9irqKTfp6ShzoaDfohsamYTAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAwIesQcCHrEGAh6xBAAesQQAPrEEAD6xBgA+sQfwPrEH+B6xBngCsQQQArEEAQKxBAASsQQAArEEAA6xB4D+sQQ==
"@
#endregion ******** $Threads16Icon ********
$PILSmallImageList.Images.Add("Threads16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Threads16Icon))))

#region ******** $LoadConfig16Icon ********
$LoadConfig16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ+Cgo+fgoLfn4KCYKJvKWuiZAD/AAAAAKdoAIWdYgD/o2UARQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACfgoLvzLWz/6duFv+tbAX/vHQP/5ZgEPqnaAD/24co/7t1
EP+jZQD/AAAAAAAAAAAAAAAAAAAAAAAAAACfgoJgsJWU//jm4v+naAD/45U6//+9b///xYD//8WA///Jgf+lZAD/oXdQ/wAAAAAAAAAAAAAAAAAAAAAAAAAAn4KCz/jm4v/45uL/+Obi/6doAP//yIT//8WA/8uG
Kv//xYD//8WA/5BSAP+2dRP/AAAAAAAAAAAAAAAAn4KCQKuPj//45uL/nF0A/+yvZf//0pP//9ab/6doAP+nczD/r3AL///lsP/en0v/0ZE3/wAAAAAAAAAAAAAAAJ+Cgr/Ls7L/9uPg/6doAP+gXwD/uJdT///W
m/+gXwD/qXM1/69vCv//7MH/oWAA/wAAAAAAAAAAAAAAAJ+CglCrkZD//fLu//jp5v/04d7/89/d/6BfAP//1pv//9ab/8WOOP//7MH//+zB/72DKP+gXwD/AAAAAAAAAACfgoLf28zK///49f//+PX//vbx/59d
AP//1pv/p2gA/6FhAP//1pv/voku/6BfAP//1pv/n10A/6doAAifgoKAt6Cf///59///+vf///r4///6+P//+vj/n10A///6+P+naAD/9NGb/7R6G//039z/n10A/6J5VGkAAAAAn4KC/9vPz////Pv///v6///7
+v///Pr///z6///8+v///Pv/rXEQ/6BfDP+jYgD/9eLf/8qysf+fgoKAAAAAAJ+CgmClior/7ejn///+/v///fz///38///9/f///v3///7+///////68e//9uTg//fl4f/WwL7/n4KCrwAAAAAAAAAAn4KCYJ+C
gv/VyMj/////////////////////////////////+ezp//jm4//55+T/48/M/5+Cgt8AAAAAAAAAAAAAAACfgoIwn4KC37ehof/n4OD/////////////////z8HB/5+Cgv+fgoL/n4KC/5+Cgv+fgoKvAAAAAAAA
AAAAAAAAAAAAAAAAAACfgoKAn4KC37GZmf/PwcH/5+Dg/7ehof+fgoJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACfgoJQn4KCn5+Cgs+fgoLvn4KCIAAAAAAAAAAAAAAAAAAA
AAAAAAAA//+sQfBHrEHwA6xB4AOsQeABrEHAAaxBwAOsQYABrEGAAKxBAAGsQQABrEEAAaxBgAGsQcABrEHwH6xB/B+sQQ==
"@
#endregion ******** $LoadConfig16Icon ********
$PILSmallImageList.Images.Add("LoadConfig16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($LoadConfig16Icon))))

#region ******** $SaveConfig16Icon ********
$SaveConfig16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAiWVIn4ZjR/+DYEX/f15D/8zMzP/MzMz/zMzM/8zMzP/MzMz/zMzM/8zMzP9kSTX/YEcz/11E
Mf9ZQi//jGhKn4pmSf+2noj/0b+u/4xoSv//////gF5E/3paQP/w8PD/+Pj4/+rq6v/29vb/W0Mw/8Gvnf+/rZz/WUIv/41oS/+4oIr/0sGv/8WyoP+MaEr//////4BeRP96WkD/8PDw//j4+P/q6ur/9vb2/1tD
MP+Ufmv/v62c/1lCL/+NaEv/1MKx/8e0ov+ynYn/jGhK//////+bgm7/kXlm//Dw8P/4+Pj/6urq//b29v9bQzD/lH5r/7+tnP9ZQi//jWhL/9TCsf+1oIz/sp2J/4xoSv/49fT/7Ozs/9jY2P/w8PD/+Pj4/+rq
6v/s6un/W0Mw/5R+a/+/rZz/WUIv/41oS//UwrH/taCM/7Kdif+Oa07/hmNH/4BeRP96WkD/c1U9/21ROv9nTDf/YUgz/15GNP+Ufmv/v62c/1lCL/+NaEv/1MKx/7WgjP+ynYn/r5qG/6yXg/+plID/ppF9/6OO
ev+ginf/nYd0/5qEcf+XgW7/lH5r/7+tnP9ZQi//jWhL/9TCsf+1oIz/xdHz/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/w8/x/5R+a/+/rZz/WUIv/41oS//UwrH/taCM//r7/v/6+/7/+vv+//r7
/v/6+/7/+vv+//r7/v/6+/7/+vv+//r7/v+Ufmv/v62c/1lCL/+NaEv/1MKx/7WgjP/G1Pr/xtT6/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/lH5r/7+tnP9ZQi//jWhL/9TCsf+1oIz/+vv+//r7
/v/6+/7/+vv+//r7/v/6+/7/+vv+//r7/v/6+/7/+vv+/5R+a/+/rZz/WUIv/41oS//UwrH/taCM/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/xtT6/8bU+v+Ufmv/v62c/1lCL/+NaEv/1MKx/7Wg
jP/6+/7/+vv+//r7/v/6+/7/+vv+//r7/v/6+/7/+vv+//r7/v/6+/7/lH5r/7+tnP9ZQi//jWhL/9TCsf9cTUH/xtT6/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/xtT6/1xNQf+/rZz/WUIv/41o
S//UwrH/0sGv/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/xtT6/8bU+v/G1Pr/xtT6/8bU+v/Br53/v62c/1lCL/+NaEv/imZJ/4ZjR/+DYEX/f15D/3xbQf94WUD/dVY+/3FUPP9uUTr/a044/2dMN/9kSTX/YEcz/11E
Mf9ZQi//gACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQ==
"@
#endregion ******** $SaveConfig16Icon ********
$PILSmallImageList.Images.Add("SaveConfig16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($SaveConfig16Icon))))

#region ******** $Calc16Icon ********
$Calc16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAISAjByAfIycgHyM6IB8jOyAfIzsgHyM7IB8jOyAfIzsgHyM7IB8jOyAfIzsgHyM6IB8iJyEg
JAcAAAAAJCMnBzMyNGxwbGvYgHx664B8euyAfHrsgHx67IB8euyAfHrsgHx67IB8euyAfHrsgHx663Bsa9czMTRqJiUpBy8uMTGSjYvznZiV/5SQjv94d3v/fnx//5uWlP+dmJX/nZiV/5uWk/9+fH//eHd7/5SQ
jv+dmJX/kY2L8TAuMS9ST1Bpop2b/4SCg/86Qlb/OUJa/zhCWf9HTVz/mpaU/5mVk/9GTFz/OUJZ/zlCWv87Qlf/hYOE/6Kdm/9RTlBlW1hZc6mkov9ESVn/Pkdg/2dvg/9qcYT/OUNb/2prcv9naXD/OUNb/0BK
Y/9WX3X/O0Vd/0ZLWv+ppKL/WFVWa15bXHOuqqj/NDxQ/0JLZP+usrv/g4mZ/z5HYP9ZXWj/V1tn/0BJYf+KkJ7/ipCe/1dfdP81PVD/sKup/1pYWWphX2BzubWz/1ZbZ/9ian3/a3KF/210hv9eZXn/dnd9/3R1
e/9gZ3r/XWV6/210h/9gaHz/V1to/7m1s/9dW1xpZGJjc8C9u/+lo6T/b3N//36Ek/95f47/cXR8/7q3tv+5t7X/c3V+/3qAj/9+hJP/bXF9/6WkpP+/vbv/X15faGdlZnPHxcT/x8XE/8LAv/+enZ7/qaen/8fF
xP/HxcT/x8XE/8fEw/+op6f/np2e/8LAv//HxcT/x8XE/2JgYmdpaGly0M7N/8HAwP9fZHH/Nj9V/zxEWf9/gYn/z83M/8/NzP98f4j/O0NY/zY/Vf9gZXP/wsHB/9DOzf9kY2RlbGttctjX1v9iZnP/OkRc/1tk
eP8+SGD/OEFY/5iZnv+Vlpz/OEFZ/zxGX/89R2D/OkNc/2Vpdv/Y19b/ZmVnZHBvcHLd3Nv/Nz9T/1Rccf+xtb7/iI6c/ztFXf9scXz/aW56/zxGXv+Ijp3/iY+d/1Nccf85QVT/393c/2hnaWNzcnNy5+bl/0tS
Yv9VXnP/g4qZ/1Nccv9TW3D/f4KL/3t/if9TXHH/Ulty/1Jbcv9VXnP/TVRk/+fm5v9qaWphfn1+We3t7f+qrLD/cXeG/3J5iv91e4z/a3B+/9TU1f/S0tP/bnOB/3R7jP9zeov/b3WE/6utsf/t7e3/e3p7U5eW
lxTs6+vr8/Py/8jIyv+jpav/qKqv/97e3v/z8/L/8/Ly/93d3f+oqq//o6Wr/8nJy//z8/L/6ejo6JKRkg8AAAAA2djXLe/u7bDw8O/L8O/vy/Dv78vw7+/L8O/vy/Dv78vw7+/L8O/vy/Dv78vw8O/L7+7tsNPS
0ioAAAAAgAGsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBgAGsQQ==
"@
#endregion ******** $Calc16Icon ********
$PILSmallImageList.Images.Add("Calc16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Calc16Icon))))

#region ******** $ListData16Icon ********
$ListData16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFtUTys4My+BNTEujTUxLo01Mi6NNTIujTYyL401MS6MGhgWbQUFBSQFBQQBAAAAAAAA
AAAAAAAAAAAAAIB2bxKgmpbzy8vL/8/Pz//W1tb/3d3d/+Tk5P/q6ur/7Ozs/6aioP5zcG7OCgkJTgcGBgUAAAAAAAAAAAAAAACCeXExrKuq/319ff+1tbX/hoaG/8bGxv+Tk5P/2NjY/42Njf+LiYf/3Nzc/5iW
lOcPDg1eBwcGBQAAAAAAAAAAhn12MX59e/+srKz/W1tb/76+vv9hYWH/0tLS/6ampv+ampr/pKKg//T09P/b29v/mJaU6AsKCVAGBQUBAAAAAId+dzGko6H/eHh4/6ampv+AgID/tra2/4uLi//Q0ND/j4+P/+rq
6f/e3Nr/8vLy/9bW1v93dHLSBwYGJgAAAACJf3kxu7m4/4ODg/+8vLz/kpKS/9DQ0P+Xl5f/29vb/6enp//o6Oj/oqKh/83Kyf+PjYv/tbKv/iEeHHgAAAAAioF6McnHxv+CgoL/kpKS/7a2tv+lpaX/kZGR/7Gx
sf/MzMz/cHBw/8vLy/91dXX/xMTE/5OTk/9DPzycAAAAAIyDfDHLysn/hISE/8HBwf+FhYX/vb29/3p6ev/IyMj/hoaG/6SkpP+JiYn/pqam/4eHh/+urq7/Qz88nQAAAACNhH0x29rZ/8HBwf/j4+P/vLy8/+Hh
4f/AwMD/5+fn/8DAwP/k5OT/ubm5/9fX1/+4uLj/z8/P/0M/PJ0AAAAAjoV+MeHf3v+NjY3/7+/v/4uLi/+0tLT/rKys/319ff+oqKj/oKCg/4GBgf/c3Nz/f39//9PT0/9DPzycAAAAAI6FfzHZ2Nf/jY2N/+3t
7f+Kior/urq6/6Wlpf+Ghob/oKCg/6Ojo/95eXn/2tra/3Nzc//U1NT/Qz88nAAAAACPhX8x4+Lh/9vb2//29vb/1dXV//Pz8//Ly8v/8vLy/8PDw//i4uL/wcHB/9jY2P+4uLj/0NDQ/0M/O5wAAAAAj4Z/MbKw
r/+UlJT/ycnJ/3R0dP/AwMD/kpKS/7q6uv9xcXH/19fX/2xsbP+6urr/goKC/5aWlv9CPjucAAAAAIuBejGXlZT/ysrK/8XFxf+bm5v/tbW1/8zMzP+vr6//lpaW//Ly8v+SkpL/wMDA/7CwsP+Li4v/Qz88kQAA
AACJf3gNpJ2Y56empf/k4+L/uLe2/+zs6/+mpaX/4eDg/62srP/h4N//pKSj/97d3P+VlJP/pqKf+2hhXEAAAAAAAAAAAIyDfBOOhX5Ek4qDRZOLhEWSioNFkoqDRZKJgkWQiIFFj4d/RY2FfkWMg3xFiH94RYV8
diQAAAAAwAesQYADrEGAAaxBgACsQYAArEGAAKxBgACsQYAArEGAAKxBgACsQYAArEGAAKxBgACsQYAArEGAAKxBwAGsQQ==
"@
#endregion ******** $ListData16Icon ********
$PILSmallImageList.Images.Add("ListData16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($ListData16Icon))))

#region ******** $Export16Icon ********
$Export16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACkpKQgoqKiWqGhoW+hoaF/oaGhhaGhoZOhoaGVoaGhlaGhoZWhoaGVoaGhk6GhoYWioqKAo6Ojd6Sk
pF+kpKQin5+fCLS0tJi3t7fZt7e337e3t+a3t7fot7e377e3t++3t7fvt7e378DAuui7u7jos7Ozuqurqzm0tLQsp6enDQAAAADR0dHZw8PD/8PDw//Dw8P/w8PD/8vLxv/Nzcb/zc3G/9XVyf9iYqH/srK8/9ra
0eUAAAAAAAAAAAAAAAAAAAAA1NTU1sXFxf/FxcX/xcXF/83Nx/91daz/Xl6k/2Fhpf9nZ6f/ISGS/wAAhf97e7LoAAAAAAAAAAAAAAAAAAAAANvb29bMzMz/zMzM/8zMzP/b29D/HByY/wUFkv8ICJP/CAiT/w4O
lP8QEJX/BweT/w0NmssQEJMSAAAAAAAAAADg4ODW0dHR/9HR0f/R0dH/4ODU/yUlpP8PD57/EhKg/xISoP8SEqD/EhKg/xISoP8QEJ7/Fham/yAglxwAAAAA5ubm1tbW1v/W1tb/1tbW/+Xl2f8gIKv/CQml/wwM
pv8MDKb/EhKn/xUVqP8MDKX/JSW29EdH1j4AAAAAAAAAAOrq6tba2tr/2tra/9ra2v/h4dv/h4fO/3R0zP92dsz/fHzO/zIyuf8GBrD/bGzO8QAAAAAAAAAAAAAAAAAAAADu7u7W3t7e/97e3v/e3t7/3t7e/+Xl
3//n59//5+ff//Hx4P9jY83/rq7a//z87uMAAAAAAAAAAAAAAAAAAAAA8vLy1uLi4v/i4uL/4uLi/+Li4v/i4uL/4uLi/+Li4v/i4uL/7Ozj/+fn4v/v7+/kAAAAAAAAAAAAAAAAAAAAAPX19dbl5eX/5eXl/+Xl
5f/l5eX/5eXl/+Xl5f/l5eX/5eXl/+Xl5f/l5eX/8/Pz5AAAAAAAAAAAAAAAAAAAAAD5+fnW6Ojo/+jo6P/o6Oj/6Ojo/+jo6P/o6Oj/6Ojo/+jo6P/o6Oj/6Ojo//b29uQAAAAAAAAAAAAAAAAAAAAA+/v71urq
6v/q6ur/6urq/+rq6v/q6ur/6urq/+rq6v/q6ur/8/Pz//X19f/////pAAAAAAAAAAAAAAAAAAAAAP39/dbs7Oz/7Ozs/+zs7P/s7Oz/7Ozs/+zs7P/s7Oz/9fX1/4CAgP9ra2v/f39/ZwAAAAAAAAAAAAAAAAAA
AAD+/v7Z7e3t/+3t7f/t7e3/7e3t/+3t7f/t7e3/7e3t//n5+f9eXl7/XV1dZQAAAAAAAAAAAAAAAAAAAAAAAAAA/v7+k+7u7s/u7u7M7u7uzO7u7szu7u7M7u7uzO7u7sz5+fnRUlJSYgAAAAAAAAAAAAAAAAAA
AAAAAAAAAACsQQAArEGAB6xBgAesQYABrEGAAKxBgAGsQYAHrEGAB6xBgAesQYAHrEGAB6xBgAesQYAHrEGAD6xBgB+sQQ==
"@
#endregion ******** $Export16Icon ********
$PILSmallImageList.Images.Add("Export16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Export16Icon))))

#region ******** $Clear16Icon ********
$Clear16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACkMDAyPFRQUmBUUFJgVFBSYFRQUmBUUFJgVFBSYFRQUmBUUFJgMDAyPAAAAKQAA
AAAAAAAAAAAAAAAAAAxRUE/q6ujm//Hw7v/x8O7/8fDu//Hw7v/x8O7/8fDu//Hw7v/x8O7/6ujm/1FQT+oAAAAMAAAAAAAAAAAAAAA1paOh////////////////////////////////////////////////////
//+lo6H/AAAANQAAAAAAAAAAAAAANqinpf/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7/qKel/wAAADYAAAAAAAAAAAAAADapqKf//f39//39/f/9/f3//f39//39/f/9/f3//f39//39
/f/9/f3//f39/6mop/8AAAA2AAAAAAAAAAAAAAA2qqmo//z7+//8+/v//Pv7//z7+//8+/v//Pv7//z7+//8+/v//Pv7//z7+/+qqaj/AAAANgAAAAAAAAAAAAAANquqqf/6+vn/+vr5//r6+f/6+vn/+vr5//r6
+f/6+vn/+vr5//r6+f/6+vn/q6qp/wAAADYAAAAAAAAAAAAAADasq6n/+Pj3//j49//4+Pf/+Pj3//j49//4+Pf/+Pj3//j49//4+Pf/+Pf2/6upqP8AAAA2AAAAAAAAAAAAAAA2rKuq//b19P/29fT/9vX0//b1
9P/29fT/9vX0//b19P/29fT/9vX0//Tz8f+npaP/AAAANgAAAAAAAAAAAAAANq2sq//08vH/9PLx//Ty8f/08vH/9PLx//Ty8f/08vH/8/Lw//Lw7v/r6OT/nJiU/wAAADYAAAAAAAAAAAAAADatrKv/8e/t//Hv
7f/x7+3/8e/t//Hv7f/x7+3/8e/t/+/s6v/n4+D/1c7H/4F5cf8AAAA2AAAAAAAAAAAAAAA2ra2s/+3r6f/t6+n/7evp/+3r6f/t6+n/7evp/+jk4f/k39r/5eHd/+Tg3P9RTkvsAAAAGwAAAAAAAAAAAAAANq6t
rP/q5+T/6ufk/+rn5P/q5+T/6ufk/+jl4v/f2dT/6OXi/+bj4P9lZWTsAAAAMQAAAAAAAAAAAAAAAAAAADWtrKv/5uPf/+bj3//m49//5uPf/+Th3f/e2dT/08zE/+vp5/9kZGTrAAAALgAAAAAAAAAAAAAAAAAA
AAAAAAAMW1tb6u/u7P/w7uz/8O7r/+7r6f/n49//1M3G/72zqP9kZGTrAAAALwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACkODg6PFxcXmBcXF5gXFxeYFhUVmBMSEZgMCwqYAAAALgAAAAAAAAAAAAAAAAAA
AAAAAAAAwAOsQYABrEGAAaxBgAGsQYABrEGAAaxBgAGsQYABrEGAAaxBgAGsQYABrEGAAaxBgAOsQYAHrEGAD6xBwB+sQQ==
"@
#endregion ******** $Clear16Icon ********
$PILSmallImageList.Images.Add("Clear16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Clear16Icon))))

#region ******** $StatusGood16Icon ********
$StatusGood16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMwAFA5kCNybbAVtG+wFbRvsCNybbAAUDmQAAADMAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAKAQcDnQ5qNf4k0Hr/P+6e/0z2s/9M9rP/P+6e/yTRe/8OajX+AQcDnQAAAAoAAAAAAAAAAAAAAAAAAAAKBBMEwyCeOP870l7/S993/0bnjP9D7Jr/Q+ya/0bnjP9L33f/O9Ne/yCe
OP8EEwTDAAAACgAAAAAAAAAAAgYAnSmQGP81vTb/OctO/zPUZP8v23T/LN9+/yzffv8v23X/M9Rk/znLT/81vTf/KZAZ/wIGAJ0AAAAAAAAAMx1RAP43qA7/N7Ul/zG+Ov8sx03/KM1a/ybRYv8m0WL/KM1a/yzH
Tf8xvzv/N7Ul/zeoD/8dUQD+AAAAMwEDAJk2jwD/OqEC/ziqEv8ysyT/Lroz/yq/P/8pwkX/KcJF/yq/P/8tujT/MrMk/zirEv86oQL/No8A/wEDAJkNIQDbOZUA/zqbAP88oQP/OqgQ/zevHf80syb/MrUr/zK1
K/80syb/N68d/zqoEP88oQP/OpsA/zmVAP8NIQDbFDMA+ziPAP8+lwb/R6AQ/0ikEP9HqBP/RqsX/0WtG/9FrRv/RqsX/0eoE/9IpBD/R6AQ/z6XBv84jwD/FDMA+xQxAPs4iQD/UZwf/1ShIv9VpSL/Vaci/1ap
Iv9WqiL/Vqoi/1apIv9VpyL/VaUi/1ShIv9RnB//OIkA/xQxAPsMHgDbPIYH/2WiOf9lpjn/Zqg5/2aqOf9mrDn/Zqw5/2asOf9mrDn/Zqo5/2aoOf9lpjn/ZaI5/zyGB/8MHgDbAQMAmTZ2Bv95q1T/eq1V/3qv
Vf96sVX/e7JV/3uyVf97slX/e7JV/3qxVf96r1X/eq1V/3mrVP82dgb/AQMAmQAAADMaPAD+falb/5K5dv+Uu3f/lLx3/5S9d/+UvXf/lL13/5S9d/+UvHf/lLt3/5K5dv99qVv/GjwA/gAAADMAAAAAAQQAnUhz
Jf+uyZn/tM6g/7TOoP+0z6D/tM+g/7TPoP+0z6D/tM6g/7TOoP+uyZn/SHMl/wEEAJ0AAAAAAAAAAAAAAAoFCwDDYIFD/9Hfxv/Z5dD/2eXQ/9nl0P/Z5dD/2eXQ/9nl0P/R38b/YIFD/wULAMMAAAAKAAAAAAAA
AAAAAAAAAAAACgEDAJ0zSSD+oreQ/+Tr3f/3+fX/9/n1/+Tr3f+it5D/M0kg/gEDAJ0AAAAKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMwECAJkLGADbHS4N+x0uDfsLGADbAQIAmQAAADMAAAAAAAAAAAAA
AAAAAAAA8A+sQcADrEGAAaxBgAGsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBgAGsQYABrEHAA6xB8A+sQQ==
"@
#endregion ******** $StatusGood16Icon ********
$PILSmallImageList.Images.Add("StatusGood16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($StatusGood16Icon))))

#region ******** $StatusBad16Icon ********
$StatusBad16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMwADBpkAJDvbAEFd+wBBXfsAJDvbAAMGmQAAADMAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAKAAIJnQAxgP4Teuz/Nab//0e8//9HvP//Nab//xN77P8AMYD+AAIJnQAAAAoAAAAAAAAAAAAAAAAAAAAKAAQawwE10P8gaf//OIn//ziZ//84o///OKP//ziZ//84if//IGn//wE1
0P8ABBrDAAAACgAAAAAAAAAAAAAJnQAX0P8MO///GFj//xhq//8YeP//GIH//xiB//8YeP//GGr//xhY//8MO///ABfQ/wAACZ0AAAAAAAAAMwAAf/4ADf7/CCn//wk9//8JTf//CVn//wlg//8JYP//CVn//wlN
//8JPf//CCn//wAO/v8AAH/+AAAAMwAABZkAAOT/AAL9/wIS/v8CI///AjL//wI8//8CQv//AkL//wI9//8CMv//AiT//wIT/v8AAv3/AADk/wAABZkAADfbAADw/wAA9/8BA/z/Aw/+/wQb//8EJP//BCn//wQp
//8EJP//BBz//wMP/v8BA/3/AAD3/wAA8P8AADfbAABU+wAA6f8GBvD/EBD2/xAQ+v8QE/3/EBf+/xAa/v8QGv7/EBf+/xAT/f8QEPr/EBD2/wYG8P8AAOn/AABU+wAAUfsAAOL/Hx/r/yIi8P8iIvP/IiL2/yIi
+P8iIvr/IiL6/yIi+P8iIvb/IiLz/yIi8P8fH+v/AADi/wAAUvsAADLbBwfc/zk55/85Oev/OTnu/zk58P85OfL/OTnz/zk58/85OfL/OTnw/zk57v85Oev/OTnn/wcH3P8AADLbAAAFmQYGxv9UVOX/VVXo/1VV
6v9VVez/VVXt/1VV7v9VVe7/VVXt/1VV7P9VVer/VVXo/1RU5f8GBsb/AAAFmQAAADMAAGj+W1vh/3Z26P93d+r/d3fr/3d37P93d+3/d3ft/3d37P93d+v/d3fq/3Z26P9bW+H/AABo/gAAADMAAAAAAAAHnSUl
rf+Zmev/oKDt/6Cg7v+goO//oKDv/6Cg7/+goO//oKDu/6Cg7f+Zmev/JSWt/wAAB50AAAAAAAAAAAAAAAoAABTDQ0Ow/8bG8v/Q0PX/0ND1/9DQ9f/Q0PX/0ND1/9DQ9f/GxvL/Q0Ow/wAAFMMAAAAKAAAAAAAA
AAAAAAAAAAAACgAAB50gIGn+kJDV/93d9v/19fz/9fX8/93d9v+QkNX/ICBp/gAAB50AAAAKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMwAABJkAACzbDQ1J+w0NSfsAACzbAAAEmQAAADMAAAAAAAAAAAAA
AAAAAAAA8A+sQcADrEGAAaxBgAGsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBgAGsQYABrEHAA6xB8A+sQQ==
"@
#endregion ******** $StatusBad16Icon ********
$PILSmallImageList.Images.Add("StatusBad16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($StatusBad16Icon))))

#region ******** $StatusInfo16Icon ********
$StatusInfo16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMwUFBJk5NzDbXFtY+1xbWPs5NzDbBQUEmQAAADMAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAKCAcDnXNpQv7bz5n/9e3G//r24f/69uH/9e3G/9vQmv9zaUL+CAcDnQAAAAoAAAAAAAAAAAAAAAAAAAAKFRIGw7KdRv/k0Xb/7N6V//Hmr//07MH/9OzB//HmsP/s3pX/5NF2/7Kd
R/8VEgbDAAAACgAAAAAAAAAABwUAnamOH//Xu0T/38li/+TSff/p2ZL/692e/+venv/p2ZL/5NJ9/9/JY//Xu0X/qY4f/wcFAJ0AAAAAAAAAM2JPAf7KpRL/0rIv/9i8Sv/dxWD/4Mtx/+LPe//iz3v/4Mxx/93F
Yf/YvEr/0rMv/8qmE/9iTwH+AAAAMwQDAJmsiQP/w50D/8uoFv/QsC3/1bhB/9i9T//ZwFf/2cBX/9i9T//VuEH/0LEu/8uoFv/DnQP/rYkD/wQDAJkoHwLbsosH/7uUA//DnQT/yqYT/86sI//QsS//0rM2/9Kz
Nv/QsS//zqwj/8qmE//DnQT/u5QD/7KLB/8oHwLbPC4E+6qDCv+0jg3/vZkU/8OgEv/IpBT/y6ga/8yqHv/Mqh7/y6ga/8ilFP/EoBL/vZoU/7SODf+qgwr/PC4E+zkrBfuheg7/s5Ep/7qYKf++nSf/wqEm/8Wk
Jf/HpST/x6Uk/8WkJf/CoSb/v50n/7qYKf+zkSn/oXoO/zkrBfsiGQTbm3QY/7WWRf+5m0P/vZ9C/8CiQP/CpED/w6U//8OlP//CpED/wKJA/72fQv+5m0P/tZZF/5t0GP8iGQTbAwIAmYhjGP+4nmH/vKJg/7+l
X//Bp17/w6le/8OqXv/Dql7/w6le/8GnXv+/pV//vKJg/7ieYf+IYxj/AwIAmQAAADNEMAv+tJtq/8Ougf/GsYH/yLOB/8m0gP/KtID/yrSA/8m0gP/Is4H/xrGB/8Ougf+0m2r/RDAL/gAAADMAAAAABAMAnX1h
Nv/Pv6P/1MWp/9XGqP/Wx6j/1seo/9bHqP/Wx6j/1cao/9TFqf/Pv6P/fWI2/wQDAJ0AAAAAAAAAAAAAAAoMCALDiHJT/+LZzP/o4NT/6OHU/+jh1P/o4dT/6OHU/+jg1P/i2cz/iHJT/w0IAsMAAAAKAAAAAAAA
AAAAAAAAAAAACgQCAZ1OPiz+u6yb/+3n4f/5+Pb/+fj2/+3n4f+7rJv/Tj4s/gQCAZ0AAAAKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMwIBAJkbEQfbMiUX+zIlF/sbEQfbAgEAmQAAADMAAAAAAAAAAAAA
AAAAAAAA8A+sQcADrEGAAaxBgAGsQQAArEEAAKxBAACsQQAArEEAAKxBAACsQQAArEEAAKxBgAGsQYABrEHAA6xB8A+sQQ==
"@
#endregion ******** $StatusInfo16Icon ********
$PILSmallImageList.Images.Add("StatusInfo16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($StatusInfo16Icon))))

#region ******** $Selected16Icon ********
$Selected16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX5JgdgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIy9mvN0qX+DAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABot4rzWMCR/2imeIUAAAAAAAAAAAAAAAAAAAAAcql9hHuzioR8s4qEfLOKhHyzioR8s4uEfLSLhHy0i4R6somEZbuN+Saz
d/9Ru4r/cKt/hwAAAAAAAAAAAAAAADabYv81qXL/NK12/zSxef8ztXz/M7h//zO7gv8yvYT/Mr2D/y+5f/8msnb/Jq5y/0q2g/9vrH+JAAAAAAAAAAArlFn/KKFn/yela/8nqW7/Jq1y/yawdf8ms3f/JbV4/yW0
eP8msnb/Jq90/yascP8np23/RLB8/2epeosAAAAAK5NY/yifZv8oo2n/LKlw/zCudf8zsnr/NLV9/zW2f/82tn//N7V+/zezfP82r3n/Nat1/zSncP9Grnv/Y6d2jUCbZv9Orn//VrWH/1e4iv9Wuoz/VryO/1a9
j/9WvpD/Vr6Q/1a9j/9Wu43/VrmM/1e3iv9XtIf/XreK/2aoeI1iqn7/cbyX/3G+mf9xwZv/ccKc/3HEnf9wxZ7/cMae/3DFnv9wxZ7/ccSd/3HCnP9xwJr/ccCY/2uqfYsAAAAAdrWO/4zKq/+MzKz/jM2t/4zO
rv+Mz6//jNCv/4zQsP+M0LD/jM+v/4vNrf+LzKz/h8qo/3Sug4kAAAAAAAAAAHOrgIR9tY2EfraOhH62joR+to6EfraOhH62joR+to6EfbWNhIzFo/ml1r7/ntO3/3euhYcAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACWxKXzstvF/3SrgIUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAmcGi836uhoMAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGCTYnYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA//+sQf+/rEH/n6xB/4+sQQAHrEEAA6xBAAGsQQAArEEAAKxBAAGsQQADrEEAB6xB/4+sQf+frEH/v6xB//+sQQ==
"@
#endregion ******** $Selected16Icon ********
$PILSmallImageList.Images.Add("Selected16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Selected16Icon))))

#region ******** $Delete16Icon ********
$Delete16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA8Ps48SErOvGhqtEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB4eqRAqKsmvLy/TjwAA
AAAAAAAAAAAAAA0Nsp8GBrf/CQm6/xQUtM8bG60QAAAAAAAAAAAAAAAAAAAAAB0dqhAnJ8bPODju/0BA9/81NdqfAAAAADc3v48HB7X/Bga3/wkJuv8MDL3/FRW1zxsbrBAAAAAAAAAAAB0dqhAjI8LPLy/k/zY2
7P87O/L/PT30/y4u0I9ERMSvVFTO/wkJt/8ICLn/Cwu8/w4OwP8WFrfPHBysEBwcqxAfH77PJiba/yws4P8xMeb/NTXr/zY27P8pKcevGRmvEEVFxM9VVc//Cwu5/woKu/8NDb//ERHD/xgYuM8bG7vPHh7R/yMj
1v8nJ9v/Kyvg/y4u4/8mJsXPHh6pEAAAAAAZGa8QRUXEz1VV0P8MDLv/DAy+/w8Pwf8TE8X/FxfJ/xsbzv8fH9L/IiLW/yUl2f8iIsHPHR2qEAAAAAAAAAAAAAAAABkZrhBGRsTPVlbQ/w4OvP8ODr//ERHD/xQU
xv8XF8r/GhrN/x4e0f8eHr7PHR2qEAAAAAAAAAAAAAAAAAAAAAAAAAAAGhquEEZGxc9NTc//DAy9/w4OwP8REcP/FBTG/xYWyf8bG7rPHByrEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkZrhAlJbjPLi7E/x0d
wf8ODr7/Dg7A/xAQwv8TE8X/GBi4zxwcrBAAAAAAAAAAAAAAAAAAAAAAAAAAABkZrxAmJrnPNDTE/zIyxf8wMMX/Ly/G/ygoxf8gIMT/Hx/F/x8fxv8jI7rPGxusEAAAAAAAAAAAAAAAABgYsBAoKLnPOTnE/zY2
xP80NMT/MjLF/zAwxf9oaNf/MTHH/y4ux/8uLsj/LS3J/yMjuc8bG60QAAAAABcXsRArK7rPPj7F/zs7xf85OcT/NjbE/zQ0xP8lJbjPT0/Hz3Bw2f8yMsb/Ly/G/y4uxv8uLsf/JCS5zxoarRBTU8qvVFTM/0FB
xv8+PsX/OzvF/zk5xP8mJrnPGRmuEBoarhBQUMfPcXHY/zMzxf8wMMX/MDDF/zAwxf8kJLivY2PQj5iY5v9TU8z/QUHG/z4+xf8oKLnPGRmvEAAAAAAAAAAAGhquEFBQx89yctj/NTXF/zIyxP8yMsT/Jye6jwAA
AABpadOfmJjm/1RUzP8rK7rPGBiwEAAAAAAAAAAAAAAAAAAAAAAZGa8QUFDIz3R02P84OMT/KSm7nwAAAAAAAAAAAAAAAGNj0I9TU8qvFxewEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkZrxBQUMivSkrFjwAA
AAAAAAAAx+OsQYPBrEEBgKxBAACsQQAArEGAAaxBwAOsQeAHrEHgB6xBwAOsQYABrEEAAKxBAACsQQGArEGDwaxBx+OsQQ==
"@
#endregion ******** $Delete16Icon ********
$PILSmallImageList.Images.Add("Delete16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Delete16Icon))))

#region ******** $Up16Icon ********
$Up16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAH8A/wB7AP8AdwD/AHIA/wBuAP8AagD/AGYA/wBjAP8AAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAACDAP+Y05P/csVr/3fKcv99z3r/hNSC/4raiv8AZgD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhwD/mNOS/zOrKP87sjL/Q7g9/0y/SP+I2If/AGoA/wAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIsA/5fSkf8wqSX/OK8v/0C2Of9IvEP/hdWD/wBuAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACPAP+W0Y//Laci/zSsKv88sjP/Q7g9/4HS
fv8AcgD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAkwD/ldCN/yqkHf8wqSX/N68t/z60Nv98znj/AHcA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJcA/5PPi/8moRj/LKUg/zKq
J/84ry//eMtz/wB7AP8AAAAAAAAAAAAAAAAAAAAAAJkA/wCZAP8AmQD/AJkA/wCZAP+RzYn/Ip4T/yeiGv8tpiH/Mqon/3PHbf8AfwD/AHsA/wB3AP8AcgD/AG4A/wCZAO9ywW//rden/43KhP93wGz/hcd8/x6a
Dv8jnhT/J6Ia/yymIP9vw2f/csVr/3THbv92yXD/Ybtd/wByAO8AmQAwAJkA73HBbf+b0JP/XLNP/1SwRv9JrTz/Pqkx/zenKv83qCr/M6gn/zerK/9mv17/Ybpc/wiAB+8AdwAwAAAAAACZADAAmQDvb8Fs/5nP
kv9as03/U7BF/0+vQv9Or0D/TbBB/0+yQ/9zwmr/ab1j/weHB+8AfwAwAAAAAAAAAAAAAAAAAJkAMACZAO9vwWv/mc+R/1qzTf9UsEf/UrBF/1KxRf9zwWr/aLxh/weOBu8AhwAwAAAAAAAAAAAAAAAAAAAAAAAA
AAAAmQAwAJkA72/Ba/+az5L/XLNQ/1iyS/93wGz/aLxg/waVBu8AjwAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACZADAAmQDvcMBs/5zPlP9+w3T/cMBo/wabBe8AlwAwAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAJkAMACZAO9xwG3/jMqH/wqcCe8AmQAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAmQAwAJkA7wCZAO8AmQAwAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA8A+sQfAPrEHwD6xB8A+sQfAPrEHwD6xB8A+sQQAArEEAAKxBAACsQYABrEHAA6xB4AesQfAPrEH4H6xB/D+sQQ==
"@
#endregion ******** $Up16Icon ********
$PILSmallImageList.Images.Add("Up16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Up16Icon))))

#region ******** $Down16Icon ********
$Down16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAcHKwwHByr7xwcq+8dHaswAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAbG6wwICCu71RUwP9HR7v/HR2r7x0dqjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAbG60wICCu71RUwP9UVMD/VFS//0hIu/8dHarvHR2qMAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGq0wICCv71NTwf9TU8D/ICCu/yAgrf9UVL//SEi7/x0dqu8dHaowAAAAAAAAAAAAAAAAAAAAAAAAAAAaGq4wHx+v71NTwf9TU8H/Hx+u/xwcrP8cHKz/ICCt/1RU
v/9ISLr/HR2q7x0dqjAAAAAAAAAAAAAAAAAaGq4wHx+w71JSwf9SUsH/Hx+v/xsbrP8bG6z/HBys/xwcq/8gIK3/VFS//0hIu/8dHarvHR2qMAAAAAAZGa8wHh6w71xY0P9lYdP/OjbD/zYywv8bG63/Gxus/xsb
rP8cHKz/NzPA/zs3wv9mYtH/UU3J/yQjse8dHaswGRmv72th5f+IeP//iHj//4h4//+IeP//Gxut/xsbrf8bG6z/Gxus/3lv6f+IeP//iHj//4h4//90Z+//HByr7xkZr/8ZGa//GRmu/xoarv8aGq7/iHj//xoa
rf8bG63/Gxut/xsbrP9gYMX/HBys/xwcrP8cHKz/HBys/xwcq/8AAAAAAAAAAAAAAAAAAAAAGhqu/4h4//8eHq//Ghqt/xsbrf8bG63/YGDF/xsbrP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkZ
rv+IeP//VFTD/z8/u/8zM7b/KSmy/2lpyf8bG63/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZGa//iHj//1xcxf9YWMT/VVXD/1NTwv+GhtT/Gxut/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAGRmv/4h4//9eXsb/XFzG/1paxf9XV8T/iYnV/xoarf8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkZr/+IeP//a2fW/2ll1f9oZNT/Z2PU/4yL2f8aGq7/AAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAYGK//iHj//4h4//+IeP//iHj//4h4//+KgPD/Ghqu/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGBiw/yoqtv8qKrX/Kiq1/yoqtf8qKrX/Kiq1/xkZr/8AAAAAAAAAAAAA
AAAAAAAA/D+sQfgfrEHwD6xB4AesQcADrEGAAaxBAACsQQAArEEAAKxB8A+sQfAPrEHwD6xB8A+sQfAPrEHwD6xB8A+sQQ==
"@
#endregion ******** $Down16Icon ********
$PILSmallImageList.Images.Add("Down16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Down16Icon))))

#region ******** $Edit16Icon ********
$Edit16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQzImWFpFM4IkGBNABQEBFAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAH5kSoz65KL/28OL/rqdceePdFW+aE47lUkzJ3E0IhtWIRQQPBQLCSkKBAMdBAICCwAAAAAAAAAAAAAAAAAAAAB3X0aM9OKf//HkoP/w5KP/9OSn//Tjqv/t2ab/4cqe/9a+
lv3SuZX5zrCQ+z4qIl8AAAAAAAAAAAAAAAABAAADoYFhw7Clev+3rID/9eao/+zgqv/u47D/8Oa2//PqvP/27cP/+vTL//bpxP87KyNYAAAAAAAAAAAAAAAAKSMiPcnAoP2uoX3/QkVF/9DKqP/67LT/8OW4//Dm
vf/y6MH/8urG//fwzv/s3sH/LyIcRgAAAAAAAAAAAAAAAGp1cpzQ2Lz/5dGb/3uWqv9snd3/u7q5//rtxv/168f/8+vJ//Ttz//69Nn/5tjA+iAWEzIAAAAAAAAAAB4VEimkz83wz9O3//nqtf/OzcP/fqnk/zFj
t/+Gm5D/9e3H//ry1//28Nf//vrk/9zMufEPCggcAAAAAAAAAABpb2mSvvz8/9bLrf/37MD/5t3A/4Slu/9Jjjb/OYES/4aqWv/38Nz//Pbk////7//HuKnfBAAADwAAAAAfFRErsdfT7sL39//eza3/9u3I//ju
z/+uu5X/jbls/1qoHP9Biwn/kLJr//337f////3/rZ2TxgAAAAQAAAAAgYeAqs/////G4dz/6tu8//bv0P/48db/9O/U/8LNsP+axnn/W6Uj/0ePEP+fvYP//////417dacAAAAAVkxFcdH39f/R////z9bM//Ll
yf/38tn/+PLc//745//h483/uMeq/6LLhP95tEz/dKhP/8bLqv9lTUeBAAAAAHJbUILG19T44v///9/ay//79d7/+vfk//r45//7+ez///73/8/Ywv+uxZz/pMuG/7DRl/+wvpj/LCsXbwAAAAAAAAAAIRYSLIx/
d6rOrpz87tnI//Hk1v/27+T/+vTs///++///////qMeT/2WlNf/J3rj/2OnN/4KXdtsABAAiAAAAAAAAAAAAAAAALR8aO5WAeLHhzcX779nO/+jQw//AqaHcoY6KraqWkL1PfxzkhLdf//f49v/o8uD/NUMqdwAA
AAAAAAAAAAAAAAAAAAAAAAAALyQhMINtY5BcTUdqDAkIEAAAAAAAAAAADBgCNkaODvK20qD/+//4/0JPN30AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYNQJjRXoc3khb
OZAAAwAI8H+sQfAArEHwAKxB4ACsQeAArEHgAKxBwACsQcAArEGAAKxBgAGsQQABrEEAAaxBgACsQeAArEH4YKxB//CsQQ==
"@
#endregion ******** $Edit16Icon ********
$PILSmallImageList.Images.Add("Edit16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Edit16Icon))))

#region ******** $Content16Icon ********
$Content16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANAAQCRwAvHJUDHxVJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAACsACwaIACkV3gBZMP0AeUf/Bm9KzQRfQ6QObVWkJ3topEOG
eaRbjISka4+JoAAAAAAAAAAAAAAAAAAAAA0AGgvCADYV/gBMIf8AYjL/AHxK/wKbav8SuZD/NtO1/2Hk0f+J8OP/qvfv/5rLxcMAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAyEAHw5lAEIiqgBzROwFi13sDm1UjCR9
aow9h3qMVI6GjGeSjYw8T0xGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgECARgQJwAAAAAAAAAAAAAAIQZdQqQGIxtCAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkAGgrHACYRxwE1
HMcBSirHAl46xwNuSdkCsoX/Lc2s/lGzodo4X1l3BwwLFwAAAAAAAAAAAAAAAAAAAAAAAAAIAB8KwQA8F/8AUCX/AGg3/wCBTf8AmWb/ArSH/yvQsf9t5tP+jc/F0lVzcG8AAAAAAAAAAAAAAAAAAAAAAAAAAAAF
ARgBDwYzAhQKMwMaDzMDIBQzAkMrewNyVKART0JcDRoXGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA//+sQf//rEH//6xB4f+sQQAHrEEAB6xBwAesQfmHrEHgAaxB4AGsQfAHrEH//6xB//+sQf//rEH//6xB//+sQQ==
"@
#endregion ******** $Content16Icon ********
$PILSmallImageList.Images.Add("Content16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Content16Icon))))

#region ******** $Header16Icon ********
$Header16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANAAMGRwkgQpUJFypJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAACsBBxSIBRlJ3g03i/0VUKf/HlGNzRxJcqQrW3ukQW6DpFh+
iqRriI6kd4yPoAAAAAAAAAAAAAAAAAAAAA0CDTTCAxpo/gcof/8NOpT/FlOq/yZ0wf9DmtX/a7/l/5DY7/+w6fb/yPP6/63IzMMAAAAAAAAAAAAAAAAAAAAAAAAAAAAEECEDEDRlCSdlqhRMnuwkZ6/sKVp+jEFw
h4xWf46MaYmRjHePlIxDTU9GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQICBhEgJwAAAAAAAAAAAAAAIR1IcaQOHihCAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkBDTXHBBVJxwgg
XMcOMG7HFUB+xx1Qi9k5kNH/ZLbh/nKovNpFXGF3CAwMFwAAAAAAAAAAAAAAAAAAAAAAAAAIAQ1BwQMccP8IK4T/D0Ca/xlXrv8lcMD/OpLS/2a64/+Z2fD+pcnT0mBydG8AAAAAAAAAAAAAAAAAAAAAAAAAAAAC
ChgCBxozBAwgMwYRJjMJFiozEDBYeyRbh6AmRldcERgcGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA//+sQf//rEH//6xB4f+sQQAHrEEAB6xBwAesQfmHrEHgAaxB4AGsQfAHrEH//6xB//+sQf//rEH//6xB//+sQQ==
"@
#endregion ******** $Header16Icon ********
$PILSmallImageList.Images.Add("Header16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Header16Icon))))

#region ******** $Demo16Icon ********
$Demo16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8B////Af///wH///8BAAAAAAAAAAAAAAALERERHhERER4AAAALAAAAAP8AAAH///8B////Af//
/wH///8B////Af///wH///8BAAAAABYWFi0cHBqQHh0a0B4dGugdHRroHB4c0BgcGpAREREtAAAAAP///wH///8B////Af///wH///8BAAAAADY0NnAvLy7yIyIh/yAeHP8gHhv/Hx8c/x0gHf8dIR7/Ghwb8hYZ
GXAAAAAA////Af///wH///8BAAAAAC0tLXZCQUH/PDs7/ygoJ/8gHh3/IR8c/x8gHf8eIR//HR8d/xweHf8dIB//GRwcdgAAAAD///8BAAAAABUVFTsmJSX6NjU1/0NCQv85ODj/Hh0c/yMiH/8hIx//GRwa/x0f
Hv8eISD/HiIh/xsfHfoVFRU7AAAAAAAAAAAdHRyrIB8f/yMjIv80MzP/RURE/3Z1df+ioqH/paWk/2xtbP8iJSP/HyMh/x4gHv8dHx//GhwcqwAAAAATExMaHh0d6h8eHf8fHh7/Hh0d/359ff/Ly8v/5OTk/+np
6f/Hx8f/eHp5/xsdHP8eHx//HB4e/xsdHeoTExMaGxsbOB0dHf0fHh3/Hh0d/zAvL/+xsbH/6Ojo/9LS0tjT09PY4uLi/7CwsP8uLy//HR0d/x0eHv8cHR39FhsWOBkZGT0eHh3/Hh4c/x0cG/83NjX/v7+///Dw
8f+ioqKooqKiqPPz8//AwMD/Njc1/xwcG/8eHh3/Hh4d/xkZGT0aGhomHR0b9R4eHP8fHhz/IyMh/5ubmv/Q0ND/7Ozs/+vr6//V1db/paWj/zw7N/8qKif/JCQh/yAgHvUaGhomAAAABB4eG8weHhz/Hh8d/x0e
HP9ERkP/qamp/9HR0f/Lysv/ra2t/3JwbP9ZVlH/SUdD/zs5Nv8tLSrMAAAABAAAAAAdHRpoHh8d/x8hH/8iJCH/HB0a/ystK/9SU1P/UlJU/zExMP9APjv/W1hU/1pYUv9SUEv/Ojo4aAAAAAAAAAAAMzMzBR0g
HbwgIiD/Hh8d/x4fHv8dIB7/GRsb/xkZG/8eHR7/LCwq/0RDQP9ZV1L/VFJNvDMzAAUAAAAA////AQAAAAAWFhYXHB4bwx0eHf8dHx3/HSAf/x0fH/8dHh//Hh4e/yIiIf82NTL/R0ZCxE1CQhcAAAAA////Af//
/wH///8BAAAAABwcHAkcHhx/Gx8c5R4hH/8dICD/HR8g/x8fH/8fHx7lKioofxwcHAkAAAAA////Af///wH///8B////Af///wH///8BAAAAABwcHBIgICBHHB4eaxweHmsgICBHHBwcEgAAAAAAAAAA////Af//
/wH///8BDCCsQRAIrEEgBKxBQAKsQYABrEGAAaxBAACsQQAArEEAAKxBAACsQQAArEGAAaxBgAGsQUACrEEgBKxBCBisQQ==
"@
#endregion ******** $Demo16Icon ********
$PILSmallImageList.Images.Add("Demo16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Demo16Icon))))

#region ******** $AddCol16Icon ********
$AddCol16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAA4AAAAWAAAAFcAAAA2AAAABQAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAADAEHAoYOSx3tGoIu/yCaM/8imC3/HXwi/xFFEesBBQGDAAAACwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIwczF9ods03/JMVL/yu7N/8tti3/LrUr/y62LP8tty//KJ8m/wsr
CdcAAAAhAAAAAAAAAAAAAAAAAAAAEQc5HN0cy2H/JsJF/y62LP83uTT/Zshi/2bIYv83uTT/LrYr/y23Lv8tsCr/DC8K2wAAAA8AAAAAAAAAAAENB5QXw2f/I8dQ/yq7N/8quzf/TMZX////////////TMZX/yq7
N/8quzf/K7o0/yqlJ/8CCgKSAAAAAAAAAAwJYTf1GtVw/ybBRf8nwUT/J8FE/0nLYv///////////0nLYv8nwUT/J8FE/yfBRP8rujT/FFAT9AAAAAsAAABCDqBe/xzSaf8jx1H/k+Oo/7vtyP/G8NH/////////
///G8NH/u+3I/5PjqP8jx1H/J8FD/x+EJP8AAABCAAAAWg+5cf8a1W7/H81e/8Px1P/////////////////////////////////D8dT/H81e/yPHUP8jmi7/AAAAWgAAAFsPunL/F9t7/xvTav+D6LD/qe/J/7fx
0f///////////7fx0f+p78n/g+iw/xvTav8hylf/IZ00/wAAAFoAAABEEJ9a/xLijP8X2Xf/F9l3/xfZd/88343///////////88343/F9l3/xfZd/8X2Xf/IshT/xyLMf8AAABCAAAADQ9aJvUP55f/EuGJ/xTf
hP8U34T/OeSX////////////OeSX/xTfhP8U34T/Gddy/yLIU/8RVR31AAAADAAAAAADCgGVFsVs/w/nl/8R5JD/EeSQ/xjllP9A6qf/QOqn/xjllP8R5JD/Ftx//x/OX/8euE//AwoBlQAAAAAAAAAAAAAAEQsz
D94U2YH/D+eX/xDlk/8P5pT/D+eW/w/mlf8S4oz/F9p6/xvUa/8azmj/CzQR3gAAABEAAAAAAAAAAAAAAAAAAAAkCTES2RPCcP8O6Jn/EOaU/xLhiv8T34T/FN6C/xPfhv8UwGz/CDIU2gAAACQAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAwABwOFCFQw7Aeebf8Ex5T/A8iY/wagcf8HVjPsAAcEhgAAAAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAA3AAAAVwAAAFcAAAA3AAAABQAAAAAAAAAAAAAAAAAA
AAAAAAAA+B+sQeAHrEHAA6xBgAGsQYABrEEAAKxBAACsQQAArEEAAKxBAACsQQAArEGAAaxBgAGsQcADrEHgB6xB+B+sQQ==
"@
#endregion ******** $AddCol16Icon ********
$PILSmallImageList.Images.Add("AddCol16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($AddCol16Icon))))

#region ******** $RemoveCol16Icon ********
$RemoveCol16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAA4AAAAWAAAAFcAAAA2AAAABQAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAADAAAB4YAAFLtAACP/wAAqv8AAKn/AACK/wAATesAAAaDAAAACwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIwAAN9oAAMD/AADW/wAA0P8AAM3/AADM/wAAzf8AAM7/AACz/wAA
MdcAAAAhAAAAAAAAAAAAAAAAAAAAEQAAPN0AANn/AADU/wAAzf8AAMz/AADM/wAAzP8AAMz/AADM/wAAzf8AAMf/AAA12wAAAA8AAAAAAAAAAAAADZQAAM7/AADY/wAA0P8AAND/AADQ/wAA0P8AAND/AADQ/wAA
0P8AAND/AADP/wAAuf8AAAuSAAAAAAAAAAwAAGb1AADi/wAA1P8AANT/AADU/wAA1P8AANT/AADU/wAA1P8AANT/AADU/wAA1P8AAM//AABa9AAAAAsAAABCAACn/wAA4P8AANj/oaHw/7q69P+6uvT/urr0/7q6
9P+6uvT/urr0/6Gh8P8AANj/AADU/wAAlP8AAABCAAAAWgAAwP8AAOH/AADc/9zc+v/////////////////////////////////c3Pr/AADc/wAA2P8AAKv/AAAAWgAAAFsAAb//AADl/wAA4P+Tk/L/qan1/6mp
9f+pqfX/qan1/6mp9f+pqfX/k5Py/wAA4P8AANr/AAGr/wAAAFoAAABEAAOg/wAA6/8AAOT/AADk/wAA5P8AAOT/AADk/wAA5P8AAOT/AADk/wAA5P8AAOT/AADZ/wADkv8AAABCAAAADQAEVvUAAe7/AADq/wAA
6P8AAOj/AADo/wAA6P8AAOj/AADo/wAA6P8AAOj/AADj/wAB2f8ABFT1AAAADAAAAAAAAAmVAAfD/wAA7v8AAOz/AADs/wAA7P8AAOz/AADs/wAA7P8AAOz/AADn/wAA3f8AB7v/AAAJlQAAAAAAAAAAAAAAEQAB
L94ACtf/AADu/wAA7f8AAO3/AADu/wAA7f8AAOv/AADl/wAA4P8ACtD/AAEw3gAAABEAAAAAAAAAAAAAAAAAAAAkAAAt2QAJvf8ACuv/AAPt/wAA6v8AAOj/AAPo/wAK5v8ACb3/AAAv2gAAACQAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAwAAAeFAABR7AAGnP8ACsb/AArI/wAGnv8AAFPsAAAHhgAAAAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAA3AAAAVwAAAFcAAAA3AAAABQAAAAAAAAAAAAAAAAAA
AAAAAAAA+B+sQeAHrEHAA6xBgAGsQYABrEEAAKxBAACsQQAArEEAAKxBAACsQQAArEGAAaxBgAGsQcADrEHgB6xB+B+sQQ==
"@
#endregion ******** $RemoveCol16Icon ********
$PILSmallImageList.Images.Add("RemoveCol16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($RemoveCol16Icon))))

#region ******** $Trash16Icon ********
$Trash16Icon = @"
AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAALMTExLivr6/3r6+v9LGxsf2xsbH9sbGx/bGxsf2xsbH9r6+v9K+vr/fPz8+vAAAAGQAA
ABIAAAAAAAAAAAAAAADOzs7g4+Pj/+zs7P/s7Oz/7Ozs/+zs7P9iupT/7Ozs/+zs7P/b29v/1tbW2gAAAAAAAAAAAAAAAAAAAADV1dUExMTE9OLk4/87sIL/Jalz/x6lbv+Wzrf/G6Jq/x2ga/8knWX/19fX/9DQ
0PHV1dUEAAAAAAAAAAAAAAAA0dHRLcHBwf3H4Nf/Jq17/6bVw/+m1ML/2+bi/zaqff+w1sf/Gp9o/87Y0//Ly8v919fXKQAAAAAAAAAAAAAAAM/Pz3jExMT/6uzr/zq0h/8som7/2OXg/+zs7P/m6uj/PKx//zeo
fP/l5eX/xMTE/9vb22kAAAAAAAAAAAAAAADNzc2nzc3N/7vZzP8rom3/LrCB/+bq6P/s7Oz/6Ovq/4fKsP/j6Ob/6+vr/8LCwv/d3d2VAAAAAAAAAAAAAAAAysrKydfX1//s7Oz/7Ozs/7rd0f8usoL/w+DW/yan
cv9BrYD/7Ozs/+zs7P/Ly8v/2NjYuwAAAAAAAAAAAAAAAMfHx+Hf39//7Ozs/+zs7P/m6un/Kap3/6bXxP8srn7/PKl4/+zs7P/s7Oz/09PT/9PT09kAAAAAAAAAAAAAAACxsbHz1NTU/9nZ2f/Z2dn/2dnZ/5jM
uv8rp3T/n8y7/9nZ2f/Z2dn/2dnZ/8jIyP+6urruAAAAAAAAAABrrZJGg4qH/Yqdlf97lIn/e5SK/3yViv98lYr/fZaL/32Wi/99lov/e5SK/3uUif+ImJH/hIyJ/V+vjjwAAAAAI6p5/yOqef8jqnn/I6p5/yOq
ef8jqnn/I6p5/yOqef8jqnn/I6p5/yOqef8jqnn/I6p5/yOqef8jqnn/MKFxAiGxf/8WvYn/Fr2J/xa9if8WvYn/Fr2J/xa9if8WvYn/Fr2J/xa9if8WvYn/Fr2J/xa9if8WvYn/IbF//zChcQUhsX//QNep/xa9
if8WvYn/Fr2J/xa9if8WvYn/Fr2J/xa9if8WvYn/Fr2J/xa9if8WvYn/NtKj/yGxf/8AAAAAFr2I/4/52f8X6LT/F+i0/xfotP8X6LT/F+i0/xfotP8X6LT/F+i0/xfotP8/7MD/F+i0/4/52f8WvYj/AAAAADnb
qksWvYj/F+i0/ymQZf8loXH/JaFx/yWhcf8loXH/JaFx/yWhcf8loXH/DL6J/z/jsP8WvYj/OduqSwAAAAAAAAAAOduqQBDRl/8Q0Zf/ENGX/xDRl/8Q0Zf/ENGX/xDRl/8Q0Zf/ENGX/xDRl/8Q0Zf/OduqQAAA
AAAAAAAAAAGsQcAHrEGAA6xBgAOsQYADrEGAA6xBgAOsQYADrEGAA6xBAAGsQQAArEEAAKxBAAGsQQABrEEAAaxBgAOsQQ==
"@
#endregion ******** $Trash16Icon ********
$PILSmallImageList.Images.Add("Trash16Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Trash16Icon))))

#endregion ******** PIL Small Image Icons ********


# ************************************************
# PILLarge ImageList
# ************************************************
#region $PILLargeImageList = [System.Windows.Forms.ImageList]::New()
$PILLargeImageList = [System.Windows.Forms.ImageList]::New($PILFormComponents)
$PILLargeImageList.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
$PILLargeImageList.ImageSize = [System.Drawing.Size]::New(48, 48)
#endregion $PILLargeImageList = [System.Windows.Forms.ImageList]::New()

#region ******** PIL Large ImageList Icons ********

#region ******** $Play48Icon ********
$Play48Icon = @"
AAABAAEAMDAAAAEAIACoJQAAFgAAACgAAAAwAAAAYAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAOzs7Abq6uov5eXlXunp6YTk5OWr4uLjvOTk5cjk5OTJ4eHivOPj46vn5+eE4uLjXujo6C/q6uoGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7u7gjq6uo85OTklenp6try8vL29/f3/vr6+/78/Pz//Pz8//z8/P79/f3//Pz8//z8
/P75+fn/9fX2/u7u7/bm5uba4N/glefm5zzr6+sIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAADu7u4E5ubmWOfn6Mf09PT9+/v7/v39/f79/f3+/f39/v39/f79/f3+/Pz7/vj49v74+Pb+/Pz7/v39/f79/f3+/f39/v39/f79/f3++fn5/u/v7/3h4eLH4eHhWOrq6gQAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7+/vAunp6ULn5+fS9vb3//39/f7+/v7//f39//f29f7u7Oj/5eHc/9zY
0f7Y08v/2NPL/9jTy/7Y08v/2NPL/9jTy/7c2NH/5eHc/+7r6P739vX//f39//7+/v/8/Pz+8vLy/9/f4NLk4+RC6+rqAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADs7OwR5OTklvLy8/f9/f3+/v7+//z7+/7w7uv/4d3W/9nUzf7Y08v/2NPL/9nTy/7Z1Mv/2tTL/9rUy/7a1Mv/2dTL/9nUy/7Z08v/2NPL/9jTyv7Z1M3/4d3W//Du
6//8+/v+/f39//v7+//r6+z33Nzclufm5hEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOXl5SPp6erM+fn5/v39/f79/f3+8O7r/t3Y
0v7Y0sr+2NPL/tnUy/7a1Mz+29XN/tvWzf7b1s7+29bO/tvWzv7b1s7+29bO/tvWzv7b1s3+29XM/trUy/7Z08v+2NPK/tfSyv7d2NL+8O7r/v39/f79/f3+9PT0/uDg4cze3t4jAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5+fnOuzs7OT8/Pz//v7+//n49//g3Nb+19HK/9jTyv7a1Mv/29XN/9vWzv7c18//3NfP/9zXz/7d2ND/3djQ/93Y0P7d2ND/3djQ/93Y
0P7c18//3NfP/9vWzv7b1s3/29XM/9rUy//Y0sr+19HK/+Dc1f/5+Pb+/v7+//n5+f/i4uLk39/fOgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADn5+c67u7u6Pz8
/P79/f3/8e/t/9rVzv/X0cr+2tTL/9vVzf7b1s7/3NfP/93Y0P7e2dH/3tnS/97Z0v7f2tP/39rT/9/a0/7f2tP/39rT/9/a0/7e2dL/3tnR/93Y0f7c19D/3NfP/9vWzv/a1cz+2dPL/9fRyf/a1c7+8e/s//39
/f/5+fn+4+Pk6N/e3zoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOXl5SPs7Ozk/Pz8/v39/f7u7Oj+19LK/tjSyv7a1Mz+29bO/tzXz/7d2ND+3tnR/t/a0/7f2tT+4NvV/uDb
1f7g3NX+4NzV/uHc1v7h3Nb+4NzV/uDc1f7g29X+39vU/t/a0/7e2dL+3djR/tzX0P7b1s7+29XN/trUy/7Y0sn+1tHJ/u7s6P79/f3++fn5/uDg4eTd3N0jAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA7OzsEenp6sz8/Pz//f39/+7s6P7V0Mf/2NLJ/9vVzP/b1s7+3djQ/97Z0f7f2tP/39vU/+Db1f7h3db/4t3X/+Ld1/7j3tj/497Y/+Pe2P7j3tj/497Y/+Le1/7i3df/4d3W/+Dc1f7g29X/39rU/97Z
0v/d2ND+3NfP/9vWzv/a1Mv+19HJ/9XQx//u7Oj+/f39//j4+P/e3t7M5eTkEQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADv7+8C5OTklvn5+f7+/v7/8e/s/9bRyf7X0sn/29XM/9zXz//d2ND+3tnS/9/a
1P7g29X/4d3W/+Ld1/7j3tj/49/Y/+Pf2f7k4Nr/5ODa/+Tg2/7k4Nv/5ODa/+Tg2v7j39n/49/Y/+Le1/7i3df/4NzV/+Db1f/f2tP+3tnR/9zXz//b1s7+2tTM/9fRyf/W0cj+8e/s//7+/v/x8fL+2NjZlujo
6AIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADp6elC8vLz9/39/f74+Pb/2dTM/9fRyP7b1cz/3NfP/93Y0P/e2dL+39vU/+Dc1f7i3df/497Y/+Pf2f7k4Nr/5eDb/+Xh3P7m4dz/5uLd/+bi3f7m4t3/5uLd/+bh
3P7l4dz/5ODb/+Tg2v7j39j/4t7X/+Hd1v/g3NX+39rU/97Z0f/c18/+29bO/9rUy//W0cj+2NTL//j49v/8/Pz+5+fo9+Dg4EIAAAAAAAAAAAAAAAAAAAAAAAAAAO7u7gTn5+fS/f39/v39/f7e2tP+1dDH/trU
y/7b1s7+3djQ/t7Z0v7f29T+4NzV/uLd1/7j39j+5ODa/uXg2/7m4dz+5+Ld/ufj3v7n5N/+6OTf/ujk3/7o5N/+6OTf/ufj3v7n497+5uLd/ubh3P7k4Nv+5ODZ/uPe2P7h3db+4NzV/t/a1P7e2dH+3NfP/tvW
zv7Z08v+1c/H/t7Z0v79/f3++Pj4/tra2tLm5uUEAAAAAAAAAAAAAAAAAAAAAObm5lj29vf//v7+/+/t6f7UzsT/2dPK/9vWzv7d2ND/3tnS/9/b1P/g3NX+4t3X/+Pf2f7k4Nv/5uHc/+fi3f7n5N//6OXg/+jl
4P7p5uH/6ubi/+rm4v7q5uL/6ebh/+nl4f7o5eD/6OTf/+fj3v7m4t3/5eHc/+Tg2v/j39j+4t3X/+Dc1f/f2tP+3tnR/9zXz//b1c3+2NLK/9TOxP/v7en+/f39/+zs7P/b29tYAAAAAAAAAAAAAAAA7u7uCOfn
6Mf9/f3++/v7/tvVzf7W0Mf+29XN/tzXz/7e2dH+39rU/uDc1f7i3df+49/Z/uTg2/7m4dz+5+Pe/tbSzP7Lx8H+6ebh/urn4/7r6OT+7Onl/uzp5f7s6eX+7Onl/uvo5P7q5+P+6ebh/ujl4P7n5N/+5+Ld/uXh
3P7k4Nr+497Y/uHd1v7g29X+39rT/t3Y0P7b1s7+2tTM/tbQx/7b1c3++/v7/vn5+f7a2trH5ubmCAAAAAAAAAAA6urqPPT09P3+/v7/7+zo/9TOxP7a1Mv/29bO/93Y0P7f2tP/4NvV/+Ld1//j39j+5ODb/+bh
3P7n497/6OXg/8vHwf5tZ13/lpGJ/9bTzv7u6+f/7uvn/+7r5/7u6+f/7uvn/+3r5/7t6ub/7Ojk/+rn4/7p5uH/6OTf/+fj3v/l4dz+5ODa/+Pe2P/h3Nb+39vU/97Z0v/c18/+29bN/9nTyv/UzsT+7uzo//7+
/v/n5+f94ODgOwAAAAAAAAAA5OTklfv7+/79/f3/3tnR/9bQxv7b1c3/3NfP/97Z0f7f29T/4d3W/+Pe2P/k4Nr+5uHc/+fj3v7o5eD/6ubi/8zIwv5tZ13/bWdd/3NtZP6oo5z/4uDb/+/t6v7v7er/7+3p/+/t
6f7u7Oj/7uvn/+zp5f7r5+P/6ebh/+jk3//n4t3+5eHc/+Pf2f/i3df+4NzV/9/a0//d2ND+3NbO/9rVzP/Vz8b+3djR//39/f/z8/P+19bXlQAAAADs7OwG6enq2/39/f739fT+1c7F/tnTyv7e2tL+3djQ/t/a
0/7g29X+4t3X/uPf2f7l4Nv+5+Ld/ujk3/7p5uH+6+jk/s7KxP5sZlz+bGZc/mxmXP5sZlz+e3Vs/rezrf7v7er+8e/s/vDu6/7v7en+7uzo/u3r5/7t6ub+6ufj/unl4f7n497+5uHc/uTg2/7j3tj+4d3W/t/b
1P7e2dH+3NfP/t7Z0f7Y0sn+1M7E/vb19P76+vr+29vc2+Pj4wbq6uov8vLy9v39/f7s6eX/08zD/93Xz/7i3tj/39rS/9/a1P7h3db/497Y/+Tg2v/m4dz+5+Tf/+nl4f7r5+P/7erm/87Lxf5tZ13/bWdd/2xm
XP5tZ13/bWdd/21nXf6Ef3b/ycbA/+7s6f7x7+z/8O7q/+/t6f7u6+f/7Onl/+rm4v/o5eD+5+Pe/+Xh3P/j39n+4t3X/+Dc1f/f2tP+3tnR/+Le1//c1s7+08zD/+zp5f/9/f3+5OTl9uDg4C/l5eVe9/f3/v39
/f7i3df/1c7E/9/a0/7j39n/4t7X/+Db1f7i3df/49/Y/+Xg2//n4t3+6OXg/+rm4v7s6eX/7evn/83KxP5tZ13/bWdd/2xmXP5tZ13/bmhe/3FqYP50bWP/eHFm/5ONg/7Iw7v/39vV/+bj3v7s6eX/7erm/+vo
5P/p5uH+5+Tf/+bh3P/k4Nr+497Y/+Hc1v/f2tT+4d3W/+Pf2f/f2dL+1c7E/+Ld1v/9/f3+7Ozs/tjY2V7p6emF+vr7/v39/f7Y0sr+1tDH/uHc1f7j39n+5eHb/uHd1/7i3df+49/Z/uXh3P7n497+6OXg/urn
4/7n5N/+3tnT/r64sP5sZlz+bGZc/mxmXP5tZ13+cWth/nVvZP56c2f+fXVp/n12av6AeGz+npeN/szGvv7Y0sv+3tnT/ubj3v7q5uL+6OTf/ufi3f7k4Nv+49/Y/uHd1v7h3Nb+5ODa/uPf2f7g3NT+1s/G/tjS
yf79/f3+8vLy/t3d3YXk5OWq/Pz8//39/f7TzMP/2NLJ/+Hd1/7k4Nr/5uLc/+bi3P7j3tj/5ODa/+bh3P/n5N/+5eHc/9vXz/7V0Mf/1dDH/7y2rv5tZ13/bWdd/2xmXP5vaV//dW5j/3pzZ/6Ad2v/g3pt/4R7
bv6Dem3/f3dr/353a/6noJf/zci//9XQx//b1s/+5eHb/+fj3v/l4dz+49/Z/+Le1//l4dz+5eHb/+Pf2f/h3db+19HJ/9PMwv/9/f3+9fX1/9bW16ri4uO7/Pz8//z7+/7RysD/2dPK/+Le2P7k4Nr/5uLc/+fj
3v7n493/5eHb/+bh3P/e2tP+1tHI/9TPxv7V0Mf/1dDH/7u2rf5tZ13/bWdd/21nXf5ya2H/eHFl/352av6Ee27/iH9y/4qBc/6If3H/g3pt/352av53cGX/f3lv/62onv/RzMP+1dDH/97Z0v/l4Nv+5ODa/+bi
3f/m497+5eHb/+Pf2f/i3df+2NLK/8/Ivv/8+/v+9/f3/9PT1Lvk5OXI/Pz8/vf39f7Nxrz+2dPK/uLe2P7k4Nr+5uLc/ufj3v7p5eD+5+Pf/trVzf7TzcT+1M7F/tTPxv7Uz8b+1M/G/ru2rf5sZlz+bGZc/m1n
Xf5zbGL+eXJm/oB3a/6GfXD+jIN0/o+Fdv6LgnT+hXxv/n93a/54cWb+cmth/m1nXf6Ff3X+ubOq/tPNw/7a1c3+5uPd/ujk3/7m497+5eHb/uPf2f7i3tf+2NLJ/szEuv739vX+9/f4/tXV1sjk5OTI/f39//f2
9f7Lw7n/2NLK/+Le2P7k4Nr/5uLc/+fj3v7o5N7/4NvU/9rVzf/Vz8X+1M7E/9POxP7Uz8X/1M/G/7u2rf5tZ13/bWdd/21nXf5ybGL/eXJm/4B3a/6GfW//i4J0/42Edf6LgnP/hXxv/393a/54cWX/cmth/21n
Xf+EfnX+ubOp/9XPxf/a1c3+4NvU/+fj3v/m497+5uLb/+Pf2f/i3tf+2NHJ/8nBtv/39vX+9/f4/9TU1cjh4eK7/Pz8//z7+v7Hv7T/19HI/+Le1/7k4Nr/5uLc/+bj3v7f2tP/3NfQ/9zY0P/b1s7+1tHH/9TO
xP7UzsT/1M7F/7u1rP5tZ13/bWdd/21nXf5xa2H/d3Bk/311af6CeW3/hn1w/4h/cf6GfXD/gnls/311af52b2T/f3hu/6ymnf/QysD+1tHH/9vWzv/c19D+3NfP/9/a0//m4t3+5eHb/+Pf2f/i3df+1s/H/8W9
sv/7+/r+9vb2/9LS07vj4+Oq/Pz8/v39/f7FvbH+1M3D/uHd1v7k4Nr+5eHb/uDb1f7b1s7+29bP/tzXz/7c19D+3NfQ/tnUzP7Vz8X+1M7E/rq0qv5sZlz+bGZc/mxmXP5uaF7+dG1j/nlxZv5+dmr+gXhs/oJ5
bf6BeGz+fXVp/nx1av6lnpT+zMa8/tXPxf7Z1Mz+3NfQ/tzX0P7b1s/+29bO/tvWzv7g29T+5eHb/uPf2f7h3db+0szC/sO7r/79/f3+9PT1/tTU1arn5+eF+fn5//39/f7Kwrj/zse9/+Dc1f7j39n/4dzW/9zW
z/7c18//3NfP/9zXz//c18/+3NfQ/9zX0P7c19D/2dTM/7y3rv5tZ13/bWdd/2xmXP5tZ13/cGlf/3RtY/54cWX/enNn/3tzaP59dmr/nJWK/8rEuv7W0Mf/2dTM/9zX0P/c19D+3NfP/9zXz//b1s7+3NfP/9zX
z//b1s7+4NvV/+Pf2f/g29T+zca8/8jBt//9/f3+8PDw/9vb24Xi4uNe9fX2/v39/f7Uz8f/x7+0/9/a0v7h3df/2NLL/9nUzP7b1s7/3NfP/9zXz//b1s7+3NfP/9zXz/7c19D/3NfQ/8G8tP5tZ13/bWdd/2xm
XP5tZ13/bWdd/29pX/5ya2H/dW9k/5CJgP7EvrX/2NPL/9vWzv7c19D/3NfQ/9zX0P/c18/+3NfP/9zXz//b1s7+3NfP/9vWzv/Y0sv+19HK/+Hd1v/e2dH+xr6z/9PNxv/9/f3+6enq/tbW1l7o6Ogv7u7v9v39
/f7j39r+v7aq/tvVzf7Y08v+083E/tbQyf7Y08v+29bO/tvWzv7b1s7+29bO/tvWzv7b1s7+3NfP/sG7s/5sZlz+bGZc/mxmXP5sZlz+bGZc/mxmXP5/eXD+ubSs/trVzv7c19D+3NfQ/tzX0P7c18/+3NfP/tvW
zv7b1s7+29bO/tvWzv7b1s7+2tXN/tfSyv7Vz8f+0szC/tjSyf7a1Mz+vbWo/uPf2f78/Pz+4eHh9t3d3S/q6uoG5ubm2/39/f7z8u//vLOn/87GvP7Mxbv/z8m//9LMw/7Vz8f/19LK/9rUzf/b1s7+3NfP/9vW
zv7c18//3NfP/8C7s/5tZ13/bWdd/2xmXP5tZ13/eXNp/6umnf7a1c3/3NfP/9zXz/7c18//3NfP/9vWzv7c18//3NfP/9zXz//b1s7+3NfP/9vWzv/Z08z+1tHJ/9TOxv/Ry8H+zse9/8vDuf/Nxbv+urKl//Px
7//4+Pj+19fY2+Dg4AYAAAAA39/glfn5+f79/f3/yMG3/7mwo/7Du6//zMS7/87Ivf7Ry8H/1M7F/9bQyf/Y0sv+2tXN/9vWzv7c18//3NfP/8C7s/5tZ13/bWdd/3JsYv6fmpD/0MvD/9vWzv7c18//3NfP/9vW
zv7c18//3NfP/9vWzv7c18//3NfP/9zXz//b1s7+2dTM/9fSyv/Vz8f+083D/9DKwP/Oxrz+y8O6/8K6rv+4r6L+x8C1//39/f/v7+/+0tLSlQAAAAAAAAAA5ubmO+/v7/3+/v7/5ODb/7OqnP65sKT/x7+1/8vD
uf7Oxr3/0Mm//9LMwv/Uzsb+1tDJ/9jSyv7Z1M3/29bO/8C7s/5tZ13/kYuC/8jDuv7c18//3NfP/9vWzv7c18//3NfP/9vWzv7c18//3NfP/9vWzv7b1s7/2tXO/9nUzP/X0cr+1tDI/9TOxf/Ry8H+z8i+/83F
u//Kwrj+x760/7ivov+yqJr+4+Db//39/f/h4eL929vbOwAAAAAAAAAA6urqCOHh4sf8/Pz++vr5/r61qv6yqZr+vbWp/se/tf7Jwrj+zMS6/s7Hvf7Qyb/+0szC/tTOxf7V0Mj+1tHJ/sjCu/6/ubH+2dTM/trV
zf7b1s7+29bO/tvWzv7b1s7+29bO/tvVzv7a1c3+2dTM/tjTy/7X0sr+1tDJ/tXPx/7TzcT+0cvB/s/Ivv7Nxrz+y8O5/snAt/7GvrT+vLSo/rGnmf68tKj++vr5/vX19f7T09TH4eHhCAAAAAAAAAAAAAAAAODg
4Fjy8vL//f39/+Pg2/6to5T/sqmb/8C4rf7FvrP/yMC2/8rCuP/MxLv+zsa9/8/Ivv7Ry8D/0szD/9TOxf7Vz8f/1tDI/9bQyf7X0cn/19HK/9fSyv7X0sr/19HK/9fRyf7W0Mn/1c/I/9TPxv7TzcT/0szC/9DK
wP/PyL7+zca8/8vEuv/Jwbf+x7+1/8S9sv+/t6z+saeZ/6uik//i39r+/Pz8/+Xl5v/U1NRYAAAAAAAAAAAAAAAAAAAAAOno6ATf3+DS+/v7//39/f7Aua7/qqGS/7OqnP7Bua7/w7yx/8a+tP/IwLb+ysK4/8vD
uv7Nxrz/zse9/8/Ivv7QysD/0cvB/9LLwv7SzMP/0szD/9PNw/7TzcP/0szD/9LMwv7Ry8H/0crA/9DJv/7Px77/zsa8/8zFu//Lw7n+ycG3/8e/tf/FvbP+w7uw/8C4rf+xqJr+qZ+Q/7+3rP/9/f3+8/Pz/9LS
09Lf3t4EAAAAAAAAAAAAAAAAAAAAAAAAAADi4eFC6+vs9/39/f7z8vD+sKeZ/qeej/6zqpz+vrar/sG5rv7Du7H+xb2z/se/tP7IwLb+ysK4/svDuf7LxLr+zcW7/s3GvP7Nxrz+zse9/s7Hvf7Ox73+zce9/s3G
vP7Nxbv+zMS7/svDuv7Kwrn+ycG3/sjAtf7GvrT+xL2y/sO7sP7AuK3+vbWp/rGom/6mnI3+rqaY/vPy8P76+vr+39/g99fX10IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADp6egC29vclvT09P7+/v7/5OHc/6ed
j/6lm4z/saia/7u0p/++tqv+wLmt/8K7sP7EvLL/xb2z/8a+tP7Hv7X/yMC2/8nBt/7Jwbf/ysK4/8rCuP7Kwrj/ycK4/8nBt/7Jwbf/yMC2/8e/tf7GvrT/xL2y/8O8sf/Buq/+wLit/722qv+6sqb+r6eZ/6Sa
i/+mnI7+5ODc//39/f/p6er+z8/QluDf3wIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5eTkEeDg4Mz5+fn//f39/9zY0/6hl4f/o5mK/6yjlv+4sKT+u7On/721qf6/t6v/wLit/8G6rv7Cu7D/w7yx/8S8
sv7EvbL/xb2y/8W+s/7FvrP/xL2y/8S9sv7EvLL/w7uw/8K6r/7Bua7/wLis/762qv+8tKj+urKm/7evo/+so5X+opiJ/5+Vhv/c2NL+/f39//Hx8f/U1NTM29raEQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAN3c3CPi4uLk+fn5/v39/f7b19L+oZiJ/p+Vhf6mnY/+s6qe/revo/64saX+urKm/ru0qP69tan+vraq/r+3q/6/t6z+v7is/r+4rP6/uKz+v7es/r+3q/6+tqv+vbaq/ry0qP67s6f+urKm/riw
pP62rqL+sqqd/qacjv6elIT+oJeI/tvX0f79/f3+8vLz/tbW1+TT0tIjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADd3d064+Pk6Pn5+f79/f3/4d7a/6adj/+bkYH+n5aG/6qh
k/6zq57/ta2h/7auov63r6P/uLCk/7iwpP65saX/ubKm/7mypv66sqb/ubGl/7mxpf64sKT/t7Ck/7auov62rqL/ta2g/7Kqnv+poJL+n5WF/5qQgP+lnI7+4d7Z//39/f/y8vP+2dnZ6NPT0joAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA3NzcOuDg4eT4+Pj//v7+//Lx7/+0raH+mI5+/5mPf/6elYX/p5+R/6+nm/6yqp7/s6uf/7Orn/60rKD/tKyg/7WtoP61raD/tKyg/7Ss
n/6zq5//s6ue/7Kqnf6vp5r/p56Q/56Uhf+Yjn7+l419/7SsoP/y8e/+/f39//Hx8f/W1tfk09LSOgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANva
2iPd3d3M8fHy/vz8/P79/f3+3NnU/qmhlP6Vi3v+lox8/piOfv6elYb+pJyO/qmhk/6spJf+rqWZ/q6mmv6vppr+rqWZ/qyjl/6poJP+pJuN/p6Vhv6Xjn7+lox8/pWLe/6poJP+3NnT/v39/f76+vr+6enq/tTU
1MzS0tIjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADi4eER2NfYlufn6Pf4+Pj+/f39//n5+P7a19H/sKmd/5iPgP6TiXn/lIp6/5WL
e/6WjHz/l41+/5iOfv6Yjn//l419/5aMfP6Vi3v/lIp6/5OJef6Yj4D/sKid/9rX0f/5+fj+/Pz8//Pz8//f3+D3z8/QltrZ2REAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5eXlAt3d3ELa2trS7Ozs/vn5+f7+/v7//f39/+7t6v7V0cv/urSq/6WdkP6WjH7/koh5/5KIef6SiHn/koh5/5aMfv6lnZD/urSq/9XRy/7u7ev//f39//39
/f/19fX+5eXm/tLS09LW1tVC397eAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADj4+IE2dnZWNnZ
2sfn5+f98/Pz/vr6+v79/f3+/f39/v39/f79/f3++vr5/vPy8P7z8vD++vr5/v39/f79/f3+/f39/vz8/P74+Pj+7+/v/uHh4v3T09TH09PTWN7d3QQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOTk5Aje3d071tbWldvb3Nrk5OX27Ozs/vLy8v719fX/9/f3//f3+P739/j/9vb2//T0
9f7w8PD/6enq/uHh4fbX19ja0dHSldnZ2Tvg398IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOLh4gbf3t4v2NfYXtzb24TV1dar09PUu9XV1sjU1NXI0tLTu9TU1Kva2dmE1dXVXtzb2y/f398GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//4AB//8AAP/8AAA//wAA//AAAA//AAD/wAAAA/8AAP+AAAAB/wAA/wAAAAD/AAD+AAAAAH8AAPwAAAAAPwAA+AAAAAAfAADwAAAAAA8AAOAA
AAAABwAA4AAAAAAHAADAAAAAAAMAAMAAAAAAAwAAgAAAAAABAACAAAAAAAEAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAQAAgAAAAAABAACAAAAAAAEAAMAAAAAAAwAAwAAAAAADAADgAAAAAAcAAOAAAAAABwAA8AAAAAAPAAD4AAAAAB8AAPwA
AAAAPwAA/gAAAAB/AAD/AAAAAP8AAP+AAAAB/wAA/8AAAAP/AAD/8AAAD/8AAP/8AAA//wAA//+AAf//AAA=
"@
#endregion ******** $Play48Icon ********
$PILLargeImageList.Images.Add("Play48Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Play48Icon))))

#region ******** $Pause48Icon ********
$Pause48Icon = @"
AAABAAEAMDAAAAEAIACoJQAAFgAAACgAAAAwAAAAYAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAOzs7Abq6uov5eXlXunp6YTk5OWr4uLjvOTk5cjk5OTJ4eHivOPj46vn5+eE4uLjXujo6C/q6uoGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7u7gjq6uo85OTklenp6try8vL29/f3/vr6+/78/Pz//Pz8//z8/P79/f3//Pz8//z8
/P75+fn/9fX2/u7u7/bm5uba4N/glefm5zzr6+sIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAADu7u4E5ubmWOfn6Mf09PT9+/v7/v39/f79/f3+/f39/v39/f79/f3+/Pz7/vj49v74+Pb+/Pz7/v39/f79/f3+/f39/v39/f79/f3++fn5/u/v7/3h4eLH4eHhWOrq6gQAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7+/vAunp6ULn5+fS9vb3//39/f7+/v7//f39//f29f7u7Oj/5eHc/9zY
0f7Y08v/2NPL/9jTy/7Y08v/2NPL/9jTy/7c2NH/5eHc/+7r6P739vX//f39//7+/v/8/Pz+8vLy/9/f4NLk4+RC6+rqAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADs7OwR5OTklvLy8/f9/f3+/v7+//z7+/7w7uv/4d3W/9nUzf7Y08v/2NPL/9nTy/7Z1Mv/2tTL/9rUy/7a1Mv/2dTL/9nUy/7Z08v/2NPL/9jTyv7Z1M3/4d3W//Du
6//8+/v+/f39//v7+//r6+z33Nzclufm5hEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOXl5SPp6erM+fn5/v39/f79/f3+8O7r/t3Y
0v7Y0sr+2NPL/tnUy/7a1Mz+29XN/tvWzf7b1s7+29bO/tvWzv7b1s7+29bO/tvWzv7b1s3+29XM/trUy/7Z08v+2NPK/tfSyv7d2NL+8O7r/v39/f79/f3+9PT0/uDg4cze3t4jAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5+fnOuzs7OT8/Pz//v7+//n49//g3Nb+19HK/9jTyv7a1Mv/29XN/9vWzv7c18//3NfP/9zXz/7d2ND/3djQ/93Y0P7d2ND/3djQ/93Y
0P7c18//3NfP/9vWzv7b1s3/29XM/9rUy//Y0sr+19HK/+Dc1f/5+Pb+/v7+//n5+f/i4uLk39/fOgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADn5+c67u7u6Pz8
/P79/f3/8e/t/9rVzv/X0cr+2tTL/9vVzf7b1s7/3NfP/93Y0P7e2dH/3tnS/97Z0v7f2tP/39rT/9/a0/7f2tP/39rT/9/a0/7e2dL/3tnR/93Y0f7c19D/3NfP/9vWzv/a1cz+2dPL/9fRyf/a1c7+8e/s//39
/f/5+fn+4+Pk6N/e3zoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOXl5SPs7Ozk/Pz8/v39/f7u7Oj+19LK/tjSyv7a1Mz+29bO/tzXz/7d2ND+3tnR/t/a0/7f2tT+4NvV/uDb
1f7g3NX+4NzV/uHc1v7h3Nb+4NzV/uDc1f7g29X+39vU/t/a0/7e2dL+3djR/tzX0P7b1s7+29XN/trUy/7Y0sn+1tHJ/u7s6P79/f3++fn5/uDg4eTd3N0jAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA7OzsEenp6sz8/Pz//f39/+7s6P7V0Mf/2NLJ/9vVzP/b1s7+3djQ/97Z0f7f2tP/39vU/+Db1f7h3db/4t3X/+Ld1/7j3tj/497Y/+Pe2P7j3tj/497Y/+Le1/7i3df/4d3W/+Dc1f7g29X/39rU/97Z
0v/d2ND+3NfP/9vWzv/a1Mv+19HJ/9XQx//u7Oj+/f39//j4+P/e3t7M5eTkEQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADv7+8C5OTklvn5+f7+/v7/8e/s/9bRyf7X0sn/29XM/9zXz//d2ND+3tnS/9/a
1P7g29X/4d3W/+Ld1/7j3tj/49/Y/+Pf2f7k4Nr/5ODa/+Tg2/7k4Nv/5ODa/+Tg2v7j39n/49/Y/+Le1/7i3df/4NzV/+Db1f/f2tP+3tnR/9zXz//b1s7+2tTM/9fRyf/W0cj+8e/s//7+/v/x8fL+2NjZlujo
6AIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADp6elC8vLz9/39/f74+Pb/2dTM/9fRyP7b1cz/3NfP/93Y0P/e2dL+39vU/+Dc1f7i3df/497Y/+Pf2f7k4Nr/5eDb/+Xh3P7m4dz/5uLd/+bi3f7m4t3/5uLd/+bh
3P7l4dz/5ODb/+Tg2v7j39j/4t7X/+Hd1v/g3NX+39rU/97Z0f/c18/+29bO/9rUy//W0cj+2NTL//j49v/8/Pz+5+fo9+Dg4EIAAAAAAAAAAAAAAAAAAAAAAAAAAO7u7gTn5+fS/f39/v39/f7e2tP+1dDH/trU
y/7b1s7+3djQ/t7Z0v7f29T+4NzV/uLd1/7j39j+5ODa/uXg2/7m4dz+5+Ld/ufj3v7n5N/+6OTf/ujk3/7o5N/+6OTf/ufj3v7n497+5uLd/ubh3P7k4Nv+5ODZ/uPe2P7h3db+4NzV/t/a1P7e2dH+3NfP/tvW
zv7Z08v+1c/H/t7Z0v79/f3++Pj4/tra2tLm5uUEAAAAAAAAAAAAAAAAAAAAAObm5lj29vf//v7+/+/t6f7UzsT/2dPK/9vWzv7d2ND/3tnS/9/b1P/g3NX+4t3X/+Pf2f7k4Nv/5uHc/+fi3f7n5N//6OXg/+jl
4P7p5uH/6ubi/+rm4v7q5uL/6ebh/+nl4f7o5eD/6OTf/+fj3v7m4t3/5eHc/+Tg2v/j39j+4t3X/+Dc1f/f2tP+3tnR/9zXz//b1c3+2NLK/9TOxP/v7en+/f39/+zs7P/b29tYAAAAAAAAAAAAAAAA7u7uCOfn
6Mf9/f3++/v7/tvVzf7W0Mf+29XN/tzXz/7e2dH+39rU/uDc1f7i3df+49/Z/uTg2/7m4dz+5+Pe/ujk3/7p5eH+6ubi/urn4/7r6OT+7Onl/uzp5f7s6eX+7Onl/uvo5P7q5+P+6ebh/ujl4P7n5N/+5+Ld/uXh
3P7k4Nr+497Y/uHd1v7g29X+39rT/t3Y0P7b1s7+2tTM/tbQx/7b1c3++/v7/vn5+f7a2trH5ubmCAAAAAAAAAAA6urqPPT09P3+/v7/7+zo/9TOxP7a1Mv/29bO/93Y0P7f2tP/4NvV/+Ld1//j39j+5ODb/+bh
3P7n497/0MzG/83Iw/7OycT/zsvF/8/Mxv7c2dX/7uvn/+7r5/7u6+f/7uvn/9zZ1P7Py8b/zsrF/83Jw/7MyML/0MvF/+fj3v/l4dz+5ODa/+Pe2P/h3Nb+39vU/97Z0v/c18/+29bN/9nTyv/UzsT+7uzo//7+
/v/n5+f94ODgOwAAAAAAAAAA5OTklfv7+/79/f3/3tnR/9bQxv7b1c3/3NfP/97Z0f7f29T/4d3W/+Pe2P/k4Nr+5uHc/+fj3v7o5eD/gXty/25oXv5uaF7/bmhe/25oXv6loZn/7+3p/+/t6v7v7er/7+3p/6Wh
mf5uaF7/bmhe/25oXv5uaF7/gXty/+jk3//n4t3+5eHc/+Pf2f/i3df+4NzV/9/a0//d2ND+3NbO/9rVzP/Vz8b+3djR//39/f/z8/P+19bXlQAAAADs7OwG6enq2/39/f739fT+1c7F/tnTyv7e2tL+3djQ/t/a
0/7g29X+4t3X/uPf2f7l4Nv+5+Ld/ujk3/7p5uH+gHpx/mxmXP5sZlz+bGZc/mxmXP6loZr+8e/s/vHv7P7x7+z+8e/s/qWhmf5sZlz+bGZc/mxmXP5sZlz+gHpx/unl4f7n497+5uHc/uTg2/7j3tj+4d3W/t/b
1P7e2dH+3NfP/t7Z0f7Y0sn+1M7E/vb19P76+vr+29vc2+Pj4wbq6uov8vLy9v39/f7s6eX/08zD/93Xz/7i3tj/39rS/9/a1P7h3db/497Y/+Tg2v/m4dz+5+Tf/+nl4f7r5+P/gHpy/2xmXP5tZ13/bWdd/2xm
XP6loJn/7+3q/+7s6f7u7On/7+3q/6Wgmf5tZ13/bWdd/2xmXP5tZ13/gHpx/+rm4v/o5eD+5+Pe/+Xh3P/j39n+4t3X/+Dc1f/f2tP+3tnR/+Le1//c1s7+08zD/+zp5f/9/f3+5OTl9uDg4C/l5eVe9/f3/v39
/f7i3df/1c7E/9/a0/7j39n/4t7X/+Db1f7i3df/49/Y/+Xg2//n4t3+6OXg/+rm4v7s6eX/gHty/2xmXP5tZ13/bWdd/2xmXP6blY3/2dTM/9jUzP7Z1Mz/2dTM/5uVjP5tZ13/bWdd/2xmXP5tZ13/gHty/+vo
5P/p5uH+5+Tf/+bh3P/k4Nr+497Y/+Hc1v/f2tT+4d3W/+Pf2f/f2dL+1c7E/+Ld1v/9/f3+7Ozs/tjY2V7p6emF+vr7/v39/f7Y0sr+1tDH/uHc1f7j39n+5eHb/uHd1/7i3df+49/Z/uXh3P7n497+6OXg/urn
4/7n5N/+f3lw/mxmXP5sZlz+bGZc/mxmXP6alIv+1tDJ/tbRyf7W0cn+1tDJ/pqUi/5sZlz+bGZc/mxmXP5sZlz+f3lw/ubj3v7q5uL+6OTf/ufi3f7k4Nv+49/Y/uHd1v7h3Nb+5ODa/uPf2f7g3NT+1s/G/tjS
yf79/f3+8vLy/t3d3YXk5OWq/Pz8//39/f7TzMP/2NLJ/+Hd1/7k4Nr/5uLc/+bi3P7j3tj/5ODa/+bh3P/n5N/+5eHc/9vXz/7V0Mf/fnhu/29oX/5vaV//b2he/25oXv6alIv/19HJ/9bQyf7X0cn/19HJ/5uU
i/5vaF7/b2lf/29oX/5uaF7/fXdu/9XQx//b1s/+5eHb/+fj3v/l4dz+49/Z/+Le1//l4dz+5eHb/+Pf2f/h3db+19HJ/9PMwv/9/f3+9fX1/9bW16ri4uO7/Pz8//z7+/7RysD/2dPK/+Le2P7k4Nr/5uLc/+fj
3v7n493/5eHb/+bh3P/e2tP+1tHI/9TPxv7V0Mf/gntx/3NsYf5zbGL/cmxh/3JrYf6clo3/19HJ/9bQyP7X0cn/19HJ/52Xjf5zbGH/c2xi/3NsYf5ybGH/gXpw/9XQx//Uz8b+1dDH/97Z0v/l4Nv+5ODa/+bi
3f/m497+5eHb/+Pf2f/i3df+2NLK/8/Ivv/8+/v+9/f3/9PT1Lvk5OXI/Pz8/vf39f7Nxrz+2dPK/uLe2P7k4Nr+5uLc/ufj3v7p5eD+5+Pf/trVzf7TzcT+1M7F/tTPxv7Uz8b+hH1z/nZvZP52b2T+dm9j/nRu
Y/6dl43+1dDH/tXQx/7V0Mf+1dDH/p6Xjf52b2T+dm9k/nZvZP51bmP+g31y/tTPxv7Uz8X+1M7E/tPNxP7a1c3+5uPd/ujk3/7m497+5eHb/uPf2f7i3tf+2NLJ/szEuv739vX+9/f4/tXV1sjk5OTI/f39//f2
9f7Lw7n/2NLK/+Le2P7k4Nr/5uLc/+fj3v7o5N7/4NvU/9rVzf/Vz8X+1M7E/9POxP7Uz8X/h4B1/3pyZ/56c2f/enJn/3hxZf6fmY7/1dDH/9TPxv7V0Mf/1dDH/6Caj/56cmf/enNn/3pyZ/55cmb/hn90/9TO
xf/TzcT+1M7E/9XPxf/a1c3+4NvU/+fj3v/m497+5uLb/+Pf2f/i3tf+2NHJ/8nBtv/39vX+9/f4/9TU1cjh4eK7/Pz8//z7+v7Hv7T/19HI/+Le1/7k4Nr/5uLc/+bj3v7f2tP/3NfQ/9zY0P/b1s7+1tHH/9TO
xP7UzsT/ioN3/352av5+dmr/fXZq/3x0af6hmpD/1dDH/9TPxv7V0Mf/1dDH/6Kbkf59dmr/fnZq/352av59dWn/iYF2/9TOxP/UzsT+1tHH/9vWzv/c19D+3NfP/9/a0//m4t3+5eHb/+Pf2f/i3df+1s/H/8W9
sv/7+/r+9vb2/9LS07vj4+Oq/Pz8/v39/f7FvbH+1M3D/uHd1v7k4Nr+5eHb/uDb1f7b1s7+29bP/tzXz/7c19D+3NfQ/tnUzP7Vz8X+jYV5/oF5bf6CeW3+gXhs/n93a/6im5H+1M/G/tTPxv7Uz8b+1M/G/qSd
kv6BeWz+gnlt/oF5bf6AeGv+jIR4/tXPxf7Z1Mz+3NfQ/tzX0P7b1s/+29bO/tvWzv7g29T+5eHb/uPf2f7h3db+0szC/sO7r/79/f3+9PT1/tTU1arn5+eF+fn5//39/f7Kwrj/zse9/+Dc1f7j39n/4dzW/9zW
z/7c18//3NfP/9zXz//c18/+3NfQ/9zX0P7c19D/kYl9/4V8b/6GfW//hHxv/4J6bf6knJH/1M7E/9PNxP7UzsT/1M7E/6Wek/6FfG//hn1v/4V8b/6De27/j4d7/9zX0P/c19D+3NfP/9zXz//b1s7+3NfP/9zX
z//b1s7+4NvV/+Pf2f/g29T+zca8/8jBt//9/f3+8PDw/9vb24Xi4uNe9fX2/v39/f7Uz8f/x7+0/9/a0v7h3df/2NLL/9nUzP7b1s7/3NfP/9zXz//b1s7+3NfP/9zXz/7c19D/lIx//4mAcv6JgHL/h35x/4V8
b/6mn5T/1dDG/9XPxv7Vz8b/1dDG/6iglf6If3H/iYBy/4mAcv6GfXD/kYl9/9zX0P/c18/+3NfP/9zXz//b1s7+3NfP/9vWzv/Y0sv+19HK/+Hd1v/e2dH+xr6z/9PNxv/9/f3+6enq/tbW1l7o6Ogv7u7v9v39
/f7j39r+v7aq/tvVzf7Y08v+083E/tbQyf7Y08v+29bO/tvWzv7b1s7+29bO/tvWzv7b1s7+lo6A/oyDdP6Ng3X+ioFz/od+cP6popj+3NfQ/tzXz/7c19D+3NfQ/qykmf6LgnP+jYN1/oyCdP6If3H+kop+/tvW
zv7b1s7+29bO/tvWzv7b1s7+2tXN/tfSyv7Vz8f+0szC/tjSyf7a1Mz+vbWo/uPf2f78/Pz+4eHh9t3d3S/q6uoG5ubm2/39/f7z8u//vLOn/87GvP7Mxbv/z8m//9LMw/7Vz8f/19LK/9rUzf/b1s7+3NfP/9vW
zv7c18//lo6A/42Ddf6OhXb/i4Jz/4d+cP6popj/3NfP/9zX0P7c19D/3NfP/6ykmf6LgnP/j4V2/42DdP6JgHL/k4p+/9zXz//b1s7+3NfP/9vWzv/Z08z+1tHJ/9TOxv/Ry8H+zse9/8vDuf/Nxbv+urKl//Px
7//4+Pj+19fY2+Dg4AYAAAAA39/glfn5+f79/f3/yMG3/7mwo/7Du6//zMS7/87Ivf7Ry8H/1M7F/9bQyf/Y0sv+2tXN/9vWzv7c18//lo2A/4uCdP6Mg3T/iYFy/4d+cP6popj/3NfP/9vWzv7c18//3NfP/6uk
mf6KgXP/jIN0/4uCc/6If3H/kop+/9zXz//b1s7+2dTM/9fSyv/Vz8f+083D/9DKwP/Oxrz+y8O6/8K6rv+4r6L+x8C1//39/f/v7+/+0tLSlQAAAAAAAAAA5ubmO+/v7/3+/v7/5ODb/7OqnP65sKT/x7+1/8vD
uf7Oxr3/0Mm//9LMwv/Uzsb+1tDJ/9jSyv7Z1M3/ysS7/8jCuf7Iw7n/yMK5/8fCuP7QysL/3NfP/9vWzv7c18//3NfP/9DLwv7Iwrn/yMO5/8jCuf7Iwrj/ycO7/9nUzP/X0cr+1tDI/9TOxf/Ry8H+z8i+/83F
u//Kwrj+x760/7ivov+yqJr+4+Db//39/f/h4eL929vbOwAAAAAAAAAA6urqCOHh4sf8/Pz++vr5/r61qv6yqZr+vbWp/se/tf7Jwrj+zMS6/s7Hvf7Qyb/+0szC/tTOxf7V0Mj+1tHJ/tjSy/7Z08z+2tTN/trV
zf7b1s7+29bO/tvWzv7b1s7+29bO/tvVzv7a1c3+2dTM/tjTy/7X0sr+1tDJ/tXPx/7TzcT+0cvB/s/Ivv7Nxrz+y8O5/snAt/7GvrT+vLSo/rGnmf68tKj++vr5/vX19f7T09TH4eHhCAAAAAAAAAAAAAAAAODg
4Fjy8vL//f39/+Pg2/6to5T/sqmb/8C4rf7FvrP/yMC2/8rCuP/MxLv+zsa9/8/Ivv7Ry8D/0szD/9TOxf7Vz8f/1tDI/9bQyf7X0cn/19HK/9fSyv7X0sr/19HK/9fRyf7W0Mn/1c/I/9TPxv7TzcT/0szC/9DK
wP/PyL7+zca8/8vEuv/Jwbf+x7+1/8S9sv+/t6z+saeZ/6uik//i39r+/Pz8/+Xl5v/U1NRYAAAAAAAAAAAAAAAAAAAAAOno6ATf3+DS+/v7//39/f7Aua7/qqGS/7OqnP7Bua7/w7yx/8a+tP/IwLb+ysK4/8vD
uv7Nxrz/zse9/8/Ivv7QysD/0cvB/9LLwv7SzMP/0szD/9PNw/7TzcP/0szD/9LMwv7Ry8H/0crA/9DJv/7Px77/zsa8/8zFu//Lw7n+ycG3/8e/tf/FvbP+w7uw/8C4rf+xqJr+qZ+Q/7+3rP/9/f3+8/Pz/9LS
09Lf3t4EAAAAAAAAAAAAAAAAAAAAAAAAAADi4eFC6+vs9/39/f7z8vD+sKeZ/qeej/6zqpz+vrar/sG5rv7Du7H+xb2z/se/tP7IwLb+ysK4/svDuf7LxLr+zcW7/s3GvP7Nxrz+zse9/s7Hvf7Ox73+zce9/s3G
vP7Nxbv+zMS7/svDuv7Kwrn+ycG3/sjAtf7GvrT+xL2y/sO7sP7AuK3+vbWp/rGom/6mnI3+rqaY/vPy8P76+vr+39/g99fX10IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADp6egC29vclvT09P7+/v7/5OHc/6ed
j/6lm4z/saia/7u0p/++tqv+wLmt/8K7sP7EvLL/xb2z/8a+tP7Hv7X/yMC2/8nBt/7Jwbf/ysK4/8rCuP7Kwrj/ycK4/8nBt/7Jwbf/yMC2/8e/tf7GvrT/xL2y/8O8sf/Buq/+wLit/722qv+6sqb+r6eZ/6Sa
i/+mnI7+5ODc//39/f/p6er+z8/QluDf3wIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5eTkEeDg4Mz5+fn//f39/9zY0/6hl4f/o5mK/6yjlv+4sKT+u7On/721qf6/t6v/wLit/8G6rv7Cu7D/w7yx/8S8
sv7EvbL/xb2y/8W+s/7FvrP/xL2y/8S9sv7EvLL/w7uw/8K6r/7Bua7/wLis/762qv+8tKj+urKm/7evo/+so5X+opiJ/5+Vhv/c2NL+/f39//Hx8f/U1NTM29raEQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAN3c3CPi4uLk+fn5/v39/f7b19L+oZiJ/p+Vhf6mnY/+s6qe/revo/64saX+urKm/ru0qP69tan+vraq/r+3q/6/t6z+v7is/r+4rP6/uKz+v7es/r+3q/6+tqv+vbaq/ry0qP67s6f+urKm/riw
pP62rqL+sqqd/qacjv6elIT+oJeI/tvX0f79/f3+8vLz/tbW1+TT0tIjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADd3d064+Pk6Pn5+f79/f3/4d7a/6adj/+bkYH+n5aG/6qh
k/6zq57/ta2h/7auov63r6P/uLCk/7iwpP65saX/ubKm/7mypv66sqb/ubGl/7mxpf64sKT/t7Ck/7auov62rqL/ta2g/7Kqnv+poJL+n5WF/5qQgP+lnI7+4d7Z//39/f/y8vP+2dnZ6NPT0joAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA3NzcOuDg4eT4+Pj//v7+//Lx7/+0raH+mI5+/5mPf/6elYX/p5+R/6+nm/6yqp7/s6uf/7Orn/60rKD/tKyg/7WtoP61raD/tKyg/7Ss
n/6zq5//s6ue/7Kqnf6vp5r/p56Q/56Uhf+Yjn7+l419/7SsoP/y8e/+/f39//Hx8f/W1tfk09LSOgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANva
2iPd3d3M8fHy/vz8/P79/f3+3NnU/qmhlP6Vi3v+lox8/piOfv6elYb+pJyO/qmhk/6spJf+rqWZ/q6mmv6vppr+rqWZ/qyjl/6poJP+pJuN/p6Vhv6Xjn7+lox8/pWLe/6poJP+3NnT/v39/f76+vr+6enq/tTU
1MzS0tIjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADi4eER2NfYlufn6Pf4+Pj+/f39//n5+P7a19H/sKmd/5iPgP6TiXn/lIp6/5WL
e/6WjHz/l41+/5iOfv6Yjn//l419/5aMfP6Vi3v/lIp6/5OJef6Yj4D/sKid/9rX0f/5+fj+/Pz8//Pz8//f3+D3z8/QltrZ2REAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5eXlAt3d3ELa2trS7Ozs/vn5+f7+/v7//f39/+7t6v7V0cv/urSq/6WdkP6WjH7/koh5/5KIef6SiHn/koh5/5aMfv6lnZD/urSq/9XRy/7u7ev//f39//39
/f/19fX+5eXm/tLS09LW1tVC397eAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADj4+IE2dnZWNnZ
2sfn5+f98/Pz/vr6+v79/f3+/f39/v39/f79/f3++vr5/vPy8P7z8vD++vr5/v39/f79/f3+/f39/vz8/P74+Pj+7+/v/uHh4v3T09TH09PTWN7d3QQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOTk5Aje3d071tbWldvb3Nrk5OX27Ozs/vLy8v719fX/9/f3//f3+P739/j/9vb2//T0
9f7w8PD/6enq/uHh4fbX19ja0dHSldnZ2Tvg398IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOLh4gbf3t4v2NfYXtzb24TV1dar09PUu9XV1sjU1NXI0tLTu9TU1Kva2dmE1dXVXtzb2y/f398GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//4AB//8AAP/8AAA//wAA//AAAA//AAD/wAAAA/8AAP+AAAAB/wAA/wAAAAD/AAD+AAAAAH8AAPwAAAAAPwAA+AAAAAAfAADwAAAAAA8AAOAA
AAAABwAA4AAAAAAHAADAAAAAAAMAAMAAAAAAAwAAgAAAAAABAACAAAAAAAEAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAQAAgAAAAAABAACAAAAAAAEAAMAAAAAAAwAAwAAAAAADAADgAAAAAAcAAOAAAAAABwAA8AAAAAAPAAD4AAAAAB8AAPwA
AAAAPwAA/gAAAAB/AAD/AAAAAP8AAP+AAAAB/wAA/8AAAAP/AAD/8AAAD/8AAP/8AAA//wAA//+AAf//AAA=
"@
#endregion ******** $Pause48Icon ********
$PILLargeImageList.Images.Add("Pause48Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Pause48Icon))))

#region ******** $Stop48Icon ********
$Stop48Icon = @"
AAABAAEAMDAAAAEAIACoJQAAFgAAACgAAAAwAAAAYAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAOzs7Abq6uov5eXlXunp6YTk5OWr4uLjvOTk5cjk5OTJ4eHivOPj46vn5+eE4uLjXujo6C/q6uoGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7u7gjq6uo85OTklenp6try8vL29/f3/vr6+/78/Pz//Pz8//z8/P79/f3//Pz8//z8
/P75+fn/9fX2/u7u7/bm5uba4N/glefm5zzr6+sIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAADu7u4E5ubmWOfn6Mf09PT9+/v7/v39/f79/f3+/f39/v39/f79/f3+/Pz7/vj49v74+Pb+/Pz7/v39/f79/f3+/f39/v39/f79/f3++fn5/u/v7/3h4eLH4eHhWOrq6gQAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7+/vAunp6ULn5+fS9vb3//39/f7+/v7//f39//f29f7u7Oj/5eHc/9zY
0f7Y08v/2NPL/9jTy/7Y08v/2NPL/9jTy/7c2NH/5eHc/+7r6P739vX//f39//7+/v/8/Pz+8vLy/9/f4NLk4+RC6+rqAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADs7OwR5OTklvLy8/f9/f3+/v7+//z7+/7w7uv/4d3W/9nUzf7Y08v/2NPL/9nTy/7Z1Mv/2tTL/9rUy/7a1Mv/2dTL/9nUy/7Z08v/2NPL/9jTyv7Z1M3/4d3W//Du
6//8+/v+/f39//v7+//r6+z33Nzclufm5hEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOXl5SPp6erM+fn5/v39/f79/f3+8O7r/t3Y
0v7Y0sr+2NPL/tnUy/7a1Mz+29XN/tvWzf7b1s7+29bO/tvWzv7b1s7+29bO/tvWzv7b1s3+29XM/trUy/7Z08v+2NPK/tfSyv7d2NL+8O7r/v39/f79/f3+9PT0/uDg4cze3t4jAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5+fnOuzs7OT8/Pz//v7+//n49//g3Nb+19HK/9jTyv7a1Mv/29XN/9vWzv7c18//3NfP/9zXz/7d2ND/3djQ/93Y0P7d2ND/3djQ/93Y
0P7c18//3NfP/9vWzv7b1s3/29XM/9rUy//Y0sr+19HK/+Dc1f/5+Pb+/v7+//n5+f/i4uLk39/fOgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADn5+c67u7u6Pz8
/P79/f3/8e/t/9rVzv/X0cr+2tTL/9vVzf7b1s7/3NfP/93Y0P7e2dH/3tnS/97Z0v7f2tP/39rT/9/a0/7f2tP/39rT/9/a0/7e2dL/3tnR/93Y0f7c19D/3NfP/9vWzv/a1cz+2dPL/9fRyf/a1c7+8e/s//39
/f/5+fn+4+Pk6N/e3zoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOXl5SPs7Ozk/Pz8/v39/f7u7Oj+19LK/tjSyv7a1Mz+29bO/tzXz/7d2ND+3tnR/t/a0/7f2tT+4NvV/uDb
1f7g3NX+4NzV/uHc1v7h3Nb+4NzV/uDc1f7g29X+39vU/t/a0/7e2dL+3djR/tzX0P7b1s7+29XN/trUy/7Y0sn+1tHJ/u7s6P79/f3++fn5/uDg4eTd3N0jAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA7OzsEenp6sz8/Pz//f39/+7s6P7V0Mf/2NLJ/9vVzP/b1s7+3djQ/97Z0f7f2tP/39vU/+Db1f7h3db/4t3X/+Ld1/7j3tj/497Y/+Pe2P7j3tj/497Y/+Le1/7i3df/4d3W/+Dc1f7g29X/39rU/97Z
0v/d2ND+3NfP/9vWzv/a1Mv+19HJ/9XQx//u7Oj+/f39//j4+P/e3t7M5eTkEQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADv7+8C5OTklvn5+f7+/v7/8e/s/9bRyf7X0sn/29XM/9zXz//d2ND+3tnS/9/a
1P7g29X/4d3W/+Ld1/7j3tj/49/Y/+Pf2f7k4Nr/5ODa/+Tg2/7k4Nv/5ODa/+Tg2v7j39n/49/Y/+Le1/7i3df/4NzV/+Db1f/f2tP+3tnR/9zXz//b1s7+2tTM/9fRyf/W0cj+8e/s//7+/v/x8fL+2NjZlujo
6AIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADp6elC8vLz9/39/f74+Pb/2dTM/9fRyP7b1cz/3NfP/93Y0P/e2dL+39vU/+Dc1f7i3df/497Y/+Pf2f7k4Nr/5eDb/+Xh3P7m4dz/5uLd/+bi3f7m4t3/5uLd/+bh
3P7l4dz/5ODb/+Tg2v7j39j/4t7X/+Hd1v/g3NX+39rU/97Z0f/c18/+29bO/9rUy//W0cj+2NTL//j49v/8/Pz+5+fo9+Dg4EIAAAAAAAAAAAAAAAAAAAAAAAAAAO7u7gTn5+fS/f39/v39/f7e2tP+1dDH/trU
y/7b1s7+3djQ/t7Z0v7f29T+4NzV/uLd1/7j39j+5ODa/uXg2/7m4dz+5+Ld/ufj3v7n5N/+6OTf/ujk3/7o5N/+6OTf/ufj3v7n497+5uLd/ubh3P7k4Nv+5ODZ/uPe2P7h3db+4NzV/t/a1P7e2dH+3NfP/tvW
zv7Z08v+1c/H/t7Z0v79/f3++Pj4/tra2tLm5uUEAAAAAAAAAAAAAAAAAAAAAObm5lj29vf//v7+/+/t6f7UzsT/2dPK/9vWzv7d2ND/3tnS/9/b1P/g3NX+4t3X/+Pf2f7k4Nv/5uHc/+fi3f7n5N//6OXg/+jl
4P7p5uH/6ubi/+rm4v7q5uL/6ebh/+nl4f7o5eD/6OTf/+fj3v7m4t3/5eHc/+Tg2v/j39j+4t3X/+Dc1f/f2tP+3tnR/9zXz//b1c3+2NLK/9TOxP/v7en+/f39/+zs7P/b29tYAAAAAAAAAAAAAAAA7u7uCOfn
6Mf9/f3++/v7/tvVzf7W0Mf+29XN/tzXz/7e2dH+39rU/uDc1f7i3df+49/Z/uTg2/7m4dz+5+Pe/ujk3/7p5eH+6ubi/urn4/7r6OT+7Onl/uzp5f7s6eX+7Onl/uvo5P7q5+P+6ebh/ujl4P7n5N/+5+Ld/uXh
3P7k4Nr+497Y/uHd1v7g29X+39rT/t3Y0P7b1s7+2tTM/tbQx/7b1c3++/v7/vn5+f7a2trH5ubmCAAAAAAAAAAA6urqPPT09P3+/v7/7+zo/9TOxP7a1Mv/29bO/93Y0P7f2tP/4NvV/+Ld1//j39j+5ODb/+bh
3P7n497/6OXg/+nm4f7r5+P/7Onl/+3q5v7u6+f/7uvn/+7r5/7u6+f/7uvn/+3r5/7t6ub/7Ojk/+rn4/7p5uH/6OTf/+fj3v/l4dz+5ODa/+Pe2P/h3Nb+39vU/97Z0v/c18/+29bN/9nTyv/UzsT+7uzo//7+
/v/n5+f94ODgOwAAAAAAAAAA5OTklfv7+/79/f3/3tnR/9bQxv7b1c3/3NfP/97Z0f7f29T/4d3W/+Pe2P/k4Nr+5uHc/+fj3v7o5eD/3dnU/9jUz/7Z1tH/2tbS/9rX0v7b2NP/29nU/9vZ1P7b2dT/29nT/9vY
0/7a19L/2tbS/9nV0P7X087/3NjT/+jk3//n4t3+5eHc/+Pf2f/i3df+4NzV/9/a0//d2ND+3NbO/9rVzP/Vz8b+3djR//39/f/z8/P+19bXlQAAAADs7OwG6enq2/39/f739fT+1c7F/tnTyv7e2tL+3djQ/t/a
0/7g29X+4t3X/uPf2f7l4Nv+5+Ld/ujk3/7p5uH+lpCI/mxmXP5sZlz+bGZc/mxmXP5sZlz+bGZc/mxmXP5sZlz+bGZc/mxmXP5sZlz+bGZc/mxmXP5sZlz+lpCI/unl4f7n497+5uHc/uTg2/7j3tj+4d3W/t/b
1P7e2dH+3NfP/t7Z0f7Y0sn+1M7E/vb19P76+vr+29vc2+Pj4wbq6uov8vLy9v39/f7s6eX/08zD/93Xz/7i3tj/39rS/9/a1P7h3db/497Y/+Tg2v/m4dz+5+Tf/+nl4f7r5+P/lpGJ/2xmXP5tZ13/bWdd/2xm
XP5tZ13/bWdd/2xmXP5tZ13/bWdd/2xmXP5tZ13/bWdd/2xmXP5tZ13/lpGJ/+rm4v/o5eD+5+Pe/+Xh3P/j39n+4t3X/+Dc1f/f2tP+3tnR/+Le1//c1s7+08zD/+zp5f/9/f3+5OTl9uDg4C/l5eVe9/f3/v39
/f7i3df/1c7E/9/a0/7j39n/4t7X/+Db1f7i3df/49/Y/+Xg2//n4t3+6OXg/+rm4v7s6eX/l5KK/2xmXP5tZ13/bWdd/21nXf5uZ13/bmdd/21nXf5tZ13/bWdd/2xmXP5tZ13/bWdd/2xmXP5tZ13/lpKJ/+vo
5P/p5uH+5+Tf/+bh3P/k4Nr+497Y/+Hc1v/f2tT+4d3W/+Pf2f/f2dL+1c7E/+Ld1v/9/f3+7Ozs/tjY2V7p6emF+vr7/v39/f7Y0sr+1tDH/uHc1f7j39n+5eHb/uHd1/7i3df+49/Z/uXh3P7n497+6OXg/urn
4/7n5N/+k42E/m1nXf5vaF7+cGlf/nFqYP5xamD+cWpg/nFqYP5waV/+bmhe/m1nXf5sZlz+bGZc/mxmXP5sZlz+ko2E/ubj3v7q5uL+6OTf/ufi3f7k4Nv+49/Y/uHd1v7h3Nb+5ODa/uPf2f7g3NT+1s/G/tjS
yf79/f3+8vLy/t3d3YXk5OWq/Pz8//39/f7TzMP/2NLJ/+Hd1/7k4Nr/5uLc/+bi3P7j3tj/5ODa/+bh3P/n5N/+5eHc/9vXz/7V0Mf/kIqA/3BqYP5ybGH/c2xi/3RtYv50bWL/dG1i/3RtYv5zbGL/cmth/3Bp
X/5uaF7/bWdd/2xmXP5tZ13/j4l//9XQx//b1s/+5eHb/+fj3v/l4dz+49/Z/+Le1//l4dz+5eHb/+Pf2f/h3db+19HJ/9PMwv/9/f3+9fX1/9bW16ri4uO7/Pz8//z7+/7RysD/2dPK/+Le2P7k4Nr/5uLc/+fj
3v7n493/5eHb/+bh3P/e2tP+1tHI/9TPxv7V0Mf/koyC/3NsYv51bmP/dm9k/3dwZf54cWX/eHFl/3dwZf52b2T/dW5j/3NsYv5xamD/bmhe/21nXf5tZ13/j4l//9XQx//Uz8b+1dDH/97Z0v/l4Nv+5ODa/+bi
3f/m497+5eHb/+Pf2f/i3df+2NLK/8/Ivv/8+/v+9/f3/9PT1Lvk5OXI/Pz8/vf39f7Nxrz+2dPK/uLe2P7k4Nr+5uLc/ufj3v7p5eD+5+Pf/trVzf7TzcT+1M7F/tTPxv7Uz8b+lI2D/nZvZP54cWX+eXJm/ntz
Z/57dGj+e3Ro/ntzZ/55cmb+eHFl/nZvY/5zbWL+cWpg/m5oXv5sZlz+j4l//tTPxv7Uz8X+1M7E/tPNxP7a1c3+5uPd/ujk3/7m497+5eHb/uPf2f7i3tf+2NLJ/szEuv739vX+9/f4/tXV1sjk5OTI/f39//f2
9f7Lw7n/2NLK/+Le2P7k4Nr/5uLc/+fj3v7o5N7/4NvU/9rVzf/Vz8X+1M7E/9POxP7Uz8X/lY+E/3hxZv57c2j/fXVp/353av5/d2r/f3dq/352av59dWn/e3Nn/3hxZf52b2T/c2xi/3BpX/5tZ13/joh//9TO
xf/TzcT+1M7E/9XPxf/a1c3+4NvU/+fj3v/m497+5uLb/+Pf2f/i3tf+2NHJ/8nBtv/39vX+9/f4/9TU1cjh4eK7/Pz8//z7+v7Hv7T/19HI/+Le1/7k4Nr/5uLc/+bj3v7f2tP/3NfQ/9zY0P/b1s7+1tHH/9TO
xP7UzsT/lo+F/3tzZ/5+dmn/gHhr/4J5bP6Dem3/g3pt/4J5bP6AeGv/fXZp/3tzZ/54cWX/dW5j/3JrYf5uaF7/joh//9TOxP/UzsT+1tHH/9vWzv/c19D+3NfP/9/a0//m4t3+5eHb/+Pf2f/i3df+1s/H/8W9
sv/7+/r+9vb2/9LS07vj4+Oq/Pz8/v39/f7FvbH+1M3D/uHd1v7k4Nr+5eHb/uDb1f7b1s7+29bP/tzXz/7c19D+3NfQ/tnUzP7Vz8X+l5CF/n11af6AeGv+g3pt/oV8b/6GfXD+hn1w/oR8b/6Cem3+f3hr/n11
af55cmb+dm9k/nNsYv5waV/+joh//tXPxf7Z1Mz+3NfQ/tzX0P7b1s/+29bO/tvWzv7g29T+5eHb/uPf2f7h3db+0szC/sO7r/79/f3+9PT1/tTU1arn5+eF+fn5//39/f7Kwrj/zse9/+Dc1f7j39n/4dzW/9zW
z/7c18//3NfP/9zXz//c18/+3NfQ/9zX0P7c19D/m5SJ/353av6CeW3/hXxv/4d/cf6KgXL/ioFy/4d/cf6EfG//gnls/352av57c2f/d3Bl/3RtYv5xamD/kYuC/9zX0P/c19D+3NfP/9zXz//b1s7+3NfP/9zX
z//b1s7+4NvV/+Pf2f/g29T+zca8/8jBt//9/f3+8PDw/9vb24Xi4uNe9fX2/v39/f7Uz8f/x7+0/9/a0v7h3df/2NLL/9nUzP7b1s7/3NfP/9zXz//b1s7+3NfP/9zXz/7c19D/nJWK/393a/6Dem3/hn5w/4qB
cv6NhHX/jYN1/4qBcv6GfXD/g3pt/393av57dGj/eHFl/3RtYv5xamD/koyD/9zX0P/c18/+3NfP/9zXz//b1s7+3NfP/9vWzv/Y0sv+19HK/+Hd1v/e2dH+xr6z/9PNxv/9/f3+6enq/tbW1l7o6Ogv7u7v9v39
/f7j39r+v7aq/tvVzf7Y08v+083E/tbQyf7Y08v+29bO/tvWzv7b1s7+29bO/tvWzv7b1s7+m5SK/n93a/6Dem3+hn5w/oqBc/6OhHX+jYN1/oqBcv6GfXD+g3pt/n93av57dGj+eHFl/nRtYv5xamD+koyC/tvW
zv7b1s7+29bO/tvWzv7b1s7+2tXN/tfSyv7Vz8f+0szC/tjSyf7a1Mz+vbWo/uPf2f78/Pz+4eHh9t3d3S/q6uoG5ubm2/39/f7z8u//vLOn/87GvP7Mxbv/z8m//9LMw/7Vz8f/19LK/9rUzf/b1s7+3NfP/9vW
zv7c18//m5SK/353av6Cem3/hXxv/4h/cf6KgXP/ioFz/4d/cf6FfG//gnls/353av57c2f/d3Bl/3RtYv5xamD/koyC/9zXz//b1s7+3NfP/9vWzv/Z08z+1tHJ/9TOxv/Ry8H+zse9/8vDuf/Nxbv+urKl//Px
7//4+Pj+19fY2+Dg4AYAAAAA39/glfn5+f79/f3/yMG3/7mwo/7Du6//zMS7/87Ivf7Ry8H/1M7F/9bQyf/Y0sv+2tXN/9vWzv7c18//0szE/83Iv/7OyMD/zsnA/8/JwP7PycD/z8nA/8/JwP7OycD/zsjA/83I
v/7Nx7//zMe+/8zGvv7Lxr7/0MvD/9zXz//b1s7+2dTM/9fSyv/Vz8f+083D/9DKwP/Oxrz+y8O6/8K6rv+4r6L+x8C1//39/f/v7+/+0tLSlQAAAAAAAAAA5ubmO+/v7/3+/v7/5ODb/7OqnP65sKT/x7+1/8vD
uf7Oxr3/0Mm//9LMwv/Uzsb+1tDJ/9jSyv7Z1M3/29bO/9vWzv7c18//3NfP/9vWzv7c18//3NfP/9vWzv7c18//3NfP/9vWzv7c18//3NfP/9vWzv7b1s7/2tXO/9nUzP/X0cr+1tDI/9TOxf/Ry8H+z8i+/83F
u//Kwrj+x760/7ivov+yqJr+4+Db//39/f/h4eL929vbOwAAAAAAAAAA6urqCOHh4sf8/Pz++vr5/r61qv6yqZr+vbWp/se/tf7Jwrj+zMS6/s7Hvf7Qyb/+0szC/tTOxf7V0Mj+1tHJ/tjSy/7Z08z+2tTN/trV
zf7b1s7+29bO/tvWzv7b1s7+29bO/tvVzv7a1c3+2dTM/tjTy/7X0sr+1tDJ/tXPx/7TzcT+0cvB/s/Ivv7Nxrz+y8O5/snAt/7GvrT+vLSo/rGnmf68tKj++vr5/vX19f7T09TH4eHhCAAAAAAAAAAAAAAAAODg
4Fjy8vL//f39/+Pg2/6to5T/sqmb/8C4rf7FvrP/yMC2/8rCuP/MxLv+zsa9/8/Ivv7Ry8D/0szD/9TOxf7Vz8f/1tDI/9bQyf7X0cn/19HK/9fSyv7X0sr/19HK/9fRyf7W0Mn/1c/I/9TPxv7TzcT/0szC/9DK
wP/PyL7+zca8/8vEuv/Jwbf+x7+1/8S9sv+/t6z+saeZ/6uik//i39r+/Pz8/+Xl5v/U1NRYAAAAAAAAAAAAAAAAAAAAAOno6ATf3+DS+/v7//39/f7Aua7/qqGS/7OqnP7Bua7/w7yx/8a+tP/IwLb+ysK4/8vD
uv7Nxrz/zse9/8/Ivv7QysD/0cvB/9LLwv7SzMP/0szD/9PNw/7TzcP/0szD/9LMwv7Ry8H/0crA/9DJv/7Px77/zsa8/8zFu//Lw7n+ycG3/8e/tf/FvbP+w7uw/8C4rf+xqJr+qZ+Q/7+3rP/9/f3+8/Pz/9LS
09Lf3t4EAAAAAAAAAAAAAAAAAAAAAAAAAADi4eFC6+vs9/39/f7z8vD+sKeZ/qeej/6zqpz+vrar/sG5rv7Du7H+xb2z/se/tP7IwLb+ysK4/svDuf7LxLr+zcW7/s3GvP7Nxrz+zse9/s7Hvf7Ox73+zce9/s3G
vP7Nxbv+zMS7/svDuv7Kwrn+ycG3/sjAtf7GvrT+xL2y/sO7sP7AuK3+vbWp/rGom/6mnI3+rqaY/vPy8P76+vr+39/g99fX10IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADp6egC29vclvT09P7+/v7/5OHc/6ed
j/6lm4z/saia/7u0p/++tqv+wLmt/8K7sP7EvLL/xb2z/8a+tP7Hv7X/yMC2/8nBt/7Jwbf/ysK4/8rCuP7Kwrj/ycK4/8nBt/7Jwbf/yMC2/8e/tf7GvrT/xL2y/8O8sf/Buq/+wLit/722qv+6sqb+r6eZ/6Sa
i/+mnI7+5ODc//39/f/p6er+z8/QluDf3wIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5eTkEeDg4Mz5+fn//f39/9zY0/6hl4f/o5mK/6yjlv+4sKT+u7On/721qf6/t6v/wLit/8G6rv7Cu7D/w7yx/8S8
sv7EvbL/xb2y/8W+s/7FvrP/xL2y/8S9sv7EvLL/w7uw/8K6r/7Bua7/wLis/762qv+8tKj+urKm/7evo/+so5X+opiJ/5+Vhv/c2NL+/f39//Hx8f/U1NTM29raEQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAN3c3CPi4uLk+fn5/v39/f7b19L+oZiJ/p+Vhf6mnY/+s6qe/revo/64saX+urKm/ru0qP69tan+vraq/r+3q/6/t6z+v7is/r+4rP6/uKz+v7es/r+3q/6+tqv+vbaq/ry0qP67s6f+urKm/riw
pP62rqL+sqqd/qacjv6elIT+oJeI/tvX0f79/f3+8vLz/tbW1+TT0tIjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADd3d064+Pk6Pn5+f79/f3/4d7a/6adj/+bkYH+n5aG/6qh
k/6zq57/ta2h/7auov63r6P/uLCk/7iwpP65saX/ubKm/7mypv66sqb/ubGl/7mxpf64sKT/t7Ck/7auov62rqL/ta2g/7Kqnv+poJL+n5WF/5qQgP+lnI7+4d7Z//39/f/y8vP+2dnZ6NPT0joAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA3NzcOuDg4eT4+Pj//v7+//Lx7/+0raH+mI5+/5mPf/6elYX/p5+R/6+nm/6yqp7/s6uf/7Orn/60rKD/tKyg/7WtoP61raD/tKyg/7Ss
n/6zq5//s6ue/7Kqnf6vp5r/p56Q/56Uhf+Yjn7+l419/7SsoP/y8e/+/f39//Hx8f/W1tfk09LSOgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANva
2iPd3d3M8fHy/vz8/P79/f3+3NnU/qmhlP6Vi3v+lox8/piOfv6elYb+pJyO/qmhk/6spJf+rqWZ/q6mmv6vppr+rqWZ/qyjl/6poJP+pJuN/p6Vhv6Xjn7+lox8/pWLe/6poJP+3NnT/v39/f76+vr+6enq/tTU
1MzS0tIjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADi4eER2NfYlufn6Pf4+Pj+/f39//n5+P7a19H/sKmd/5iPgP6TiXn/lIp6/5WL
e/6WjHz/l41+/5iOfv6Yjn//l419/5aMfP6Vi3v/lIp6/5OJef6Yj4D/sKid/9rX0f/5+fj+/Pz8//Pz8//f3+D3z8/QltrZ2REAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5eXlAt3d3ELa2trS7Ozs/vn5+f7+/v7//f39/+7t6v7V0cv/urSq/6WdkP6WjH7/koh5/5KIef6SiHn/koh5/5aMfv6lnZD/urSq/9XRy/7u7ev//f39//39
/f/19fX+5eXm/tLS09LW1tVC397eAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADj4+IE2dnZWNnZ
2sfn5+f98/Pz/vr6+v79/f3+/f39/v39/f79/f3++vr5/vPy8P7z8vD++vr5/v39/f79/f3+/f39/vz8/P74+Pj+7+/v/uHh4v3T09TH09PTWN7d3QQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOTk5Aje3d071tbWldvb3Nrk5OX27Ozs/vLy8v719fX/9/f3//f3+P739/j/9vb2//T0
9f7w8PD/6enq/uHh4fbX19ja0dHSldnZ2Tvg398IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOLh4gbf3t4v2NfYXtzb24TV1dar09PUu9XV1sjU1NXI0tLTu9TU1Kva2dmE1dXVXtzb2y/f398GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//4AB//8AAP/8AAA//wAA//AAAA//AAD/wAAAA/8AAP+AAAAB/wAA/wAAAAD/AAD+AAAAAH8AAPwAAAAAPwAA+AAAAAAfAADwAAAAAA8AAOAA
AAAABwAA4AAAAAAHAADAAAAAAAMAAMAAAAAAAwAAgAAAAAABAACAAAAAAAEAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAQAAgAAAAAABAACAAAAAAAEAAMAAAAAAAwAAwAAAAAADAADgAAAAAAcAAOAAAAAABwAA8AAAAAAPAAD4AAAAAB8AAPwA
AAAAPwAA/gAAAAB/AAD/AAAAAP8AAP+AAAAB/wAA/8AAAAP/AAD/8AAAD/8AAP/8AAA//wAA//+AAf//AAA=
"@
#endregion ******** $Stop48Icon ********
$PILLargeImageList.Images.Add("Stop48Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Stop48Icon))))

#endregion ******** PIL Large ImageList Icons ********

# ************************************************
# PIL Form
# ************************************************
#region $PILForm = [System.Windows.Forms.Form]::New()
$PILForm = [System.Windows.Forms.Form]::New()
$PILForm.BackColor = [MyConfig]::Colors.Back
$PILForm.ControlBox = $True
$PILForm.Enabled = $True
$PILForm.Font = [MyConfig]::Font.Regular
$PILForm.ForeColor = [MyConfig]::Colors.Fore
$PILForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$PILForm.Icon = [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($PILFormIcon)))
$PILForm.KeyPreview = $True
$PILForm.MinimizeBox = $True
$PILForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * [MyConfig]::FormMinWidth), ([MyConfig]::Font.Height * [MyConfig]::FormMinHeight))
$PILForm.Name = "PILForm"
$PILForm.ShowIcon = $True
$PILForm.ShowInTaskbar = $True
$PILForm.TabIndex = 0
$PILForm.TabStop = $True
$PILForm.Tag = (-not [MyConfig]::Production)
$PILForm.Text = "$([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
#endregion $PILForm = [System.Windows.Forms.Form]::New()

#region ******** Function Start-PILFormClosing ********
function Start-PILFormClosing
{
  <#
    .SYNOPSIS
      Closing Event for the PIL Form Control
    .DESCRIPTION
      Closing Event for the PIL Form Control
    .PARAMETER Sender
       The  Control that fired the Closing Event
    .PARAMETER EventArg
       The Event Arguments for the  Closing Event
    .EXAMPLE
       Start-PILFormClosing -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Form]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Closing Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0

  if ([MyConfig]::Production)
  {
    [Void][Console.Window]::Show()
    [System.Console]::Title = "CLOSING: $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
    $PILForm.Tag = $True
  }

  Write-Verbose -Message "Exit Closing Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILFormClosing ********
$PILForm.add_Closing({Start-PILFormClosing -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILFormKeyDown ********
function Start-PILFormKeyDown
{
  <#
    .SYNOPSIS
      KeyDown Event for the PIL Form Control
    .DESCRIPTION
      KeyDown Event for the PIL Form Control
    .PARAMETER Sender
       The  Control that fired the KeyDown Event
    .PARAMETER EventArg
       The Event Arguments for the  KeyDown Event
    .EXAMPLE
       Start-PILFormKeyDown -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Form]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter KeyDown Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0
  
  If ($EventArg.Control -and $EventArg.Alt)
  {
    Switch ($EventArg.KeyCode)
    {
      "F10"
      {
        If ($PILForm.Tag)
        {
          # Hide Console Window
          #$Script:VerbosePreference = "SilentlyContinue"
          #$Script:DebugPreference = "SilentlyContinue"
          [System.Console]::Title = "RUNNING: $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
          [Void][Console.Window]::Hide()
          $PILForm.Tag = $False
        }
        Else
        {
          # Show Console Window
          #$Script:VerbosePreference = "Continue"
          #$Script:DebugPreference = "Continue"
          [Void][Console.Window]::Show()
          [System.Console]::Title = "DEBUG: $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
          $PILForm.Tag = $True
        }
        $PILForm.Activate()
        $PILForm.Select()
        Break
      }
      "Back"
      {
        $Response = Get-UserResponse -Title "Abort Processing" -Message "Would you like to Abort Processing the Current Item List?" -ButtonLeft Yes -ButtonRight No -ButtonDefault No -Icon ([System.Drawing.SystemIcons]::Question)
        If (-not $Response.Success)
        {
          Try
          {
            If (-not $PILTopMenuStrip.Items["ProcessItems"].Enabled)
            {
              $TmpRSPool = Get-MyRSPool
              Stop-MyRSJob
              $TmpRSPool.SyncedHash["Terminate"] = $True
              $TmpRSPool.SyncedHash["Pause"] = $False
              Close-MyRSPool
            }
          }
          Catch
          {
            
          }
        }
      }
    }
  }
  Else
  {
    Switch ($EventArg.KeyCode)
    {
      "F2"
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Show Change Log for $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
        $ScriptContents = ($Script:MyInvocation.MyCommand.ScriptBlock).ToString()
        $CLogStart = ($ScriptContents.IndexOf("<#") + 2)
        $CLogEnd = $ScriptContents.IndexOf("#>")
        Show-ChangeLog -ChangeText ($ScriptContents.SubString($CLogStart, ($CLogEnd - $CLogStart)))
        Break
      }
    }
  }
  
  Write-Verbose -Message "Exit KeyDown Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILFormKeyDown ********
$PILForm.add_KeyDown({Start-PILFormKeyDown -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILFormLoad ********
function Start-PILFormLoad
{
  <#
    .SYNOPSIS
      Load Event for the PIL Form Control
    .DESCRIPTION
      Load Event for the PIL Form Control
    .PARAMETER Sender
       The  Control that fired the Load Event
    .PARAMETER EventArg
       The Event Arguments for the  Load Event
    .EXAMPLE
       Start-PILFormLoad -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Form]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Load Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0

  $Screen = ([System.Windows.Forms.Screen]::FromControl($Sender)).WorkingArea
  $Sender.Left = [Math]::Floor(($Screen.Width - $Sender.Width) / 2)
  $Sender.Top = [Math]::Floor(($Screen.Height - $Sender.Height) / 2)

  if ([MyConfig]::Production)
  {
    # Disable Control Close Menu / [X]
    #[ControlBox.Menu]::DisableFormClose($PILForm.Handle)

    [System.Console]::Title = "RUNNING: $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
    [Void][Console.Window]::Hide()
    $PILForm.Tag = $False
  }
  else
  {
    [Void][Console.Window]::Show()
    [System.Console]::Title = "DEBUG: $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
    $PILForm.Tag = $True
  }

  Write-Verbose -Message "Exit Load Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILFormLoad ********
$PILForm.add_Load({Start-PILFormLoad -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILFormMove ********
function Start-PILFormMove
{
  <#
    .SYNOPSIS
      Move Event for the PIL Form Control
    .DESCRIPTION
      Move Event for the PIL Form Control
    .PARAMETER Sender
       The  Control that fired the Move Event
    .PARAMETER EventArg
       The Event Arguments for the  Move Event
    .EXAMPLE
       Start-PILFormMove -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Form]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Move Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0


  Write-Verbose -Message "Exit Move Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILFormMove ********
$PILForm.add_Move({Start-PILFormMove -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILFormResize ********
function Start-PILFormResize
{
  <#
    .SYNOPSIS
      Resize Event for the PIL Form Control
    .DESCRIPTION
      Resize Event for the PIL Form Control
    .PARAMETER Sender
       The  Control that fired the Resize Event
    .PARAMETER EventArg
       The Event Arguments for the  Resize Event
    .EXAMPLE
       Start-PILFormResize -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Form]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Resize Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0


  Write-Verbose -Message "Exit Resize Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILFormResize ********
$PILForm.add_Resize({Start-PILFormResize -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILFormShown ********
function Start-PILFormShown
{
  <#
    .SYNOPSIS
      Shown Event for the PIL Form Control
    .DESCRIPTION
      Shown Event for the PIL Form Control
    .PARAMETER Sender
       The  Control that fired the Shown Event
    .PARAMETER EventArg
       The Event Arguments for the  Shown Event
    .EXAMPLE
       Start-PILFormShown -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Form]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Shown Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0

  $Sender.Refresh()

  #Write-KPIEvent -Source "Utility" -EntryType "Information" -EventID 1 -Category 0 -Message "Begin Running $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"

  $HashTable = @{"ShowHeader" = $True; "ConfigFile" = $ConfigFile; "ImportFile" = $ImportFile}
  $ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Display-InitiliazePILUtility -RichTextBox $RichTextBox -HashTable $HashTable }
  $DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title "Initializing $([MyConfig]::ScriptName)" -ButtonMid "OK" -HashTable $HashTable

  if ([MyConfig]::Production)
  {
    # Enable $PILTimer
    $PILTimer.Enabled = ([MyConfig]::AutoExitMax -gt 0)
  }

  Write-Verbose -Message "Exit Shown Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILFormShown ********
$PILForm.add_Shown({Start-PILFormShown -Sender $This -EventArg $PSItem})

#region $PILTimer = [System.Windows.Forms.Timer]::New()
$PILTimer = [System.Windows.Forms.Timer]::New($PILFormComponents)
$PILTimer.Enabled = $False
$PILTimer.Interval = [MyConfig]::AutoExitTic
#$PILTimer.Tag = [System.Object]::New()
#endregion $PILTimer = [System.Windows.Forms.Timer]::New()

#region ******** Function Start-PILTimerTick ********
function Start-PILTimerTick
{
  <#
    .SYNOPSIS
      Tick Event for the PIL Timer Control
    .DESCRIPTION
      Tick Event for the PIL Timer Control
    .PARAMETER Sender
       The  Control that fired the Tick Event
    .PARAMETER EventArg
       The Event Arguments for the  Tick Event
    .EXAMPLE
       Start-PILTimerTick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Timer]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Tick Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit += 1
  Write-Verbose -Message "Auto Exit in $([MyConfig]::AutoExitMax - [MyConfig]::AutoExit) Minutes"
  if ([MyConfig]::AutoExit -ge [MyConfig]::AutoExitMax)
  {
    $PILForm.Close()
  }
  ElseIf (([MyConfig]::AutoExitMax - [MyConfig]::AutoExit) -le 5)
  {
    $PILBtmStatusStrip.Items["Status"].Text = "Auto Exit in $([MyConfig]::AutoExitMax - [MyConfig]::AutoExit) Minutes"
  }

  #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

  Write-Verbose -Message "Exit Tick Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILTimerTick ********
$PILTimer.add_Tick({Start-PILTimerTick -Sender $This -EventArg $PSItem})

# ************************************************
# PILMain Panel
# ************************************************
#region $PILMainPanel = [System.Windows.Forms.Panel]::New()
$PILMainPanel = [System.Windows.Forms.Panel]::New()
$PILForm.Controls.Add($PILMainPanel)
$PILMainPanel.BackColor = [MyConfig]::Colors.Back
$PILMainPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$PILMainPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$PILMainPanel.Enabled = $True
$PILMainPanel.Font = [MyConfig]::Font.Regular
$PILMainPanel.ForeColor = [MyConfig]::Colors.Fore
$PILMainPanel.Name = "PILMainPanel"
#$PILMainPanel.TabIndex = 0
#$PILMainPanel.TabStop = $False
#$PILMainPanel.Tag = [System.Object]::New()
#endregion $PILMainPanel = [System.Windows.Forms.Panel]::New()

#region ******** $PILMainPanel Controls ********

#region $PILItemListListView = [System.Windows.Forms.ListView]::New()
$PILItemListListView = [System.Windows.Forms.ListView]::New()
$PILMainPanel.Controls.Add($PILItemListListView)
$PILItemListListView.AllowColumnReorder = $True
$PILItemListListView.BackColor = [MyConfig]::Colors.TextBack
$PILItemListListView.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$PILItemListListView.CheckBoxes = $True
$PILItemListListView.Dock = [System.Windows.Forms.DockStyle]::Fill
#$PILItemListListView.FocusedItem = [System.Windows.Forms.ListViewItem]::New()
$PILItemListListView.Font = [MyConfig]::Font.Bold
$PILItemListListView.ForeColor = [MyConfig]::Colors.TextFore
$PILItemListListView.FullRowSelect = $True
$PILItemListListView.GridLines = $True
$PILItemListListView.LargeImageList = $PILSmallImageList
$PILItemListListView.ListViewItemSorter = [MyCustom.ListViewSort]::New()
$PILItemListListView.MultiSelect = $False
$PILItemListListView.Name = "PILItemListListView"
$PILItemListListView.OwnerDraw = $True
$PILItemListListView.ShowGroups = $False
$PILItemListListView.SmallImageList = $PILSmallImageList
#$PILItemListListView.TabStop = $True
#$PILItemListListView.Tag = [System.Object]::New()
#$PILItemListListView.Text = "PILItemListListView"
#$PILItemListListView.TopItem = [System.Windows.Forms.ListViewItem]::New()
$PILItemListListView.View = [System.Windows.Forms.View]::Details
#endregion $PILItemListListView = [System.Windows.Forms.ListView]::New()

#region ******** Function Start-PILItemListListViewColumnClick ********
function Start-PILItemListListViewColumnClick
{
  <#
    .SYNOPSIS
      ColumnClick Event for the PILItemList ListView Control
    .DESCRIPTION
      ColumnClick Event for the PILItemList ListView Control
    .PARAMETER Sender
       The ItemList Control that fired the ColumnClick Event
    .PARAMETER EventArg
       The Event Arguments for the ItemList ColumnClick Event
    .EXAMPLE
       Start-PILItemListListViewColumnClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ListView]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter ColumnClick Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0
  
  $Sender.ListViewItemSorter.Column = $EventArg.Column
  $Sender.ListViewItemSorter.Ascending = (-not $Sender.ListViewItemSorter.Ascending)
  $Sender.Sort()

  Write-Verbose -Message "Exit ColumnClick Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItemListListViewColumnClick ********
$PILItemListListView.add_ColumnClick({Start-PILItemListListViewColumnClick -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILItemListListViewDrawColumnHeader ********
function Start-PILItemListListViewDrawColumnHeader
{
  <#
    .SYNOPSIS
      DrawColumnHeader Event for the PILItemList ListView Control
    .DESCRIPTION
      DrawColumnHeader Event for the PILItemList ListView Control
    .PARAMETER Sender
       The ItemList Control that fired the DrawColumnHeader Event
    .PARAMETER EventArg
       The Event Arguments for the ItemList DrawColumnHeader Event
    .EXAMPLE
       Start-PILItemListListViewDrawColumnHeader -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ListView]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter DrawColumnHeader Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0

  $EventArg.Graphics.FillRectangle(([System.Drawing.SolidBrush]::New([MyConfig]::Colors.TitleBack)), $EventArg.Bounds)
  $EventArg.Graphics.DrawRectangle(([System.Drawing.Pen]::New([MyConfig]::Colors.TitleFore)), $EventArg.Bounds.X, $EventArg.Bounds.Y, $EventArg.Bounds.Width, ($EventArg.Bounds.Height - 1))
  $EventArg.Graphics.DrawString($EventArg.Header.Text, $Sender.Font, ([System.Drawing.SolidBrush]::New([MyConfig]::Colors.TitleFore)), ($EventArg.Bounds.X + [MyConfig]::FormSpacer), ($EventArg.Bounds.Y + (($EventArg.Bounds.Height - [MyConfig]::Font.Height) / 1)))

  Write-Verbose -Message "Exit DrawColumnHeader Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItemListListViewDrawColumnHeader ********
$PILItemListListView.add_DrawColumnHeader({Start-PILItemListListViewDrawColumnHeader -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILItemListListViewDrawItem ********
function Start-PILItemListListViewDrawItem
{
  <#
    .SYNOPSIS
      DrawItem Event for the PILItemList ListView Control
    .DESCRIPTION
      DrawItem Event for the PILItemList ListView Control
    .PARAMETER Sender
       The ItemList Control that fired the DrawItem Event
    .PARAMETER EventArg
       The Event Arguments for the ItemList DrawItem Event
    .EXAMPLE
       Start-PILItemListListViewDrawItem -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ListView]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter DrawItem Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0

  # Return to Default Draw
  $EventArg.DrawDefault = $True

  Write-Verbose -Message "Exit DrawItem Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItemListListViewDrawItem ********
$PILItemListListView.add_DrawItem({Start-PILItemListListViewDrawItem -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILItemListListViewDrawSubItem ********
function Start-PILItemListListViewDrawSubItem
{
  <#
    .SYNOPSIS
      DrawSubItem Event for the PILItemList ListView Control
    .DESCRIPTION
      DrawSubItem Event for the PILItemList ListView Control
    .PARAMETER Sender
       The ItemList Control that fired the DrawSubItem Event
    .PARAMETER EventArg
       The Event Arguments for the ItemList DrawSubItem Event
    .EXAMPLE
       Start-PILItemListListViewDrawSubItem -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ListView]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter DrawSubItem Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0

  # Return to Default Draw
  $EventArg.DrawDefault = $True

  Write-Verbose -Message "Exit DrawSubItem Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItemListListViewDrawSubItem ********
$PILItemListListView.add_DrawSubItem({Start-PILItemListListViewDrawSubItem -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILItemListListViewMouseDown ********
function Start-PILItemListListViewMouseDown
{
  <#
    .SYNOPSIS
      MouseDown Event for the PILItemList ListView Control
    .DESCRIPTION
      MouseDown Event for the PILItemList ListView Control
    .PARAMETER Sender
       The ItemList Control that fired the MouseDown Event
    .PARAMETER EventArg
       The Event Arguments for the ItemList MouseDown Event
    .EXAMPLE
       Start-PILItemListListViewMouseDown -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ListView]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter MouseDown Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0
  
  If ($EventArg.Button -eq [System.Windows.Forms.MouseButtons]::Right)
  {
    If (-not [String]::IsNullOrEmpty(($TmpItem = $Sender.GetItemAt($EventArg.Location.X, $EventArg.Location.Y))))
    {
      # Show Item Selected Context Menu
      $Sender.SelectedItems.Clear()
      $Sender.SelectedIndices.Add($TmpItem.Index)
      
      If ($TmpItem.Checked)
      {
        $TmpMenuText = "All Checked"
      }
      Else
      {
        $TmpMenuText = "Selected"
      }
    }
    Else
    {
      $TmpMenuText = "Zero"
    }
    
    $PILItemListContextMenuStrip.Items["Process"].Text = $PILItemListContextMenuStrip.Items["Process"].Tag -f $TmpMenuText
    $PILItemListContextMenuStrip.Items["Process"].Enabled = ($TmpMenuText -ne "Zero")
    $PILItemListContextMenuStrip.Items["Export"].Text = $PILItemListContextMenuStrip.Items["Export"].Tag -f $TmpMenuText
    $PILItemListContextMenuStrip.Items["Export"].Enabled = ($TmpMenuText -ne "Zero")
    $PILItemListContextMenuStrip.Items["Delete"].Text = $PILItemListContextMenuStrip.Items["Delete"].Tag -f $TmpMenuText
    $PILItemListContextMenuStrip.Items["Delete"].Enabled = ($TmpMenuText -ne "Zero")
    
    $PILItemListContextMenuStrip.Items["Check"].Enabled = ($TmpMenuText -ne "Zero")
    $PILItemListContextMenuStrip.Items["Uncheck"].Enabled = ($TmpMenuText -ne "Zero")
    
    $PILItemListContextMenuStrip.Items["AddCol"].Enabled = ($PILItemListListView.Columns.Count -lt ([MyRuntime]::MaxColumns - 1))
    $PILItemListContextMenuStrip.Items["RemoveCol"].Enabled = ($PILItemListListView.Columns.Count -gt ([MyRuntime]::MinColumns + 1))
    
    $PILItemListContextMenuStrip.Show($Sender, $EventArg.Location)
  }
  
  Write-Verbose -Message "Exit MouseDown Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItemListListViewMouseDown ********
$PILItemListListView.add_MouseDown({Start-PILItemListListViewMouseDown -Sender $This -EventArg $PSItem})

For ($I = 0; $I -lt ([MyRuntime]::CurrentColumns); $I++)
{
  $TmpColName = [MyRuntime]::ThreadConfig.ColumnNames[$I]
  $PILItemListListView.Columns.Insert($I, $TmpColName, $TmpColName, -2)
}
$PILItemListListView.Columns[0].Width = -2
$PILItemListListView.Columns.Insert([MyRuntime]::CurrentColumns, "Blank", " ", ($PILForm.Width * 4))


# ************************************************
# PILItemList ContextMenuStrip
# ************************************************
#region $PILItemListContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
$PILItemListContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
#$PILItemListListView.Controls.Add($PILItemListContextMenuStrip)
$PILItemListContextMenuStrip.BackColor = [MyConfig]::Colors.Back
$PILItemListContextMenuStrip.Enabled = $True
$PILItemListContextMenuStrip.Font = [MyConfig]::Font.Regular
$PILItemListContextMenuStrip.ForeColor = [MyConfig]::Colors.Fore
$PILItemListContextMenuStrip.ImageList = $PILSmallImageList
$PILItemListContextMenuStrip.ImageScalingSize = [System.Drawing.Size]::New(16, 16)
$PILItemListContextMenuStrip.Name = "PILItemListContextMenuStrip"
$PILItemListContextMenuStrip.ShowCheckMargin = $False
$PILItemListContextMenuStrip.ShowImageMargin = $True
$PILItemListContextMenuStrip.ShowItemToolTips = $False
#$PILItemListContextMenuStrip.TabIndex = 0
#$PILItemListContextMenuStrip.TabStop = $False
#$PILItemListContextMenuStrip.Tag = [System.Object]::New()
$PILItemListContextMenuStrip.TextDirection = [System.Windows.Forms.ToolStripTextDirection]::Horizontal
#endregion $PILItemListContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()


#region ******** Function Start-PILItemListContextMenuStripOpening ********
function Start-PILItemListContextMenuStripOpening
{
  <#
    .SYNOPSIS
      Opening Event for the PILItemList ContextMenuStrip Control
    .DESCRIPTION
      Opening Event for the PILItemList ContextMenuStrip Control
    .PARAMETER Sender
       The ItemList Control that fired the Opening Event
    .PARAMETER EventArg
       The Event Arguments for the ItemList Opening Event
    .EXAMPLE
       Start-PILItemListContextMenuStripOpening -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ContextMenuStrip]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Opening Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0
  
  # Do Not Show Context Menu if it is Disabled
  $EventArg.Cancel = (-not $PILItemListContextMenuStrip.Enabled)

  Write-Verbose -Message "Exit Opening Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItemListContextMenuStripOpening ********
$PILItemListContextMenuStrip.add_Opening({Start-PILItemListContextMenuStripOpening -Sender $This -EventArg $PSItem})

#region ******** Function Start-PILItemListContextMenuStripItemClick ********
Function Start-PILItemListContextMenuStripItemClick
{
  <#
    .SYNOPSIS
      ItemClicked Event for the PILItemList ContextMenuStrip Control
    .DESCRIPTION
      ItemClicked Event for the PILItemList ContextMenuStrip Control
    .PARAMETER Sender
       The ItemList Control that fired the ItemClicked Event
    .PARAMETER EventArg
       The Event Arguments for the ItemList ItemClicked Event
    .EXAMPLE
       Start-PILItemListContextMenuStripItemClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  Param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ToolStripItem]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter ItemClicked Event for $($MyInvocation.MyCommand)"
  
  [MyConfig]::AutoExit = 0
  
  # Play Sound
  #[System.Console]::Beep(2000, 10)
  
  $TmpLisTViewItems = @($PILItemListListView.SelectedItems)
  $TmpListText = "Selected"
  If ($TmpLisTViewItems[0].Checked)
  {
    $TmpLisTViewItems = @($PILItemListListView.CheckedItems)
    $TmpListText = "All Checked"
  }
  
  Switch ($Sender.Name)
  {
    "Process"
    {
      #region Start List Processing
      
      If ([String]::IsNullOrEmpty([MyRuntime]::ThreadConfig.ThreadScript))
      {
        $Response = Get-UserResponse -Title "No PIL Configureation" -Message "There is no PIL Script Configured!" -ButtonMid OK -ButtonDefault OK -Icon ([System.Drawing.SystemIcons]::Warning)
      }
      else
      {
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Start Processing $($TmpListText) Item List"
        $PILBtmStatusStrip.Refresh()
        
        # Disable Main Menu Iteme
        $PILTopMenuStrip.Items["AddItems"].Enabled = $False
        $PILTopMenuStrip.Items["Configure"].Enabled = $False
        $PILTopMenuStrip.Items["ProcessItems"].Enabled = $False
        $PILTopMenuStrip.Items["ListData"].Enabled = $False
        
        # Disable Right Click Menu
        $PILItemListContextMenuStrip.Enabled = $False
        
        # Disable ListView Sort
        $PILItemListListView.ListViewItemSorter.Enable = $False
        
        # Build RunSpace Pool
        $HashTable = @{ "ShowHeader" = $True; "ListItems" = $TmpLisTViewItems } 
        $ScriptBlock = { [CmdletBinding()] Param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Start-ProcessingItems -RichTextBox $RichTextBox -HashTable $HashTable }
        $DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title "Initializing $([MyConfig]::ScriptName)" -ButtonMid "OK" -HashTable $HashTable -AutoClose -AutoCloseWait 1
        
        # Set Processing ToolStrip Menu Items
        $PILPlayProcButton.Enabled = $False
        $PILPlayPauseButton.Enabled = $True
        $PILPlayStopButton.Enabled = $True
        $PILPlayBarPanel.Visible = $True
        $PILForm.Refresh()
        
        $PILBtmStatusStrip.Items["Status"].Text = "Processing $($TmpListText.Count) List Items"
        $PILBtmStatusStrip.Refresh()
        
        Monitor-RunspacePoolThreads
      }
      Break
      #endregion Start List Processing
    }
    "Export"
    {
      #region Export Slected / Checked Items
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Export $($TmpListText) Items"
      $PILBtmStatusStrip.Refresh()
      
      # Save Export File
      $PILSaveFileDialog.FileName = ""
      $PILSaveFileDialog.Filter = "CSV File (*.csv)|*.csv"
      $PILSaveFileDialog.FilterIndex = 1
      $PILSaveFileDialog.Title = "Export PIL CSV Report"
      $PILSaveFileDialog.Tag = $Null
      $Response = $PILSaveFileDialog.ShowDialog()
      If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
      {
        Try
        {
          $TmpCount = ([MyRuntime]::CurrentColumns - 1)
          $StringBuilder = [System.Text.StringBuilder]::New()
          [Void]$StringBuilder.AppendLine(($PILItemListListView.Columns[0..$($TmpCount)] | Select-Object -ExpandProperty Text) -Join ",")
          $TmpLisTViewItems | ForEach-Object -Process { [Void]$StringBuilder.AppendLine("`"{0}`"" -f (($PSItem.SubItems[0 .. $($TmpCount)] | Select-Object -ExpandProperty Text) -join "`",`"")) }
          ConvertFrom-Csv -InputObject (($StringBuilder.ToString())) -Delimiter "," | Export-Csv -Path $PILSaveFileDialog.FileName -NoTypeInformation -Encoding ASCII
          $StringBuilder.Clear()
          
          # Save Current Directory
          $PILSaveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILSaveFileDialog.FileName)
          $PILBtmStatusStrip.Items["Status"].Text = "Success Exporting $($TmpListText) Items"
        }
        Catch
        {
          $Response = Get-UserResponse -Title "Error Exporting" -Message "There was an Error Exporting the PIL Report Data" -ButtonMid OK -ButtonDefault OK -Icon ([System.Drawing.SystemIcons]::Error)
          $PILBtmStatusStrip.Items["Status"].Text = "Error Exporting $($TmpListText) Items"
        }
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Exporting $($TmpListText) Items"
      }
      Break
      #endregion Export Slected / Checked Items
    }
    "Header"
    {
      #region Resize Header
      $PILItemListListView.BeginUpdate()
      $PILItemListListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
      If ($PILItemListListView.Items.Count -gt 0)
      {
        $PILItemListListView.Columns[0].Width = -1
      }
      $PILItemListListView.Columns[([MyRuntime]::CurrentColumns)].Width = ($PILForm.Width * 4)
      $PILItemListListView.EndUpdate()
      Break
      #endregion Resize Header
    }
    "Content"
    {
      #region Resize Content
      $PILItemListListView.BeginUpdate()
      If ($PILItemListListView.Items.Count -eq 0)
      {
        $PILItemListListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
      }
      Else
      {
        $PILItemListListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
      }
      $PILItemListListView.Columns[([MyRuntime]::CurrentColumns)].Width = ($PILForm.Width * 4)
      $PILItemListListView.EndUpdate()
      Break
      #endregion Resize Content
    }
    "Check"
    {
      #region Check All
      $TmpChecked = @($PILItemListListView.Items | Where-Object -FilterScript { -not $PSItem.Checked })
      $TmpChecked | ForEach-Object -Process { $PSItem.Checked = $True }
      Break
      #endregion Check All
    }
    "UnCheck"
    {
      #region Uncheck All
      $TmpChecked = @($PILItemListListView.Items | Where-Object -FilterScript { $PSItem.Checked })
      $TmpChecked | ForEach-Object -Process { $PSItem.Checked = $False }
      Break
      #endregion Uncheck All
    }
    "Delete"
    {
      #region Clear Selected / Checked Item List
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Clear $($TmpListText) Items"
      $PILBtmStatusStrip.Refresh()
      
      $Response = Get-UserResponse -Title "Clear Item List?" -Message "Do you want to Clear the $($TmpListText) Items?" -ButtonLeft Yes -ButtonRight No -ButtonDefault No -Icon ([System.Drawing.SystemIcons]::Question)
      If ($Response.Success)
      {
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Clearing $($TmpListText) Items"
      }
      Else
      {
        # Clear Item List
        $TmpLisTViewItems | ForEach-Object {
          $PILItemListListView.Items.Remove($PSItem)
        }
        
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Successfully Cleared $($TmpListText) Items"
      }
      Break
      #endregion Clear Selected / Checked Item List
    }
    "AddCol"
    {
      #region Add PIL Column
      $PILBtmStatusStrip.Items["Status"].Text = "Add New PIL Column"
      $PILBtmStatusStrip.Refresh()
      
      $TmpColNames = @($PILItemListListView.Columns[0..([MyRuntime]::CurrentColumns - 1)] | ForEach-Object -Process { [MyListItem]::new($PSItem.Text, $PSItem.Index, $PSItem.Name) })
      $DialogResult = Add-NewPILColumn -Title "Add New PIL Column" -Message "Select Which Column to add New Column After?." -Items $TmpColNames -DisplayMember "Text" -ValueMember "Value" -SelectText "Select PIL Column to Add New Column After"
      If ($DialogResult.Success)
      {
        $AfterIndex = $DialogResult.Index + 1
        $AddColName = $DialogResult.Name
        
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = $Null
        $PILItemListListView.BeginUpdate()
        $PILItemListListView.Columns.Insert($AfterIndex, $AddColName, $AddColName, -2)
        ForEach ($ListItem In $PILItemListListView.Items)
        {
          $ListItem.SubItems.Insert($AfterIndex, ([System.Windows.Forms.ListViewItem+ListViewSubItem]::new($ListItem, "", [MyConfig]::Colors.TextFore, [MyConfig]::Colors.TextBack, [MyConfig]::Font.Regular)))
        }
        $PILItemListListView.EndUpdate()
        
        # Update Column Names
        [MyRuntime]::AddPILColumn($AfterIndex, $AddColName)
        
        # Update Status
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = "Selected16Icon"
        $PILBtmStatusStrip.Items["Status"].Text = "Success Adding PIL Column $($RemoveName)"
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Adding PIL Column"
      }
      Break
      #endregion Add PIL Column
    }
    "RemoveCol"
    {
      #region Remove PIL Columns
      $PILBtmStatusStrip.Items["Status"].Text = "Remove Existing PIL Column"
      $PILBtmStatusStrip.Refresh()
      $TmpColNames = @($PILItemListListView.Columns[1..([MyRuntime]::CurrentColumns - 1)] | ForEach-Object -Process { [MyListItem]::new($PSItem.Text, $PSItem.Index, $PSItem.Name) })
      
      $DialogResult = Get-CheckedListBoxOprion -Title "Remove Existing Columns" -Message "Select Which Column do you would like to Remove." -DisplayMember "Text" -ValueMember "Value" -Items $TmpColNames
      If ($DialogResult.Success)
      {
        # Success
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = $Null
        $PILItemListListView.BeginUpdate()
        
        $TmpColumns = $DialogResult.Items | Sort-Object -Property Value -Descending
        :HitLimit ForEach ($TmpColumn In $TmpColumns)
        {
          $PILItemListListView.Columns.RemoveAt($TmpColumn.Value)
          ForEach ($ListItem In $PILItemListListView.Items)
          {
            $ListItem.SubItems.RemoveAt($TmpColumn.Value)
          }
          
          # Update Column Names
          [MyRuntime]::RemovePILComun($TmpColumn.Value)
          If ([MyRuntime]::CurrentColumns -le [MyRuntime]::MinColumns)
          {
            Break HitLimit
          }
        }
        
        # Resize Columns
        $PILItemListListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
        If ($PILItemListListView.Items.Count -gt 0)
        {
          $PILItemListListView.Columns[0].Width = -1
        }
        $PILItemListListView.Columns[([MyRuntime]::CurrentColumns)].Width = ($PILForm.Width * 4)
        $PILItemListListView.EndUpdate()
        
        # Update Status
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = "Selected16Icon"
        $PILBtmStatusStrip.Items["Status"].Text = "Success Removing PIL Column"
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Removing PIL Column"
      }
      Break
      #endregion Remove PIL Columns
    }
    "Reset"
    {
      #region Reset PIL
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Reseting $([MyConfig]::ScriptName)"
      $PILBtmStatusStrip.Refresh()
      
      $Response = Get-UserResponse -Title "Reset All?" -Message "Do you want to Reset $([MyConfig]::ScriptName) to Default Settings??" -ButtonLeft Yes -ButtonRight No -ButtonDefault No -Icon ([System.Drawing.SystemIcons]::Question)
      If ($Response.Success)
      {
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Reseting $([MyConfig]::ScriptName)"
      }
      Else
      {
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = $Null
        
        [MyRuntime]::UpdateTotalColumn([MyRuntime]::StartColumns)
        $PILItemListListView.BeginUpdate()
        $PILItemListListView.Columns.Clear()
        $PILItemListListView.Items.Clear()
        For ($I = 0; $I -lt ([MyRuntime]::CurrentColumns); $I++)
        {
          $TmpColName = [MyRuntime]::ThreadConfig.ColumnNames[$I]
          $PILItemListListView.Columns.Insert($I, $TmpColName, $TmpColName, -2)
        }
        $PILItemListListView.Columns[0].Width = -2
        $PILItemListListView.Columns.Insert([MyRuntime]::CurrentColumns, "Blank", " ", ($PILForm.Width * 4))
        $PILItemListListView.EndUpdate()
        [MyRuntime]::ConfigName = "Unknown Configuration"
 
        # Set Status Message
        $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = "Selected16Icon"
        $PILBtmStatusStrip.Items["Status"].Text = "Successfully Reset $([MyConfig]::ScriptName)"
      }
      Break
      #endregion Reset PIL
    }
  }
  
  Write-Verbose -Message "Exit ItemClicked Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItemListContextMenuStripItemClick ********

(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Process" -Name "Process" -Tag "Process {0} Items" -DisplayStyle "ImageAndText" -ImageKey "Process16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Export" -Name "Export" -Tag "Export {0} Items" -DisplayStyle "ImageAndText" -ImageKey "Export16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $PILItemListContextMenuStrip
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Check All" -Name "Check" -Tag "Check" -DisplayStyle "ImageAndText" -ImageKey "CheckIcon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Uncheck All" -Name "Uncheck" -Tag "Uncheck" -DisplayStyle "ImageAndText" -ImageKey "UnCheckIcon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $PILItemListContextMenuStrip
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Delete" -Name "Delete" -Tag "Delete {0} Items" -DisplayStyle "ImageAndText" -ImageKey "Trash16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $PILItemListContextMenuStrip
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Resize Header" -Name "Header" -Tag "Header" -DisplayStyle "ImageAndText" -ImageKey "Header16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Resize Content" -Name "Content" -Tag "Content" -DisplayStyle "ImageAndText" -ImageKey "Content16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $PILItemListContextMenuStrip
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Add Column" -Name "AddCol" -Tag "AddCol" -DisplayStyle "ImageAndText" -ImageKey "AddCol16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Remove Columns" -Name "RemoveCol" -Tag "RemoveCol" -DisplayStyle "ImageAndText" -ImageKey "RemoveCol16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $PILItemListContextMenuStrip
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Reset PIL" -Name "Reset" -Tag "Reset" -DisplayStyle "ImageAndText" -ImageKey "PILFormIcon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})

#endregion ******** $PILMainPanel Controls ********

# ************************************************
# PILPlayBar Panel
# ************************************************
#region $PILPlayBarPanel = [System.Windows.Forms.Panel]::New()
$PILPlayBarPanel = [System.Windows.Forms.Panel]::New()
$PILForm.Controls.Add($PILPlayBarPanel)
$PILPlayBarPanel.BackColor = [MyConfig]::Colors.Back
$PILPlayBarPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$PILPlayBarPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$PILPlayBarPanel.Enabled = $True
$PILPlayBarPanel.Font = [MyConfig]::Font.Regular
$PILPlayBarPanel.ForeColor = [MyConfig]::Colors.Fore
$PILPlayBarPanel.Name = "PILPlayBarPanel"
#$PILPlayBarPanel.TabIndex = 0
#$PILPlayBarPanel.TabStop = $False
#$PILPlayBarPanel.Tag = [System.Object]::New()
$PILPlayBarPanel.Visible = $False
#endregion $PILPlayBarPanel = [System.Windows.Forms.Panel]::New()

#region ******** PILPlayBar Controls ********

# ************************************************
# PILPlayCtrls Panel
# ************************************************
#region $PILPlayPanel = [System.Windows.Forms.Panel]::New()
$PILPlayCtrlsPanel = [System.Windows.Forms.Panel]::New()
$PILPlayBarPanel.Controls.Add($PILPlayCtrlsPanel)
$PILPlayCtrlsPanel.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
$PILPlayCtrlsPanel.BackColor = [MyConfig]::Colors.Back
$PILPlayCtrlsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$PILPlayCtrlsPanel.Enabled = $True
$PILPlayCtrlsPanel.Font = [MyConfig]::Font.Regular
$PILPlayCtrlsPanel.ForeColor = [MyConfig]::Colors.Fore
$PILPlayCtrlsPanel.Name = "PILPlayCtrlsPanel"
#$PILPlayCtrlsPanel.TabIndex = 0
#$PILPlayCtrlsPanel.TabStop = $False
#$PILPlayCtrlsPanel.Tag = [System.Object]::New()
#endregion $PILPlayCtrlsPanel = [System.Windows.Forms.Panel]::New()

#region ******** PILPlayCtrls Controls ********

$TmpButtonSize = [System.Drawing.Size]::New(($PILLargeImageList.ImageSize.Width + [MyConfig]::FormSpacer), ($PILLargeImageList.ImageSize.Height + [MyConfig]::FormSpacer))

#region $PILPLayProcButton = [System.Windows.Forms.Button]::New()
$PILPLayProcButton = [System.Windows.Forms.Button]::New()
$PILPlayCtrlsPanel.Controls.Add($PILPLayProcButton)
$PILPLayProcButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$PILPlayProcButton.BackColor = [System.Drawing.Color]::Transparent
$PILPLayProcButton.Enabled = $True
$PILPlayProcButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$PILPlayProcButton.FlatAppearance.BorderSize = 0
$PILPlayProcButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Transparent
$PILPlayProcButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Transparent
$PILPLayProcButton.Font = [MyConfig]::Font.Bold
$PILPLayProcButton.ForeColor = [MyConfig]::Colors.ButtonFore
$PILPlayProcButton.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$PILPLayProcButton.ImageKey = "Play48Icon"
$PILPLayProcButton.ImageList = $PILLargeImageList
$PILPLayProcButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
$PILPLayProcButton.Name = "PILPLayProcButton"
$PILPLayProcButton.TabStop = $True
$PILPLayProcButton.Size = $TmpButtonSize
#endregion $PILPLayProcButton = [System.Windows.Forms.Button]::New()

#region ******** Function Start-PILPlayProcButtonClick ********
function Start-PILPlayProcButtonClick
{
  <#
    .SYNOPSIS
      Click Event for the PILPlayProc Button Control
    .DESCRIPTION
      Click Event for the PILPlayProc Button Control
    .PARAMETER Sender
       The PlayProc Control that fired the Click Event
    .PARAMETER EventArg
       The Event Arguments for the PlayProc Click Event
    .EXAMPLE
       Start-PILPlayProcButtonClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Button]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Click Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0
  
  # Set Status Message
  $PILBtmStatusStrip.Items["Status"].Text = "Resume Processing List Items"
  $PILBtmStatusStrip.Refresh()
  
  # Play Sound
  [System.Console]::Beep(1000, 30)
  
  # Uppause Runspace Pool Threads
  (Get-MyRSPool).SyncedHash["Pause"] = $False
  
  # Set Processing ToolStrip Menu Items
  $PILPlayPauseButton.Enabled = $True
  $PILPlayStopButton.Enabled = $True
  $PILPlayProcButton.Enabled = $False
  
  # Set Status Message
  $PILBtmStatusStrip.Items["Status"].Text = "Processing Item List has Resumed"
  
  Write-Verbose -Message "Exit Click Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILPlayProcButtonClick ********
$PILPlayProcButton.add_Click({Start-PILPlayProcButtonClick -Sender $This -EventArg $PSItem})

#region $PILPLayPauseButton = [System.Windows.Forms.Button]::New()
$PILPLayPauseButton = [System.Windows.Forms.Button]::New()
$PILPlayCtrlsPanel.Controls.Add($PILPLayPauseButton)
$PILPLayPauseButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$PILPLayPauseButton.BackColor = [System.Drawing.Color]::Transparent
$PILPLayPauseButton.Enabled = $True
$PILPLayPauseButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$PILPLayPauseButton.FlatAppearance.BorderSize = 0
$PILPLayPauseButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Transparent
$PILPLayPauseButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Transparent
$PILPLayPauseButton.Font = [MyConfig]::Font.Bold
$PILPLayPauseButton.ForeColor = [MyConfig]::Colors.ButtonFore
$PILPLayPauseButton.ImageKey = "Pause48Icon"
$PILPLayPauseButton.ImageList = $PILLargeImageList
$PILPLayPauseButton.Location = [System.Drawing.Point]::New(($PILPLayProcButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
$PILPLayPauseButton.Name = "PILPLayPauseButton"
$PILPLayPauseButton.TabStop = $True
$PILPLayPauseButton.Size = $TmpButtonSize
#endregion $PILPLayPauseButton = [System.Windows.Forms.Button]::New()

#region ******** Function Start-PILPlayPauseButtonClick ********
function Start-PILPlayPauseButtonClick
{
  <#
    .SYNOPSIS
      Click Event for the PILPlayPause Button Control
    .DESCRIPTION
      Click Event for the PILPlayPause Button Control
    .PARAMETER Sender
       The PlayPause Control that fired the Click Event
    .PARAMETER EventArg
       The Event Arguments for the PlayPause Click Event
    .EXAMPLE
       Start-PILPlayPauseButtonClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Button]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Click Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0
  
  # Set Status Message
  $PILBtmStatusStrip.Items["Status"].Text = "Pause Processing List Items"
  $PILBtmStatusStrip.Refresh()
  
  # Play Sound
  [System.Console]::Beep(1000, 30)
  
  # Pause Runspace Pol Threads
  (Get-MyRSPool).SyncedHash["Pause"] = $True
  
  # Set Pauseing ToolStrip Menu Items
  $PILPlayPauseButton.Enabled = $False
  $PILPlayProcButton.Enabled = $True
  $PILPlayStopButton.Enabled = $True
  
  # Set Status Message
  $PILBtmStatusStrip.Items["Status"].Text = "Processing Item List has been Paused"
  
  Write-Verbose -Message "Exit Click Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILPlayPauseButtonClick ********
$PILPlayPauseButton.add_Click({Start-PILPlayPauseButtonClick -Sender $This -EventArg $PSItem})

#region $PILPLayStopButton = [System.Windows.Forms.Button]::New()
$PILPLayStopButton = [System.Windows.Forms.Button]::New()
$PILPlayCtrlsPanel.Controls.Add($PILPLayStopButton)
$PILPLayStopButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$PILPLayStopButton.BackColor = [System.Drawing.Color]::Transparent
$PILPLayStopButton.Enabled = $True
$PILPLayStopButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$PILPLayStopButton.FlatAppearance.BorderSize = 0
$PILPLayStopButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Transparent
$PILPLayStopButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Transparent
$PILPLayStopButton.Font = [MyConfig]::Font.Bold
$PILPLayStopButton.ForeColor = [MyConfig]::Colors.ButtonFore
$PILPlayStopButton.ImageKey = "Stop48Icon"
$PILPlayStopButton.ImageList = $PILLargeImageList
$PILPLayStopButton.Location = [System.Drawing.Point]::New(($PILPLayPauseButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
$PILPLayStopButton.Name = "PILPLayStopButton"
$PILPLayStopButton.TabStop = $True
$PILPLayStopButton.Size = $TmpButtonSize
#endregion $PILPLayStopButton = [System.Windows.Forms.Button]::New()

#region ******** Function Start-PILPlayStopButtonClick ********
function Start-PILPlayStopButtonClick
{
  <#
    .SYNOPSIS
      Click Event for the PILPlayStop Button Control
    .DESCRIPTION
      Click Event for the PILPlayStop Button Control
    .PARAMETER Sender
       The PlayStop Control that fired the Click Event
    .PARAMETER EventArg
       The Event Arguments for the PlayStop Click Event
    .EXAMPLE
       Start-PILPlayStopButtonClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Button]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Click Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0
  
  # Set Status Message
  $PILBtmStatusStrip.Items["Status"].Text = "Terminate Proccessing List Items"
  $PILBtmStatusStrip.Refresh()
  
  # Play Sound
  [System.Console]::Beep(1000, 30)
  
  # Play Sound
  #[System.Console]::Beep(1000, 30)
  
  # Terminate Threads
  $TmpRSPool = Get-MyRSPool
  $TmpRSPool.SyncedHash["Terminate"] = $True
  
  $PILPlayStopButton.Enabled = $False
  $PILPlayPauseButton.Enabled = $True
  $PILPlayProcButton.Enabled = $False
  $PILPlayStopButton.Enabled = $True
  
  $TmpRSPool.SyncedHash["Pause"] = $False
  
  # Set Status Message
  $PILBtmStatusStrip.Items["Status"].Text = "Proccessing List Items has been Terminated"
  
  Write-Verbose -Message "Exit Click Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILPlayStopButtonClick ********
$PILPlayStopButton.add_Click({Start-PILPlayStopButtonClick -Sender $This -EventArg $PSItem})

$PILPlayCtrlsPanel.ClientSize = [System.Drawing.Size]::New(($($PILPlayCtrlsPanel.Controls[$PILPlayCtrlsPanel.Controls.Count - 1]).Right + [MyConfig]::FormSpacer), $($PILPlayCtrlsPanel.Controls[$PILPlayCtrlsPanel.Controls.Count - 1]).Bottom)
$PILPlayCtrlsPanel.Location = [System.Drawing.Point]::New((($PILPlayBarPanel.ClientSize.Width - $PILPlayCtrlsPanel.Width) / 2), 0)
$PILPlayBarPanel.ClientSize = [System.Drawing.Size]::New($PILPlayBarPanel.ClientSize.Width, $PILPlayCtrlsPanel.Height)

#endregion ******** PILPlayCtrls Controls ********

#region $PILLeftProgressBar = [System.Windows.Forms.ProgressBar]::New()
$PILLeftProgressBar = [System.Windows.Forms.ProgressBar]::New()
$PILPlayBarPanel.Controls.Add($PILLeftProgressBar)
$PILLeftProgressBar.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
$PILLeftProgressBar.BackColor = [MyConfig]::Colors.Back
$PILLeftProgressBar.ForeColor = [MyConfig]::Colors.Fore
$PILLeftProgressBar.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, (($PILPlayCtrlsPanel.ClientSize.Height - $PILLeftProgressBar.Height) / 2))
$PILLeftProgressBar.Maximum = 100
$PILLeftProgressBar.Minimum = 0
$PILLeftProgressBar.Name = "PILLeftProgressBar"
$PILLeftProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
$PILLeftProgressBar.TabStop = $False
$PILLeftProgressBar.Value = 0
$PILLeftProgressBar.Width = ($PILPlayCtrlsPanel.Left - [MyConfig]::FormSpacer)
#endregion $PILLeftProgressBar = [System.Windows.Forms.ProgressBar]::New()

#region $PILRightProgressBar = [System.Windows.Forms.ProgressBar]::New()
$PILRightProgressBar = [System.Windows.Forms.ProgressBar]::New()
$PILPlayBarPanel.Controls.Add($PILRightProgressBar)
$PILRightProgressBar.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
$PILRightProgressBar.BackColor = [MyConfig]::Colors.Back
$PILRightProgressBar.ForeColor = [MyConfig]::Colors.Fore
$PILRightProgressBar.Location = [System.Drawing.Point]::New(($PILPlayCtrlsPanel.Right + [MyConfig]::FormSpacer), (($PILPlayCtrlsPanel.ClientSize.Height - $PILRightProgressBar.Height) / 2))
$PILRightProgressBar.Maximum = 100
$PILRightProgressBar.Minimum = 0
$PILRightProgressBar.Name = "PILRightProgressBar"
$PILRightProgressBar.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$PILRightProgressBar.RightToLeftLayout = $True
$PILRightProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
$PILRightProgressBar.TabStop = $False
$PILRightProgressBar.Value = 0
$PILRightProgressBar.Width = ($PILPlayBarPanel.ClientSize.Width - ($PILRightProgressBar.Left + [MyConfig]::FormSpacer))
#endregion $PILRightProgressBar = [System.Windows.Forms.ProgressBar]::New()

#region ******** Function Start-PILPlayBarPanelResize ********
function Start-PILPlayBarPanelResize
{
  <#
    .SYNOPSIS
      Resize Event for the PILPlayBar Panel Control
    .DESCRIPTION
      Resize Event for the PILPlayBar Panel Control
    .PARAMETER Sender
       The PlayBar Control that fired the Resize Event
    .PARAMETER EventArg
       The Event Arguments for the PlayBar Resize Event
    .EXAMPLE
       Start-PILPlayBarPanelResize -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.Panel]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Resize Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0
  
  $PILLeftProgressBar.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, (($PILPlayCtrlsPanel.ClientSize.Height - $PILLeftProgressBar.Height) / 2))
  $PILLeftProgressBar.Width = ($PILPlayCtrlsPanel.Left - [MyConfig]::FormSpacer)
  
  $PILRightProgressBar.Location = [System.Drawing.Point]::New(($PILPlayCtrlsPanel.Right + [MyConfig]::FormSpacer), (($PILPlayCtrlsPanel.ClientSize.Height - $PILRightProgressBar.Height) / 2))
  $PILRightProgressBar.Width = ($PILPlayBarPanel.ClientSize.Width - ($PILRightProgressBar.Left + [MyConfig]::FormSpacer))
  
  #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

  Write-Verbose -Message "Exit Resize Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILPlayBarPanelResize ********
$PILPlayBarPanel.add_Resize({Start-PILPlayBarPanelResize -Sender $This -EventArg $PSItem})

#endregion ******** PILPlayBar Controls ********

# ************************************************
# PILTop MenuStrip
# ************************************************
#region $PILTopMenuStrip = [System.Windows.Forms.MenuStrip]::New()
$PILTopMenuStrip = [System.Windows.Forms.MenuStrip]::New()
$PILForm.Controls.Add($PILTopMenuStrip)
$PILForm.MainMenuStrip = $PILTopMenuStrip
$PILTopMenuStrip.BackColor = [MyConfig]::Colors.Back
$PILTopMenuStrip.Dock = [System.Windows.Forms.DockStyle]::Top
$PILTopMenuStrip.Enabled = $True
$PILTopMenuStrip.Font = [MyConfig]::Font.Regular
$PILTopMenuStrip.ForeColor = [MyConfig]::Colors.Fore
$PILTopMenuStrip.ImageList = $PILSmallImageList
$PILTopMenuStrip.ImageScalingSize = [System.Drawing.Size]::New(16, 16)
$PILTopMenuStrip.Name = "PILTopMenuStrip"
$PILTopMenuStrip.ShowItemToolTips = $False
#$PILTopMenuStrip.TabIndex = 0
#$PILTopMenuStrip.TabStop = $False
#$PILTopMenuStrip.Tag = [System.Object]::New()
$PILTopMenuStrip.TextDirection = [System.Windows.Forms.ToolStripTextDirection]::Horizontal
#endregion $PILTopMenuStrip = [System.Windows.Forms.MenuStrip]::New()

#region ******** Function Start-PILTopMenuStripItemClick ********
function Start-PILTopMenuStripItemClick
{
  <#
    .SYNOPSIS
      Click Event for the PILTop MenuStripItem Control
    .DESCRIPTION
      Click Event for the PILTop MenuStripItem Control
    .PARAMETER Sender
       The Top Control that fired the Click Event
    .PARAMETER EventArg
       The Event Arguments for the Top Click Event
    .EXAMPLE
       Start-PILTopMenuStripItemClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By kensw
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ToolStripItem]$Sender,
    [parameter(Mandatory = $True)]
    [Object]$EventArg
  )
  Write-Verbose -Message "Enter Click Event for $($MyInvocation.MyCommand)"

  [MyConfig]::AutoExit = 0
  
  # Play Sound
  #[System.Console]::Beep(2000, 10)
  
  Switch ($Sender.Name)
  {
    "AddList"
    {
      #region Add New Items List
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Add New Items for Processing"
      $PILBtmStatusStrip.Refresh()
      
      $DialogResult = Get-TextBoxInput -Title "Get Item List" -Message "Enter the list of items to add for processing" -Multi -NoDuplicates -ValidChars "."
      If ($DialogResult.Success)
      {
        $NewCount = 0
        $TmpSubItems = @("") * [MyRuntime]::CurrentColumns
        $PILItemListListView.BeginUpdate()
        ForEach ($TmpItem In $DialogResult.Items)
        {
          If (-not $PILItemListListView.Items.ContainsKey($TmpItem))
          {
            $TmpListItem = [System.Windows.Forms.ListViewItem]::New($TmpItem, "StatusInfo16Icon", [MyConfig]::Colors.TextFore, [MyConfig]::Colors.TextBack, [MyConfig]::Font.Regular)
            $TmpListItem.Name = $TmpItem
            $TmpListItem.SubItems.AddRange($TmpSubItems)
            [Void]$PILItemListListView.Items.Add($TmpListItem)
            $NewCount++
          }
        }
        $PILItemListListView.EndUpdate()
        
        # Success
        $PILBtmStatusStrip.Items["Status"].Text = "Successfully Added $($NewCount) New Items for Processing"
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Adding New Items for Processing"
      }
      
      Break
      #endregion Add New Items List
    }
    "ImportList"
    {
      #region Import Item List
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Importing Item List"
      $PILBtmStatusStrip.Refresh()
      
      # Get File to Import
      $PILOpenFileDialog.FileName = ""
      $PILOpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|Supported Files|*.csv;*.txt|All Files (*.*)|*.*"
      $PILOpenFileDialog.FilterIndex = 3
      $PILOpenFileDialog.Multiselect = $False
      $PILOpenFileDialog.Title = "Select an Item List Import File"
      $PILOpenFileDialog.Tag = $Null
      $Response = $PILOpenFileDialog.ShowDialog()
      If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
      {
        Try
        {
          If ([System.IO.path]::GetExtension($PILOpenFileDialog.SafeFileName) -eq ".csv")
          {
            $TmpCSV = Import-Csv -Path $PILOpenFileDialog.FileName
            If ($TmpCSV.Count -gt 0)
            {
              $TmpColNames = @($TmpCSV[0].PSObject.Properties)
              If ($TmpColNames.Count -gt 1)
              {
                $DialogResult = Get-ComboBoxOption -Title "Select Item Column Name" -Message "Select the CSV Column Name that has the Lost of Items you want to Import." -SelectText "Select the CSV Column Name" -Items $TmpColNames -DisplayMember "Name" -ValueMember "MemberType"
                If ($DialogResult.Success)
                {
                  $TmpColName = $DialogResult.Item.Name
                }
                Else
                {
                  $TmpColName = "No Column Was Selected"
                }
              }
              Else
              {
                $TmpColName = $TmpColNames[0].Name
              }
              
              $TmpItems = @($TmpCSV | Select-Object -ExpandProperty $TmpColName)
            }
          }
          Else
          {
            $TmpItems = @(Get-Content -Path $PILOpenFileDialog.FileName)
          }
          
          If ($TmpItems.Count -eq 0)
          {
            $PILBtmStatusStrip.Items["Status"].Text = "No Items were Found to Import"
          }
          Else
          {
            $NewCount = 0
            $TmpSubItems = @("") * [MyRuntime]::CurrentColumns
            $PILItemListListView.BeginUpdate()
            ForEach ($TmpItem In $TmpItems)
            {
              If (-not $PILItemListListView.Items.ContainsKey($TmpItem))
              {
                $TmpListItem = [System.Windows.Forms.ListViewItem]::New($TmpItem, "StatusInfo16Icon", [MyConfig]::Colors.TextFore, [MyConfig]::Colors.TextBack, [MyConfig]::Font.Regular)
                $TmpListItem.Name = $TmpItem
                $TmpListItem.SubItems.AddRange($TmpSubItems)
                [Void]$PILItemListListView.Items.Add($TmpListItem)
                $NewCount++
              }
            }
            $PILItemListListView.Columns[0].Width = -2
            $PILItemListListView.EndUpdate()
            
            # Success
            $PILBtmStatusStrip.Items["Status"].Text = "Successfully Added $($NewCount) New Items for Processing"
          }
        }
        Catch
        {
          $PILBtmStatusStrip.Items["Status"].Text = "Error Importing Item List"
        }
        
        # Save Current Directory
        $PILOpenFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILOpenFileDialog.FileName)
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Importing Item List"
      }
      Break
      #endregion Import Item List
    }
    "LoadExport"
    {
      #region Load Exported Data
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Load Exported PIL Data"
      $PILBtmStatusStrip.Refresh()
      
      # Get File to Import
      $PILOpenFileDialog.FileName = ""
      $PILOpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
      $PILOpenFileDialog.FilterIndex = 1
      $PILOpenFileDialog.Multiselect = $False
      $PILOpenFileDialog.Title = "Select a PIL Export Data File"
      $PILOpenFileDialog.Tag = $Null
      $Response = $PILOpenFileDialog.ShowDialog()
      If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
      {
        $HashTable = @{"ShowHeader" = $True; "ImportFile" = $PILOpenFileDialog.FileName }
        $ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Load-PILDataExport -RichTextBox $RichTextBox -HashTable $HashTable }
        $DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title "Initializing $([MyConfig]::ScriptName)" -ButtonMid "OK" -HashTable $HashTable
        If ($DialogResult.Success)
        {
          $PILItemListListView.Columns[0].Width = -2
          $PILBtmStatusStrip.Items["Status"].Text = "Successfully Loaded Exported PIL Data"
        }
        Else
        {
          $PILBtmStatusStrip.Items["Status"].Text = "Errors Loading Exported PIL Data"
        }
        
        # Save Current Directory
        $PILOpenFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILOpenFileDialog.FileName)
      }
      Break
      #endregion Load Exported Data
    }
    "TotalColumns"
    {
      #region Set Total Columns
      $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = $Null
      [MyRuntime]::UpdateTotalColumn($Sender.Tag)
      $PILItemListListView.BeginUpdate()
      $PILItemListListView.Columns.Clear()
      If ($PILItemListListView.Items.Count -gt 0)
      {
        $TmpWidth =  -1
      }
      Else
      {
        $TmpWidth = -2
      }
      For ($I = 0; $I -lt ([MyRuntime]::CurrentColumns); $I++)
      {
        $TmpColName = [MyRuntime]::ThreadConfig.ColumnNames[$I]
        $PILItemListListView.Columns.Insert($I, $TmpColName, $TmpColName, $TmpWidth)
      }
      $PILItemListListView.Columns[0].Width = $TmpWidth
      $PILItemListListView.Columns.Insert([MyRuntime]::CurrentColumns, "Blank", " ", ($PILForm.Width * 4))
      $PILItemListListView.EndUpdate()
      
      [MyRuntime]::ConfigName = "Unknown Configuration"
      $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = "Selected16Icon"
      #endregion Set Total Columns
    }
    "ColumnNames"
    {
      #region Set Column Names
      $PILBtmStatusStrip.Items["Status"].Text = "Update Column Names"
      $PILBtmStatusStrip.Refresh()
      $DialogResult = Get-MultiTextBoxInput -Title "Update Column Names" -Message "Enter the New Column Names for the $([MyConfig]::ScriptName) Utility" -OrderedItems ([MyRuntime]::ThreadConfig.GetColumnNames()) -AllRequired -ValidChars "."
      If ($DialogResult.Success)
      {
        # Success
        $TmpNames = @($DialogResult.OrderedItems.Values)
        $Max = $TmpNames.Count
        For ($I = 0; $I -lt $Max; $I++)
        {
          $PILItemListListView.Columns[$I].Text = $TmpNames[$I]
        }
        [MyRuntime]::ThreadConfig.SetColumnNames($TmpNames)
        $PILBtmStatusStrip.Items["Status"].Text = "Successfully Updated Column Names"
      }
      Else
      {
        # Failed
        $PILBtmStatusStrip.Items["Status"].Text = "Failed to Update Column Names"
      }
      Break
      #endregion Set Column Names
    }
    "ThreadConfig"
    {
      #region Update Thread Config
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Update PIL Threads Configuration"
      $PILBtmStatusStrip.Refresh()
      $DialogResult = Update-ThreadConfiguration
      If ($DialogResult.Success)
      {
        $Response = Get-UserResponse -Title "Save Configuration" -Message "Would you like to Save the PIL Configuration?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
        If ($Response.Success)
        {
          # Save Export File
          $PILSaveFileDialog.FileName = [MyRuntime]::ConfigName
          $PILSaveFileDialog.Filter = "PIL Config Files|*.Json;*.Xml"
          $PILSaveFileDialog.FilterIndex = 1
          $PILSaveFileDialog.Title = "Save PIL Configuration File"
          $PILSaveFileDialog.Tag = $Null
          $Response = $PILSaveFileDialog.ShowDialog()
          If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
          {
            Try
            {
              If ([System.IO.Path]::GetExtension($PILSaveFileDialog.FileName) -eq ".Json")
              {
                # Save Config
                [MyRuntime]::ThreadConfig | ConvertTo-Json -Compress | Out-File -FilePath $PILSaveFileDialog.FileName -Encoding ASCII
              }
              Else
              {
                # Save Config
                [MyRuntime]::ThreadConfig | Export-Clixml -Path $PILSaveFileDialog.FileName -Encoding ASCII
              }
              
              # Update PIL Config Name
              [MyRuntime]::ConfigName = [System.IO.Path]::GetFileName($PILSaveFileDialog.FileName)
              
              # Save Current Directory
              $PILSaveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILSaveFileDialog.FileName)
            }
            Catch
            {
              $Response = Get-UserResponse -Title "Error Saving Config" -Message "There was an Error Saving the PIL Confiuration" -ButtonMid OK -ButtonDefault OK -Icon ([System.Drawing.SystemIcons]::Error)
              $PILBtmStatusStrip.Items["Status"].Text = "Error Exporting CSV Report"
            }
          }
        }
        $PILBtmStatusStrip.Items["Status"].Text = "Successfully Updated PIL Threads Configuration"
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Updating PIL Threads Configuration"
      }
      Break
      #endregion Update Thread Config
    }
    "LoadConfig"
    {
      #region Load PIL Configuration
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Load PIL Configuration File"
      $PILBtmStatusStrip.Refresh()
      
      # Open Selected File
      $PILOpenFileDialog.FileName = ""
      $PILOpenFileDialog.Filter = "PIL Config Files|*.Json;*.Xml"
      $PILOpenFileDialog.FilterIndex = 1
      $PILOpenFileDialog.Multiselect = $False
      $PILOpenFileDialog.Title = "Load PIL Configuration File"
      $PILOpenFileDialog.Tag = $Null
      $Response = $PILOpenFileDialog.ShowDialog()
      If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
      {
        $HashTable = @{"ShowHeader" = $True; "ConfigFile" = $PILOpenFileDialog.FileName }
        if ([System.IO.Path]::GetExtension($PILOpenFileDialog.FileName) -eq ".Json")
        {
          $ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Load-PILConfigFIleJson -RichTextBox $RichTextBox -HashTable $HashTable }
        }
        else
        {
          $ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Load-PILConfigFIleXml -RichTextBox $RichTextBox -HashTable $HashTable }
        }
        $DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title ($PILBtmStatusStrip.Items["Status"].Text) -ButtonMid "OK" -HashTable $HashTable
        
        If ($DialogResult.Success)
        {
          # Update PIL Config Name
          [MyRuntime]::ConfigName = [System.IO.Path]::GetFileName($PILOpenFileDialog.FileName)
          
          # Save Current Directory
          $PILOpenFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILOpenFileDialog.FileName)
          $PILBtmStatusStrip.Items["Status"].Text = "Success Loading PIL Configuration File"
        }
        Else
        {
          $PILBtmStatusStrip.Items["Status"].Text = "Errors Loading PIL Configuration File"
        }
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Loading PIL Configuration File"
      }
      Break
      #endregion Load PIL Configuration
    }
    "SaveConfig"
    {
      #region Save PIL Configuration File
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Save PIL Configuration File"
      $PILBtmStatusStrip.Refresh()
      
      # Save Export File
      $PILSaveFileDialog.FileName = [MyRuntime]::ConfigName
      $PILSaveFileDialog.Filter = "PIL Config Files|*.Json;*.Xml"
      $PILSaveFileDialog.FilterIndex = 1
      $PILSaveFileDialog.Title = "Save PIL Configuration File"
      $PILSaveFileDialog.Tag = $Null
      $Response = $PILSaveFileDialog.ShowDialog()
      If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
      {
        Try
        {
          If ([System.IO.Path]::GetExtension($PILSaveFileDialog.FileName) -eq ".Json")
          {
            # Save Config
            [MyRuntime]::ThreadConfig | ConvertTo-Json -Compress | Out-File -FilePath $PILSaveFileDialog.FileName -Encoding ASCII
          }
          Else
          {
            # Save Config
            [MyRuntime]::ThreadConfig | Export-Clixml -Path $PILSaveFileDialog.FileName -Encoding ASCII
          }
          
          # Update PIL Config Name
          [MyRuntime]::ConfigName = [System.IO.Path]::GetFileName($PILSaveFileDialog.FileName)
          
          # Save Current Directory
          $PILSaveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILSaveFileDialog.FileName)
          $PILBtmStatusStrip.Items["Status"].Text = "Success Saving PIL Configuration File"
        }
        Catch
        {
          $Response = Get-UserResponse -Title "Error Saving Config" -Message "There was an Error Saving the PIL Configuration" -ButtonMid OK -ButtonDefault OK -Icon ([System.Drawing.SystemIcons]::Error)
          $PILBtmStatusStrip.Items["Status"].Text = "Error Exporting CSV Report"
        }
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Saving PIL Configuration File"
      }
      Break
      #endregion Save PIL Configuration File
    }
    "ProcessItems"
    {
      #region Start List Processing
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Start Processing All List Items"
      $PILBtmStatusStrip.Refresh()
      
      If ($PILItemListListView.Items.Count -eq 0)
      {
        $Response = Get-UserResponse -Title "No List Items" -Message "There are no List Items to Proccess!" -ButtonMid OK -ButtonDefault OK -Icon ([System.Drawing.SystemIcons]::Warning)
      }
      Else
      {
        If ([String]::IsNullOrEmpty([MyRuntime]::ThreadConfig.ThreadScript))
        {
          $Response = Get-UserResponse -Title "No PIL Configureation" -Message "There is no PIL Script Configured!" -ButtonMid OK -ButtonDefault OK -Icon ([System.Drawing.SystemIcons]::Warning)
        }
        Else
        {
          # Disable Main Menu Iteme
          $PILTopMenuStrip.Items["AddItems"].Enabled = $False
          $PILTopMenuStrip.Items["Configure"].Enabled = $False
          $PILTopMenuStrip.Items["ProcessItems"].Enabled = $False
          $PILTopMenuStrip.Items["ListData"].Enabled = $False
          
          # Disable Right Click Menu
          $PILItemListContextMenuStrip.Enabled = $False
          
          # Disable ListView Sort
          $PILItemListListView.ListViewItemSorter.Enable = $False
          
          # Build RunSpace Pool
          $HashTable = @{ "ShowHeader" = $True; "ListItems" = @($PILItemListListView.Items) }
          $ScriptBlock = { [CmdletBinding()] Param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Start-ProcessingItems -RichTextBox $RichTextBox -HashTable $HashTable }
          $DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title "Initializing $([MyConfig]::ScriptName)" -ButtonMid "OK" -HashTable $HashTable -AutoClose -AutoCloseWait 1
          
          # Set Processing ToolStrip Menu Items
          $PILPlayProcButton.Enabled = $False
          $PILPlayPauseButton.Enabled = $True
          $PILPlayStopButton.Enabled = $True
          $PILPlayBarPanel.Visible = $True
          $PILForm.Refresh()
          
          $PILBtmStatusStrip.Items["Status"].Text = "Processing $($PILItemListListView.Items.Count) List Items"
          $PILBtmStatusStrip.Refresh()
          
          Monitor-RunspacePoolThreads
        }
      }
      Break
      #endregion Start List Processing
    }
    "ExportCSV"
    {
      #region Export CSV Report
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Export CSV Report"
      $PILBtmStatusStrip.Refresh()
      
      If ($PILItemListListView.Items.Count -eq 0)
      {
        $PILBtmStatusStrip.Items["Status"].Text = "No List Items to Export"
      }
      Else
      {
        # Save Export File
        $PILSaveFileDialog.FileName = ""
        $PILSaveFileDialog.Filter = "CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
        $PILSaveFileDialog.FilterIndex = 1
        $PILSaveFileDialog.Title = "Export PIL CSV Report"
        $PILSaveFileDialog.Tag = $Null
        $Response = $PILSaveFileDialog.ShowDialog()
        If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
        {
          Try
          {
            $TmpCount = ([MyRuntime]::CurrentColumns - 1)
            $StringBuilder = [System.Text.StringBuilder]::New()
            [Void]$StringBuilder.AppendLine(($PILItemListListView.Columns[0..$($TmpCount)] | Select-Object -ExpandProperty Text) -Join ",")
            $PILItemListListView.Items | ForEach-Object -Process { [Void]$StringBuilder.AppendLine("`"{0}`"" -f (($PSItem.SubItems[0..$($TmpCount)] | Select-Object -ExpandProperty Text) -join "`",`"")) }
            ConvertFrom-Csv -InputObject (($StringBuilder.ToString())) -Delimiter "," | Export-Csv -Path $PILSaveFileDialog.FileName -NoTypeInformation -Encoding ASCII
            $StringBuilder.Clear()
            
            # Save Current Directory
            $PILSaveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILSaveFileDialog.FileName)
            $PILBtmStatusStrip.Items["Status"].Text = "Success Exporting CSV Report"
          }
          Catch
          {
            $Response = Get-UserResponse -Title "Error Exporting" -Message "There was an Error Exporting the PIL Report Data" -ButtonMid OK -ButtonDefault OK -Icon ([System.Drawing.SystemIcons]::Error)
            $PILBtmStatusStrip.Items["Status"].Text = "Error Exporting CSV Report"
          }
        }
        Else
        {
          $PILBtmStatusStrip.Items["Status"].Text = "Canceled Exporting CSV Report"
        }
      }
      Break
      #endregion Export CSV Report
    }
    "DeleteList"
    {
      #region Clear Item List
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Clear Item List?"
      $PILBtmStatusStrip.Refresh()
      
      If ($PILItemListListView.Items.Count -eq 0)
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Previously Cleared Item List"
      }
      Else
      {
        $Response = Get-UserResponse -Title "Clear Item List?" -Message "Do you want to Clear the Item List?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
        If ($Response.Success)
        {
          # Clear Item List
          $PILItemListListView.Items.Clear()
          
          # Set Status Message
          $PILBtmStatusStrip.Items["Status"].Text = "Successfully Cleared Item List"
        }
        Else
        {
          # Set Status Message
          $PILBtmStatusStrip.Items["Status"].Text = "Canceled Clearing Item List"
        }
      }
      Break
      #endregion Clear Item List
    }
    "Sample"
    {
      #region Load Sample Configuration
      $PILBtmStatusStrip.Items["Status"].Text = "Loading Sample PIL Configuration"
      $PILBtmStatusStrip.Refresh()
      
      Switch ($Sender.Tag)
      {
        "SampleDemo"
        {
          $ConfigObject = $SampleDemo
          $ConfigName = "Sample - Demo"
          Break
        }
        "StarterConfig"
        {
          $ConfigObject = $StarterConfig
          $ConfigName = "Starter Config"
          Break
        }
        "GetWorkstationInfo"
        {
          $ConfigObject = $GetWorkstationInfo
          $ConfigName = "Get-WorkstationInfo"
          Break
        }
        "GetDomainComps"
        {
          $ConfigObject = $GetDomainComputer
          $ConfigName = "Get-DomainComputers"
          Break
        }
        "GetDomainUsers"
        {
          $ConfigObject = $GetDomainUser
          $ConfigName = "Get-DomainUsers"
          Break
        }
        "GraphAPIDevice"
        {
          $ConfigObject = $GraphAPIDevice
          $ConfigName = "Graph API Device"
          Break
        }
        "GraphAPIUser"
        {
          $ConfigObject = $GraphAPIUser
          $ConfigName = "Graph API User"
          Break
        }
      }
      $HashTable = @{"ShowHeader" = $True; "ConfigObject" = $ConfigObject; "ConfigName" = $ConfigName}
      $ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Load-PILConfigFIleJson -RichTextBox $RichTextBox -HashTable $HashTable }
      $DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title ($PILBtmStatusStrip.Items["Status"].Text) -ButtonMid "OK" -HashTable $HashTable
      If ($DialogResult.Success)
      {
        [MyRuntime]::ConfigName = $ConfigName
        
        If ($ConfigName -eq "Sample - Demo")
        {
          $PILItemListListView.BeginUpdate()
          $TmpSubItems = @("") * [MyRuntime]::CurrentColumns
          For ($I = 0; $I -lt 50; $I++)
          {
            $TmpListItem = [System.Windows.Forms.ListViewItem]::New(("Sample List Item {0:00}" -f $I), "StatusInfo16Icon")
            $TmpListItem.Name = $TmpItem
            $TmpListItem.Font = [MyConfig]::Font.Regular
            $TmpListItem.SubItems.AddRange($TmpSubItems)
            [Void]$PILItemListListView.Items.Add($TmpListItem)
          }
          $PILItemListListView.Columns[0].Width = -2
          $PILItemListListView.EndUpdate()
        }
        
        $PILBtmStatusStrip.Items["Status"].Text = "Success Loading Sample PIL Configuration"
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Errors Loading Sample PIL Configuration"
      }
      Break
      #endregion Load Sample Configuration
    }
    "Help"
    {
      #region Show Help
      $PILBtmStatusStrip.Items["Status"].Text = "Show Help"
      $PILBtmStatusStrip.Refresh()
      Show-MyWebReport -ReportURL ([MyConfig]::HelpURL)
      $PILBtmStatusStrip.Items["Status"].Text = "Success Help Shown"
      Break
      #endregion Show Help
    }
    "Exit"
    {
      #region Exit Utility
      If ([MyConfig]::Production)
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Exiting $([MyConfig]::ScriptName)"
        $PILBtmStatusStrip.Refresh()
        $PILForm.Close()
      }
      Else
      {
        # **** Testing - Exit to Nested Prompt ****
        Write-Host -Object "Line Num: $((Get-PSCallStack).ScriptLineNumber)"
        $Host.EnterNestedPrompt()
        # **** Testing - Exit to Nested Prompt ****
      }
      Break
      #endregion Exit Utility
    }
  }

  Write-Verbose -Message "Exit Click Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILTopMenuStripItemClick ********

$DropDownMenu = New-MenuItem -Menu $PILTopMenuStrip -Text "Add Items $([char]0x00BB)" -Name "AddItems" -Tag "AddItems" -DisplayStyle "ImageAndText" -ImageKey "AddItems16Icon" -TextImageRelation "ImageBeforeText" -PassThru
(New-MenuItem -Menu $DropDownMenu -Text "Add Item List" -Name "AddList" -Tag "AddList" -DisplayStyle "ImageAndText" -ImageKey "Add16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $DropDownMenu -Text "Import Item List" -Name "ImportList" -Tag "ImportList" -DisplayStyle "ImageAndText" -ImageKey "Import16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $DropDownMenu
(New-MenuItem -Menu $DropDownMenu -Text "Load PIL Data" -Name "LoadExport" -Tag "LoadExport" -DisplayStyle "ImageAndText" -ImageKey "LoadData16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})

$DropDownMenu = New-MenuItem -Menu $PILTopMenuStrip -Text "Configure $([char]0x00BB)" -Name "Configure" -Tag "Configure" -DisplayStyle "ImageAndText" -ImageKey "Config16Icon" -TextImageRelation "ImageBeforeText" -PassThru
$SubDropDownMenu = New-MenuItem -Menu $DropDownMenu -Text "Number of Columns" -Name "TotalColumns" -Tag "TotalColumns" -DisplayStyle "ImageAndText" -ImageKey "Calc16Icon" -TextImageRelation "ImageBeforeText" -PassThru
For ($I = [MyRuntime]::MinColumns; $I -le [MyRuntime]::MaxColumns; $I++)
{
  (New-MenuItem -Menu $SubDropDownMenu -Text ("{0:00} Total Columns" -f $I) -ToolTip "Set the Number of Item List Columns" -Name "TotalColumns" -Tag $I -DisplayStyle "ImageAndText" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
}
$SubDropDownMenu.DropDownItems[[MyRuntime]::CurrentColumns - [MyRuntime]::MinColumns].ImageKey = "Selected16Icon"
(New-MenuItem -Menu $DropDownMenu -Text "Set Column Names" -Name "ColumnNames" -Tag "ColumnNames" -DisplayStyle "ImageAndText" -ImageKey "Column16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})

(New-MenuItem -Menu $DropDownMenu -Text "Edit Configuration" -Name "ThreadConfig" -Tag "ThreadConfig" -DisplayStyle "ImageAndText" -ImageKey "Threads16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $DropDownMenu
(New-MenuItem -Menu $DropDownMenu -Text "Load Configuration" -Name "LoadConfig" -Tag "LoadConfig" -DisplayStyle "ImageAndText" -ImageKey "LoadConfig16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $DropDownMenu -Text "Save Configuration" -Name "SaveConfig" -Tag "SaveConfig" -DisplayStyle "ImageAndText" -ImageKey "SaveConfig16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $DropDownMenu

$SubDropDownMenu = New-MenuItem -Menu $DropDownMenu -Text "Sample PIL Configs" -Name "Examples" -Tag "Examples" -DisplayStyle "ImageAndText" -ImageKey "Demo16Icon" -PassThru
(New-MenuItem -Menu $SubDropDownMenu -Text "Get Workstation Info" -Name "Sample" -Tag "GetWorkstationInfo" -DisplayStyle "ImageAndText" -ImageKey "Demo16Icon" -PassThru).add_Click({ Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $SubDropDownMenu
(New-MenuItem -Menu $SubDropDownMenu -Text "Get Domain Computers" -Name "Sample" -Tag "GetDomainComps" -DisplayStyle "ImageAndText" -ImageKey "Demo16Icon" -PassThru).add_Click({ Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $SubDropDownMenu -Text "Get Domain Users" -Name "Sample" -Tag "GetDomainUsers" -DisplayStyle "ImageAndText" -ImageKey "Demo16Icon" -PassThru).add_Click({ Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $SubDropDownMenu
(New-MenuItem -Menu $SubDropDownMenu -Text "Graph API Devices" -Name "Sample" -Tag "GraphAPIDevice" -DisplayStyle "ImageAndText" -ImageKey "Demo16Icon" -PassThru).add_Click({ Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $SubDropDownMenu -Text "Graph API User" -Name "Sample" -Tag "GraphAPIUser" -DisplayStyle "ImageAndText" -ImageKey "Demo16Icon" -PassThru).add_Click({ Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $SubDropDownMenu
(New-MenuItem -Menu $SubDropDownMenu -Text "PIL Starter Config" -Name "Sample" -Tag "StarterConfig" -DisplayStyle "ImageAndText" -ImageKey "Demo16Icon" -PassThru).add_Click({ Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $SubDropDownMenu
(New-MenuItem -Menu $SubDropDownMenu -Text "PIL Demo Script" -Name "Sample" -Tag "SampleDemo" -DisplayStyle "ImageAndText" -ImageKey "Demo16Icon" -PassThru).add_Click({ Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})

New-MenuSeparator -Menu $PILTopMenuStrip
(New-MenuItem -Menu $PILTopMenuStrip -Text "Process Items" -Name "ProcessItems" -Tag "ProcessItems" -DisplayStyle "ImageAndText" -ImageKey "Process16Icon" -TextImageRelation "ImageBeforeText" -ClickOnCheck -Check -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})

$DropDownMenu = New-MenuItem -Menu $PILTopMenuStrip -Text "List Data $([char]0x00BB)" -Name "ListData" -Tag "ListData" -DisplayStyle "ImageAndText" -ImageKey "ListData16Icon" -TextImageRelation "ImageBeforeText" -PassThru
(New-MenuItem -Menu $DropDownMenu -Text "Export CSV Report" -Name "ExportCSV" -Tag "ExportCSV" -DisplayStyle "ImageAndText" -ImageKey "Export16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $DropDownMenu
(New-MenuItem -Menu $DropDownMenu -Text "Delete All Items" -Name "DeleteList" -Tag "DeleteLists" -DisplayStyle "ImageAndText" -ImageKey "Trash16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $PILTopMenuStrip

#(New-MenuItem -Menu $PILTopMenuStrip -Text "&Help" -Name "Help" -Tag "Help" -DisplayStyle "ImageAndText" -ImageKey "HelpIcon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $PILTopMenuStrip -Text "E&xit" -Name "Exit" -Tag "Exit" -DisplayStyle "ImageAndText" -ImageKey "ExitIcon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})

# ************************************************
# PILBtm StatusStrip
# ************************************************
#region $PILBtmStatusStrip = [System.Windows.Forms.StatusStrip]::New()
$PILBtmStatusStrip = [System.Windows.Forms.StatusStrip]::New()
$PILForm.Controls.Add($PILBtmStatusStrip)
$PILBtmStatusStrip.BackColor = [MyConfig]::Colors.Back
$PILBtmStatusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom
$PILBtmStatusStrip.Enabled = $True
$PILBtmStatusStrip.Font = [MyConfig]::Font.Regular
$PILBtmStatusStrip.ForeColor = [MyConfig]::Colors.Fore
$PILBtmStatusStrip.ImageList = $PILSmallImageList
$PILBtmStatusStrip.ImageScalingSize = [System.Drawing.Size]::New(16, 16)
$PILBtmStatusStrip.Name = "PILBtmStatusStrip"
$PILBtmStatusStrip.ShowItemToolTips = $False
#$PILBtmStatusStrip.TabIndex = 0
#$PILBtmStatusStrip.TabStop = $False
#$PILBtmStatusStrip.Tag = [System.Object]::New()
$PILBtmStatusStrip.TextDirection = [System.Windows.Forms.ToolStripTextDirection]::Horizontal
#endregion $PILBtmStatusStrip = [System.Windows.Forms.StatusStrip]::New()

New-MenuLabel -Menu $PILBtmStatusStrip -Text "Status" -Name "Status" -Tag "Status"

#endregion ******** Controls for PIL Form ********

#region ******** Start Form  ********
Try
{
  [System.Windows.Forms.Application]::Run($PILForm)
}
Catch
{
  Write-Warning -Message "Error Running $([MyConfig]::ScriptName) Form: $($_.Exception.Message)"
}
Finally
{
  Write-Verbose -Message "Disposing $([MyConfig]::ScriptName) Form Components"
  $PILOpenFileDialog.Dispose()
  $PILSaveFileDialog.Dispose()
  $PILFormComponents.Dispose()
  $PILForm.Dispose()
}
#endregion ******** Start Form  ********

if ([MyConfig]::Production)
{
  [System.Environment]::Exit(0)
}

