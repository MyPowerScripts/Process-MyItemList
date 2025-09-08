# ----------------------------------------------------------------------------------------------------------------------
#  Script: PIL
# ----------------------------------------------------------------------------------------------------------------------
<#
Change Log for PIL
------------------------------------------------------------------------------------------------
1.0.0.0 - Initial Version
------------------------------------------------------------------------------------------------
#>

#requires -version 5.0

Using namespace System.Windows.Forms
Using namespace System.Drawing
Using namespace System.Collections
Using namespace System.Collections.Generic
Using namespace System.Collections.Specialized

<#
  .SYNOPSIS
  .DESCRIPTION
  .PARAMETER <Parameter-Name>
  .EXAMPLE
  .NOTES
    My Script PIL Version 1.0 by kensw on 08/27/2025
    Created with "Form Code Generator" Version 7.0.0.2
#>
[CmdletBinding()]
Param (
  [ValidateRange(2, 24)]
  [uint16]$MaxColumns = 16,
  [String]$ConfigFile,
  [String]$ExportFile
)
Import-Module PowerShellGet


$ErrorActionPreference = "Stop"

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
  static [bool]$Production = $False

  static [String]$ScriptName = "Process-ItemList"
  static [Version]$ScriptVersion = [Version]::New("1.0.0.0")
  static [String]$ScriptAuthor = "Ken Sweet"

  # Script Configuration
  static [String]$ScriptRoot = ""
  static [String]$ConfigFile = ""
  static [PSCustomObject]$ConfigData = [PSCustomObject]@{ }

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

  Static [OrderedDictionary]$RequiredModules = [Ordered]@{
    "Az.Accounts" = "4.0.2"
    "Az.KeyVault" = "6.3.1"
    "Az.Automation" = "1.11.1"
    "Microsoft.Graph.Authentication" = "2.28.0"
  }

  # Azure Logon Information
  static [String]$TenantID = ""
  static [String]$SubscriptionID = ""
  static [Object]$AADLogonInfo = $Null
  static [Object]$AccessToken = $Null
  static [HashTable]$AuthToken = @{ }

  # Default Form Color Mode
  static [Bool]$DarkMode = ((Get-Itemproperty -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -ErrorAction "SilentlyContinue").AppsUseLightTheme -eq "0")

  # Form Auto Exit
  static [Int]$AutoExit = 0
  static [Int]$AutoExitMax = 60
  static [Int]$AutoExitTic = 60000

  # Administrative Rights
  static [Bool]$IsLocalAdmin = ([Security.Principal.WindowsPrincipal]::New([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  static [Bool]$IsPowerUser = ([Security.Principal.WindowsPrincipal]::New([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::PowerUser)

  # KPI Event Logging
  static [Bool]$KPILogExists = $False
  static [String]$KPILogName = "KPI Event Log"

  # Network / Internet
  static [__ComObject]$IsConnected = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]"{DCB00C01-570F-4A9B-8D69-199FDBA5723B}"))

  # Default Script Credentials
  static [String]$Domain = "Domain"
  static [String]$UserID = "UserID"
  static [String]$Password = "P@ssw0rd"

  # Default SMTP Configuration
  static [String]$SMTPServer = "smtp.mydomain.local"
  static [int]$SMTPPort = 25

  # Default MEMCM Configuration
  static [String]$MEMCMServer = "MyMEMCM.MyDomain.Local"
  static [String]$MEMCMSite = "XYZ"
  static [String]$MEMCMNamespace = "Root\SMS\Site_XYZ"

  # Help / Issues Uri's
  static [String]$HelpURL = "https://www.microsoft.com/"
  static [String]$BugURL = "https://www.amazon.com/"

  # CertKet for Cert Encryption
  static [String]$CertKey = ""

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
  [String[]]$ColumnNames
  [OrderedDictionary]$Modules = [OrderedDictionary]::New()
  [HashTable]$Functions = [HashTable]::New()
  [HashTable]$Variables = [HashTable]::New()
  [uint16]$ThreadCount = 8
  [String]$ThreadScript = $Null
  
  PILThreadConfig ([uint16]$MaxColumns)
  {
    $This.ColumnNames = [String[]]::New($MaxColumns)
    For ($I = 0; $I -lt $MaxColumns; $I++)
    {
      $This.ColumnNames[$I] = ("Column Name {0:00}" -f $I)
    }
  }
  
  [Void] SetColumnNames ([String[]]$ColumnNames)
  {
    $Max = $ColumnNames.Count
    For ($I = 0; $I -lt $Max; $I++)
    {
      $This.ColumnNames[$I] = $ColumnNames[$I]
    }
  }
  
  [OrderedDictionary] GetColumnNames ()
  {
    $TmpValue = [Ordered]@{ }
    
    $Max = $This.ColumnNames.GetUpperBound(0)
    For ($I = 0; $I -le $Max; $I++)
    {
      $TmpValue.Add(("Column Name {0:00}" -f $I), $This.ColumnNames[$I])
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
  # Max Number of Columns
  Static [Uint16]$MaxColumns = $MaxColumns
  
  # Thread Configuration
  Static [PILThreadConfig]$ThreadConfig = [PILThreadConfig]::New($MaxColumns)
  
  # Path to Module Install Locatiosn
  Static [String]$AUModules = "$($ENV:ProgramFiles)\WindowsPowerShell\Modules"
  Static [String]$CUModules = "$([Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments))\WindowsPowerShell\Modules"
  
  # List of Installed Modules
  Static [HashTable]$Modules = [HashTable]::New()
  
  # Loaded Functions
  Static [HashTable]$Functions = [HashTable]::New()
  
  Static [Void] UpdateTotalColumn ([Uint16]$MaxColumns)
  {
    [MyRuntime]::MaxColumns = $MaxColumns
    [MyRuntime]::ThreadConfig = [PILThreadConfig]::New($MaxColumns)
  }
}

#endregion ******** PIL Runtime  Values ********

#region ******** My Default Enumerations ********

#region ******** enum MyAnswer ********
[Flags()]
enum MyAnswer
{
  Unknown = 0
  No      = 1
  Yes     = 2
  Maybe   = 3
}
#endregion ******** enum MyAnswer ********

#region ******** enum MyDigit ********
enum MyDigit
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
enum MyBits
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

#region ******** enum PILColumns ********
Enum PILColumns
{
  Column00 = 0
  Column01 = 1
  Column02 = 2
  Column03 = 3
  Column04 = 4
  Column05 = 5
  Column06 = 6
  Column07 = 7
  Column08 = 8
  Column09 = 9
  Column10 = 10
  Column11 = 11
  Column12 = 12
  Column13 = 13
  Column14 = 14
  Column15 = 15
  Column16 = 16
  Column17 = 17
  Column18 = 18
  Column19 = 19
  Column20 = 20
  Column21 = 21
  Column22 = 22
  Column23 = 23
}
#endregion ******** enum PILColumns ********

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

#region ******** Sample Functions ********

#region function Verb-Noun
Function Verb-Noun ()
{
  <#
    .SYNOPSIS
      Function to do something specific
    .DESCRIPTION
      Function to do something specific
    .PARAMETER Value
      Value Command Line Parameter
    .EXAMPLE
      Verb-Noun -Value "String"
    .NOTES
      Original Function By %YourName%
      
      %Date% - Initial Release
  #>
  [CmdletBinding(DefaultParameterSetName = "ByValue")]
  Param (
    [parameter(Mandatory = $False, ParameterSetName = "ByValue")]
    [String[]]$Value = "Default Value"
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"
  
  # Loop and Proccess all Values
  ForEach ($Item In $Value)
  {
  }
  
  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Verb-Noun

#region function Verb-NounPiped
Function Verb-NounPiped()
{
  <#
    .SYNOPSIS
      Function to do something specific
    .DESCRIPTION
      Function to do something specific
    .PARAMETER Value
      Value Command Line Parameter
    .EXAMPLE
      Verb-NounPiped -Value "String"
    .EXAMPLE
      $Value | Verb-NounPiped
    .NOTES
      Original Function By %YourName%
      
      %Date% - Initial Release
  #>
  [CmdletBinding(DefaultParameterSetName = "ByValue")]
  Param (
    [parameter(Mandatory = $False, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True, ParameterSetName = "ByValue")]
    [String[]]$Value = "Default Value"
  )
  Begin
  {
    Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand) Begin Block"
    # This Code is Executed Once when the Function Begins
    
    Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand) Begin Block"
  }
  Process
  {
    Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand) Process Block"
    
    # Loop and Proccess all Values
    ForEach ($Item In $Value)
    {
    }
    
    Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand) Process Block"
  }
  End
  {
    Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand) End Block"
    # This Code is Executed Once whent he Function Ends
    
    Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand) End Block"
  }
}
#endregion function Verb-NounPiped

#endregion ******** Sample Functions ********

#region ******* Microsoft Forms Functions ********

#region function New-MyListItem
function New-MyListItem
{
  <#
    .SYNOPSIS
      Creates and adds a new list item to a ComboBox or ListBox control.
    .DESCRIPTION
      This function creates a new list item as a PSCustomObject with Text, Value, and Tag properties,
      and adds it to the Items collection of the specified ComboBox or ListBox control.
      Optionally, the new item can be returned via the PassThru switch.
    .PARAMETER Control
      The ComboBox or ListBox control to which the new item will be added. This parameter is mandatory.
    .PARAMETER Text
      The display text for the new list item. This parameter is mandatory.
    .PARAMETER Value
      The value associated with the new list item. This parameter is mandatory.
    .PARAMETER Tag
      An optional object to associate additional data with the new list item.
    .PARAMETER PassThru
      If specified, the function returns the newly created list item object instead of $null.
    .EXAMPLE
      New-MyListItem -Control $comboBox -Text "Option 1" -Value "1" -Tag "First Option"
      Adds a new item with text "Option 1", value "1", and tag "First Option" to the $comboBox control.
    .EXAMPLE
      $item = New-MyListItem -Control $listBox -Text "Item A" -Value "A" -PassThru
      Adds a new item to $listBox and returns the created item object.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [Object]$Control,
    [Parameter(Mandatory = $true)]
    [String]$Text,
    [Parameter(Mandatory = $true)]
    [String]$Value,
    [Object]$Tag,
    [switch]$PassThru
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  $item = [PSCustomObject]@{
    Text  = $Text
    Value = $Value
    Tag   = $Tag
  }

  if ($PassThru)
  {
    $Control.Items.Add($item)
    $item
  }
  else
  {
    [Void]$Control.Items.Add($item)
  }

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function New-MyListItem

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

#region function New-ListViewItem
function New-ListViewItem()
{
  <#
    .SYNOPSIS
      Creates and adds a new ListViewItem to a ListView control.
    .DESCRIPTION
      This function creates a new System.Windows.Forms.ListViewItem with the specified properties and adds it to the provided ListView control.
      It supports customization of text, name, subitems, tag, indentation, group, tooltip, checked state, font, colors, and image (by index or key).
      The new ListViewItem can optionally be returned via the PassThru switch.
    .PARAMETER ListView
      The ListView control to which the new ListViewItem will be added. This parameter is mandatory.
    .PARAMETER BackColor
      The background color of the ListViewItem. Defaults to [MyConfig]::Colors.TextBack.
    .PARAMETER ForeColor
      The foreground (text) color of the ListViewItem. Defaults to [MyConfig]::Colors.TextFore.
    .PARAMETER Font
      The font to use for the ListViewItem text. Defaults to [MyConfig]::Font.Regular.
    .PARAMETER Name
      The name of the new ListViewItem. If not specified, the Text value is used.
    .PARAMETER Text
      The display text for the new ListViewItem. This parameter is mandatory.
    .PARAMETER SubItems
      An array of strings to add as subitems to the ListViewItem.
    .PARAMETER Tag
      An object to associate additional data with the new ListViewItem.
    .PARAMETER IndentCount
      The number of indentation levels to apply to the ListViewItem. Only used with ImageIndex or ImageKey parameter sets.
    .PARAMETER ImageIndex
      The index of the image to display for the ListViewItem. Used when specifying images by index.
    .PARAMETER ImageKey
      The key of the image to display for the ListViewItem. Used when specifying images by key.
    .PARAMETER Group
      The ListViewGroup to which the new ListViewItem will be added.
    .PARAMETER ToolTip
      The tooltip text to display when the mouse hovers over the ListViewItem.
    .PARAMETER Checked
      If specified, sets the ListViewItem's Checked property to $true.
    .PARAMETER PassThru
      If specified, returns the newly created ListViewItem object.
    .EXAMPLE
      $NewItem = New-ListViewItem -ListView $listView -Text "Text" -Tag "Tag"
      Adds a new ListViewItem with text "Text" and tag "Tag" to $listView.
    .EXAMPLE
      $item = New-ListViewItem -ListView $listView -Text "Item1" -ImageIndex 2 -SubItems @("Sub1","Sub2") -PassThru
      Adds a new ListViewItem with an image at index 2 and subitems, and returns the created ListViewItem object.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param(
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ListView]$ListView,
    [System.Drawing.Color]$BackColor = [MyConfig]::Colors.TextBack,
    [System.Drawing.Color]$ForeColor = [MyConfig]::Colors.TextFore,
    [System.Drawing.Font]$Font = [MyConfig]::Font.Regular,
    [String]$Name,
    [parameter(Mandatory = $True)]
    [String]$Text,
    [String[]]$SubItems,
    [Object]$Tag,
    [parameter(Mandatory = $False, ParameterSetName = "Index")]
    [parameter(Mandatory = $False, ParameterSetName = "Key")]
    [Int]$IndentCount = 0,
    [parameter(Mandatory = $True, ParameterSetName = "Index")]
    [Int]$ImageIndex = -1,
    [parameter(Mandatory = $True, ParameterSetName = "Key")]
    [String]$ImageKey,
    [System.Windows.Forms.ListViewGroup]$Group,
    [String]$ToolTip,
    [Switch]$Checked,
    [switch]$PassThru
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  #region $TempListViewItem = [System.Windows.Forms.ListViewItem]
  if ($PSCmdlet.ParameterSetName -eq "Default")
  {
    $TempListViewItem = [System.Windows.Forms.ListViewItem]::New($Text, $Group)
  }
  else
  {
    if ($PSCmdlet.ParameterSetName -eq "Index")
    {
      $TempListViewItem = [System.Windows.Forms.ListViewItem]::New($Text, $ImageIndex, $Group)
    }
    else
    {
      $TempListViewItem = [System.Windows.Forms.ListViewItem]::New($Text, $ImageKey, $Group)
    }
    $TempListViewItem.IndentCount = $IndentCount
  }

  if ($PSBoundParameters.ContainsKey("Name"))
  {
    $TempListViewItem.Name = $Name
  }
  else
  {
    $TempListViewItem.Name = $Text
  }

  $TempListViewItem.Tag = $Tag
  $TempListViewItem.ToolTipText = $ToolTip
  $TempListViewItem.Checked = $Checked.IsPresent

  $TempListViewItem.BackColor = $BackColor
  $TempListViewItem.ForeColor = $ForeColor
  $TempListViewItem.Font = $Font
  if ($PSBoundParameters.ContainsKey("SubItems"))
  {
    $TempListViewItem.SubItems.AddRange($SubItems)
  }
  #endregion $TempListViewItem = [System.Windows.Forms.ListViewItem]

  [Void]$ListView.Items.Add($TempListViewItem)

  if ($PassThru.IsPresent)
  {
    $TempListViewItem
  }

  $TempListViewItem = $Null

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function New-ListViewItem

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

#region function New-ListViewGroup
function New-ListViewGroup()
{
  <#
    .SYNOPSIS
      Creates and adds a new ListViewGroup to a ListView control.
    .DESCRIPTION
      This function creates a new System.Windows.Forms.ListViewGroup with the specified properties and adds it to the provided ListView control.
      It supports customization of the group header, name, tag, and alignment. The new ListViewGroup can optionally be returned via the PassThru switch.
    .PARAMETER ListView
      The ListView control to which the new ListViewGroup will be added. This parameter is mandatory.
    .PARAMETER Header
      The display header text for the new ListViewGroup. This parameter is mandatory.
    .PARAMETER Name
      The name of the new ListViewGroup. If not specified, the Header value is used.
    .PARAMETER Tag
      An object to associate additional data with the new ListViewGroup.
    .PARAMETER Alignment
      The alignment of the group header text. Defaults to 'Left'.
    .PARAMETER PassThru
      If specified, returns the newly created ListViewGroup object.
    .EXAMPLE
      $NewItem = New-ListViewGroup -ListView $listView -Header "Header" -Tag "Tag"
      Adds a new ListViewGroup with header "Header" and tag "Tag" to $listView.
    .EXAMPLE
      $group = New-ListViewGroup -ListView $listView -Header "Group1" -Name "GroupOne" -Alignment Center -PassThru
      Adds a new ListViewGroup with header "Group1", name "GroupOne", centered alignment, and returns the created group object.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param(
    [parameter(Mandatory = $True)]
    [System.Windows.Forms.ListView]$ListView,
    [parameter(Mandatory = $True)]
    [String]$Header,
    [String]$Name,
    [Object]$Tag,
    [System.Windows.Forms.HorizontalAlignment]$Alignment = "Left",
    [switch]$PassThru
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  #region $TempListViewGroup = [System.Windows.Forms.ListViewGroup]
  $TempListViewGroup = [System.Windows.Forms.ListViewGroup]::New()
  [Void]$ListView.Groups.Add($TempListViewGroup)
  $TempListViewGroup.Tag = $Tag
  $TempListViewGroup.Header = $Header
  if ($PSBoundParameters.ContainsKey("Name"))
  {
    $TempListViewGroup.Name = $Name
  }
  else
  {
    $TempListViewGroup.Name = $Header
  }
  $TempListViewGroup.HeaderAlignment = $Alignment
  #endregion $TempListViewGroup = [System.Windows.Forms.ListViewGroup]

  if ($PassThru.IsPresent)
  {
    $TempListViewGroup
  }

  $TempListViewGroup = $Null

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function New-ListViewGroup

#region function Scale-MyForm
function Scale-MyForm()
{
  <#
    .SYNOPSIS
      Scales a Windows Forms control and its child controls by a specified factor.
    .DESCRIPTION
      This function recursively scales the size and font of a Windows Forms control (such as a Form, Panel, GroupBox, etc.) and all its child controls by the specified scale factor. 
      It is useful for DPI scaling or dynamically resizing UI elements to accommodate different display settings or user preferences.
      The function handles controls with child controls in the Controls collection, as well as controls with an Items collection (such as ListBox, ComboBox, etc.).
    .PARAMETER Control
      The Windows Forms control to scale. This can be a Form or any control derived from System.Windows.Forms.Control. 
      If not specified, defaults to the global variable $FCGForm.
    .PARAMETER Scale
      The scaling factor to apply to the control and its children. 
      For example, a value of 1.25 increases size by 25%, while 0.8 reduces size by 20%. 
      The default value is 1 (no scaling).
    .EXAMPLE
      Scale-MyForm -Control $Form -Scale 1.5
      Scales the specified form and all its child controls by 150%.
    .EXAMPLE
      Scale-MyForm -Scale 0.9
      Scales the default form ($FCGForm) and all its child controls by 90%.
    .NOTES
      Original Function By Ken Sweet
      Recursively scales all controls, including those with an Items collection.
  #>
  [CmdletBinding()]
  param (
    [Object]$Control = $FCGForm,
    [Single]$Scale = 1
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  if ($Control -is [System.Windows.Forms.Form])
  {
    $Control.Scale($Scale)
  }

  $Control.Font = [System.Drawing.Font]::New($Control.Font.FontFamily, ($Control.Font.Size * $Scale), $Control.Font.Style)

  if ([String]::IsNullOrEmpty($Control.PSObject.Properties.Match("Items")))
  {
    if ($Control.Controls.Count)
    {
      foreach ($ChildControl in $Control.Controls)
      {
        Scale-MyForm -Control $ChildControl -Scale $Scale
      }
    }
  }
  else
  {
    foreach ($Item in $Control.Items)
    {
      Scale-MyForm -Control $Item -Scale $Scale
    }
  }

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Scale-MyForm

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

#endregion ******* Microsoft Forms Functions ********

#region ******* Encrypt / Encode Data Functions ********

#region function Encode-MyData
function Encode-MyData()
{
  <#
    .SYNOPSIS
      Encodes or decodes data to and from Base64 format.
    .DESCRIPTION
      This function encodes a string to Base64 with optional line length, or decodes a Base64 string back to its original form. 
      It supports output as a string or as an array of characters.
    .PARAMETER Data
      The string data to encode or decode. When encoding, this is the plain text to convert to Base64. When decoding, this is the Base64 string to convert back.
    .PARAMETER LineLength
      The maximum length of each line in the encoded Base64 output. Only used when encoding. Default is 160.
    .PARAMETER Decode
      Switch to indicate that the operation should decode the input Base64 string instead of encoding.
    .PARAMETER AsString
      When decoding, outputs the result as a string instead of an array of characters.
    .EXAMPLE
      Encode-MyData -Data "MySecret" 
      Encodes the string "MySecret" to Base64.
    .EXAMPLE
      Encode-MyData -Data $Base64String -Decode
      Decodes the Base64 string back to its original value.
    .EXAMPLE
      Encode-MyData -Data $Base64String -Decode -AsString
      Decodes the Base64 string and returns the result as a string.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Encode")]
  param (
    [parameter(Mandatory = $True)]
    [String]$Data,
    [parameter(Mandatory = $False, ParameterSetName = "Encode")]
    [Int]$LineLength = 160,
    [parameter(Mandatory = $True, ParameterSetName = "Decode")]
    [Switch]$Decode,
    [parameter(Mandatory = $False, ParameterSetName = "Decode")]
    [Switch]$AsString
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  $MemoryStream = [System.IO.MemoryStream]::New()

  if ($PSCmdlet.ParameterSetName -eq "Encode")
  {
    $StreamWriter = [System.IO.StreamWriter]::New($MemoryStream, [System.Text.Encoding]::UTF8)
    $StreamWriter.Write($Data)
    $StreamWriter.Close()

    $Encoded = [System.Text.StringBuilder]::New()
    ForEach ($Line in @([System.Convert]::ToBase64String($MemoryStream.ToArray()) -split "(?<=\G.{$LineLength})(?=.)"))
    {
      [Void]$Encoded.AppendLine($Line)
    }
    $Encoded.ToString()
    $MemoryStream.Close()
  }
  else
  {
    $CompressedData = [System.Convert]::FromBase64String($Data)
    $MemoryStream.Write($CompressedData, 0, $CompressedData.Length)
    [Void]$MemoryStream.Seek(0, 0)
    $StreamReader = [System.IO.StreamReader]::New($MemoryStream, [System.Text.Encoding]::UTF8)

    if ($AsString.IsPresent)
    {
      $StreamReader.ReadToEnd()
    }
    else
    {
      $ArrayList = [System.Collections.ArrayList]::New()
      $Buffer = [System.Char[]]::New(4096)
      While ($StreamReader.EndOfStream -eq $False)
      {
        $Bytes = $StreamReader.Read($Buffer, 0, 4096)
        if ($Bytes)
        {
          $ArrayList.AddRange($Buffer[0 .. ($Bytes - 1)])
        }
      }
      $ArrayList
      $ArrayList.Clear()
    }
    $StreamReader.Close()
    $MemoryStream.Close()
  }

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Encode-MyData

#region Function Protect-MySensitiveData
Function Protect-MySensitiveData
{
  <#
    .SYNOPSIS
      Encrypts or decrypts text string data using AES encryption.
    .DESCRIPTION
      This function encrypts or decrypts a plain text string using a passphrase and optional salt, hash algorithm, cipher mode, and padding mode. 
      It supports both encryption and decryption operations based on the -Decrypt switch.
    .PARAMETER String
      The plain text string to encrypt, or the encrypted Base64 string to decrypt.
    .PARAMETER PassPhrase
      The passphrase used to derive the encryption key.
    .PARAMETER Salt
      The salt value used in key derivation. Must be at least 8 characters. Default is "Pepper".
    .PARAMETER HashAlgorithm
      The hash algorithm used for key derivation. Default is SHA256.
    .PARAMETER CipherMode
      The cipher mode for AES encryption. Default is CBC.
    .PARAMETER PaddingMode
      The padding mode for AES encryption. Default is PKCS7.
    .PARAMETER Decrypt
      Switch to indicate that the function should decrypt the input string instead of encrypting.
    .EXAMPLE
      $EncryptedData = Protect-MySensitiveData -String "SecretText" -PassPhrase "MyPass" -Salt "MySalt"
      Encrypts the string "SecretText" using the specified passphrase and salt.
    .EXAMPLE
      $DecryptedData = Protect-MySensitiveData -String $EncryptedData -PassPhrase "MyPass" -Salt "MySalt" -Decrypt
      Decrypts the previously encrypted string using the same passphrase and salt.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [String]$String,
    [parameter(Mandatory = $True)]
    [String]$PassPhrase = "PassPhrase",
    [parameter(Mandatory = $False)]
    [String]$Salt = "Pepper",
    [System.Security.Cryptography.HashAlgorithmName]$HashAlgorithm = [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.CipherMode]$CipherMode = [System.Security.Cryptography.CipherMode]::CBC,
    [System.Security.Cryptography.PaddingMode]$PaddingMode = [System.Security.Cryptography.PaddingMode]::PKCS7,
    [Switch]$Decrypt
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  # Create Cryptography AES Object
  $Aes = [System.Security.Cryptography.Aes]::Create()
  $Aes.Mode = $CipherMode
  $Aes.Padding = $PaddingMode
  # Salt Needs to be at least 8 Characters
  $SaltBytes = [System.Text.Encoding]::UTF8.GetBytes($Salt.PadRight(8, "*"))
  $Aes.Key = [System.Security.Cryptography.Rfc2898DeriveBytes]::New($PassPhrase, $SaltBytes, 8, $HashAlgorithm).GetBytes($Aes.Key.Length)

  if ($Decrypt.IsPresent)
  {
    # Decrypt Encrypted Data
    $DecodeBytes = [System.Convert]::FromBase64String($String)
    $Aes.IV = $DecodeBytes[0..15]
    $Decryptor = $Aes.CreateDecryptor()
    [System.Text.Encoding]::UTF8.GetString(($Decryptor.TransformFinalBlock($DecodeBytes, 16, ($DecodeBytes.Length - 16))))
  }
  else
  {
    # Encrypt String Data
    $EncodeBytes = [System.Text.Encoding]::UTF8.GetBytes($String)
    $Encryptor = $Aes.CreateEncryptor()
    $EncryptedBytes = [System.Collections.ArrayList]::New($Aes.IV)
    $EncryptedBytes.AddRange($Encryptor.TransformFinalBlock($EncodeBytes, 0, $EncodeBytes.Length))
    [System.Convert]::ToBase64String($EncryptedBytes)
    $EncryptedBytes.Clear()
  }

  $Aes.Dispose()

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion Function Protect-MySensitiveData

#region function Protect-WithCert
Function Protect-WithCert ()
{
  <#
    .SYNOPSIS
      Encrypts or decrypts text data using certificate information as key material.
    .DESCRIPTION
      This function encrypts or decrypts a text string using properties from a specified certificate as the passphrase and salt for AES encryption. 
      You can optionally use a salt derived from the certificate's validity period, and choose to use universal (UTC) or local time for salt generation. 
      Decryption uses the same parameters as encryption.
    .PARAMETER CertKey
      The thumbprint or subject name of the certificate in the LocalMachine\Root store to use for encryption/decryption.
    .PARAMETER TextString
      The text string to encrypt or decrypt.
    .PARAMETER Salt
      An integer (0-3) specifying which salt format to use, based on the certificate's NotBefore/NotAfter properties. Only used if specified.
    .PARAMETER Local
      Switch to use Local time for salt generation instead of UTC time. Only relevant if -Salt is specified.
    .PARAMETER Decrypt
      Switch to indicate that the function should decrypt the input string instead of encrypting.
    .EXAMPLE
      # Encrypt with local salt
      $EncryptedText = Protect-WithCert -CertKey $CertKey -Salt 0 -TextString $TextString
    .EXAMPLE
      # Encrypt with universal salt
      $EncryptedText = Protect-WithCert -CertKey $CertKey -Salt 1 -Universal -TextString $TextString
    .EXAMPLE
      # Encrypt with no salt (uses certificate subject as salt)
      $EncryptedText = Protect-WithCert -CertKey $CertKey -TextString $TextString
    .EXAMPLE
      # Decrypt previously encrypted text
      $DecryptedText = Protect-WithCert -CertKey $CertKey -Salt 0 -TextString $EncryptedText -Decrypt
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "NoSalt")]
  Param (
    [parameter(Mandatory = $True)]
    [String]$CertKey,
    [parameter(Mandatory = $True)]
    [String]$TextString,
    [parameter(Mandatory = $True, ParameterSetName = "WithSalt")]
    [ValidateRange(0, 3)]
    [Int]$Salt,
    [parameter(Mandatory = $False, ParameterSetName = "WithSalt")]
    [Switch]$Local,
    [Switch]$Decrypt
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  $Cert = Get-ChildItem -Path "Cert:\LocalMachine\Root\$($CertKey)"
  If ($PSCmdlet.ParameterSetName -eq "WithSalt")
  {
    If ($Universal.IsPresent)
    {
      $TmpNotBefore = $Cert.NotBefore.ToUniversalTime()
      $TmpNotAfter = $Cert.NotAfter.ToUniversalTime()
    }
    Else
    {
      $TmpNotBefore = $Cert.NotBefore
      $TmpNotAfter = $Cert.NotAfter
    }
    $SaltInit = @($TmpNotBefore.ToString("yyyyMMddhhmmss"), $TmpNotBefore.ToString("hhmmssyyyyMMdd"), $TmpNotAfter.ToString("yyyyMMddhhmmss"), $TmpNotAfter.ToString("hhmmssyyyyMMdd"))[$Salt]
  }
  Else
  {
    $SaltInit = $Cert.Subject
  }
  Protect-MySensitiveData -PassPhrase ($Cert.SerialNumber) -Salt $SaltInit -String $TextString -Decrypt:($Decrypt.IsPresent)

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Protect-WithCert

#region function Encrypt-MyTextString
function Encrypt-MyTextString()
{
  <#
    .SYNOPSIS
      Encrypts or decrypts a text string using Windows Data Protection API (DPAPI).
    .DESCRIPTION
      This function encrypts a plain text string or decrypts an encrypted Base64 string using the Windows Data Protection API (DPAPI). 
      You can specify the protection scope (CurrentUser or LocalMachine) and optionally provide an additional encryption key (entropy) for extra security.
      When encrypting, the function returns a Base64-encoded string. When decrypting, it returns the original plain text.
    .PARAMETER TextString
      The text string to encrypt or decrypt. When encrypting, this is the plain text to secure. When decrypting, this is the Base64-encoded string to restore.
    .PARAMETER ProtectionScope
      Specifies the scope of protection. 
      'CurrentUser' restricts decryption to the current user (default).
      'LocalMachine' allows any user on the machine to decrypt.
    .PARAMETER EncryptKey
      An optional string used as additional entropy (extra encryption key) for added security. If not specified, no extra entropy is used.
    .PARAMETER Decrypt
      Switch to indicate that the function should decrypt the input string instead of encrypting.
    .EXAMPLE
      Encrypt-MyTextString -TextString "MyPassword"
      Encrypts the string "MyPassword" for the current user.
    .EXAMPLE
      Encrypt-MyTextString -TextString $EncryptedString -Decrypt
      Decrypts the previously encrypted string for the current user.
    .EXAMPLE
      Encrypt-MyTextString -TextString "MyPassword" -ProtectionScope LocalMachine -EncryptKey "ExtraSecret"
      Encrypts the string "MyPassword" for any user on the machine, using "ExtraSecret" as additional entropy.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [String]$TextString,
    [ValidateSet("LocalMachine", "CurrentUser")]
    [String]$ProtectionScope = "CurrentUser",
    [String]$EncryptKey = $Null,
    [Switch]$Decrypt
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  if ([String]::IsNullOrEmpty(([Management.Automation.PSTypeName]::New("System.Security.Cryptography.ProtectedData")).Type))
  {
    [Void][System.Reflection.Assembly]::LoadWithPartialName("System.Security")
  }

  if ($PSBoundParameters.ContainsKey("EncryptKey"))
  {
    $OptionalEntropy = [System.Text.Encoding]::ASCII.GetBytes($EncryptKey)
  }
  else
  {
    $OptionalEntropy = $Null
  }

  if ($Decrypt.IsPresent)
  {
    $EncryptedData = [System.Convert]::FromBase64String($TextString)
    $DecryptedData = [System.Security.Cryptography.ProtectedData]::Unprotect($EncryptedData, $OptionalEntropy, ([System.Security.Cryptography.DataProtectionScope]$ProtectionScope))
    [System.Text.Encoding]::ASCII.GetString($DecryptedData)
  }
  else
  {
    $TempData = [System.Text.Encoding]::ASCII.GetBytes($TextString)
    $EncryptedData = [System.Security.Cryptography.ProtectedData]::Protect($TempData, $OptionalEntropy, ([System.Security.Cryptography.DataProtectionScope]$ProtectionScope))
    [System.Convert]::ToBase64String($EncryptedData)
  }

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Encrypt-MyTextString

#region function Decode-MySecureString
function Decode-MySecureString ()
{
  <#
    .SYNOPSIS
      Decodes a SecureString to plain text.
    .DESCRIPTION
      This function converts a System.Security.SecureString object to its plain text string representation. 
      It is useful for retrieving the original value from a SecureString, such as when you need to use the password or sensitive data in plain text form.
    .PARAMETER SecureString
      The SecureString object to decode. This should be a System.Security.SecureString instance containing the sensitive data you want to convert to plain text.
    .EXAMPLE
      $secure = Read-Host "Enter secret" -AsSecureString
      Decode-MySecureString -SecureString $secure
      Decodes the entered SecureString and outputs the plain text value.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [System.Security.SecureString]$SecureString
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString))

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Decode-MySecureString

#region function Convert-MyImageToBase64
function Convert-MyImageToBase64()
{
  <#
    .SYNOPSIS
      Converts an image or icon to a Base64-encoded text block for embedding in scripts.
    .DESCRIPTION
      This function converts an image file (such as .ico, .gif, .jpg, etc.) or a System.Drawing.Icon object into a Base64-encoded string, formatted for easy inclusion in PowerShell scripts. 
      The output includes region markers and variable assignment for direct use. You can specify the output variable name and the maximum line length for the Base64 string.
    .PARAMETER ScriptName
      The name of the script or variable prefix to use in the generated code for referencing the image list. Optional.
    .PARAMETER Icon
      A System.Drawing.Icon object to convert to Base64. Use this parameter set to encode an icon object directly.
    .PARAMETER Path
      The file path to the image to convert. Supported formats include .ico, .gif, .jpg, and others supported by System.Drawing.Image.
    .PARAMETER Name
      The variable name to assign the Base64 string to in the generated code. This should be a valid PowerShell variable name.
    .PARAMETER LineSize
      The maximum number of characters per line in the Base64 output. Default is 160.
    .EXAMPLE
      Convert-MyImageToBase64 -Path "C:\Icons\myicon.ico" -Name "MyIcon"
      Converts the specified .ico file to a Base64 string and outputs PowerShell code assigning it to $MyIcon.
    .EXAMPLE
      $icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\notepad.exe")
      Convert-MyImageToBase64 -Icon $icon -Name "NotepadIcon"
      Converts the provided Icon object to Base64 and outputs PowerShell code assigning it to $NotepadIcon.
    .NOTES
      Original Function By Ken Sweet. Useful for embedding images or icons in PowerShell GUIs or scripts.
  #>
  [CmdletBinding(DefaultParameterSetName = "File")]
  Param (
    [String]$ScriptName,
    [parameter(Mandatory = $True, ParameterSetName = "Icon")]
    [System.Drawing.Icon]$Icon,
    [parameter(Mandatory = $True, ParameterSetName = "File")]
    [String]$Path,
    [parameter(Mandatory = $True)]
    [String]$Name,
    [int]$LineSize = 160
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"

  $StringBuilder = [System.Text.StringBuilder]::New()

  $ImageName = $Name.Replace(".", "").Replace("-", "").Replace(" ", "").Replace("ico", "Icon")
  [Void]$StringBuilder.AppendLine("#region ******** `$$($ImageName) ********")
  [Void]$StringBuilder.AppendLine("`$$($ImageName) = @`"")
  $MemoryStream = [System.IO.MemoryStream]::New()
  if ($PSCmdlet.ParameterSetName -eq "File")
  {
    Switch ([System.IO.Path]::GetExtension($Path))
    {
      ".ico"
      {
        $Image = [System.Drawing.Icon]::New($Path)
        $Image.Save($MemoryStream)
        Break
      }
      ".gif"
      {
        $Image = [System.Drawing.Image]::FromFile($Path)
        $Image.Save($MemoryStream, [System.Drawing.Imaging.ImageFormat]::Gif)
        Break
      }
      Default
      {
        $Image = [System.Drawing.Image]::FromFile($Path)
        $Image.Save($MemoryStream, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        Break
      }
    }
  }
  else
  {
    $Image = $Icon
    $Image.Save($MemoryStream)
  }
  ForEach ($Line in @([System.Convert]::ToBase64String($MemoryStream.ToArray()) -split "(?<=\G.{$LineSize})(?=.)"))
  {
    [Void]$StringBuilder.AppendLine($Line)
  }
  $MemoryStream.Close()
  [Void]$StringBuilder.AppendLine("`"@")
  [Void]$StringBuilder.AppendLine("#endregion ******** `$$($ImageName) ********")
  if (([System.IO.Path]::GetExtension($path) -eq ".ico") -or ($PSCmdlet.ParameterSetName -eq "Icon"))
  {
    #[Void]$StringBuilder.AppendLine("#`$Form.Icon = [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String(`$$($ImageName))))")
    [Void]$StringBuilder.AppendLine("`$$($ScriptName)ImageList.Images.Add(`"$($ImageName)`", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String(`$$($ImageName)))))")
  }
  else
  {
    [Void]$StringBuilder.AppendLine("#`$PictureBox.Image = [System.Drawing.Image]::FromStream([System.IO.MemoryStream]::New([System.Convert]::FromBase64String(`$$($ImageName))))")
  }
  $StringBuilder.ToString()

  $Image = $Null
  $MemoryStream = $Null
  $StringBuilder = $Null

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Convert-MyImageToBase64

#region function Compress-MyData
Function Compress-MyData()
{
  <#
    .SYNOPSIS
      Compress / Decompress String Data
    .DESCRIPTION
      Compress / Decompress String Data
    .PARAMETER Data
      Path to Text File to Compress
    .PARAMETER DataName
      Name to put in Data Region Comments
    .PARAMETER Path
      Data to Compress / Decompress
    .PARAMETER LineLength
      Max Line Length
    .PARAMETER Decompress
      Decompress the Data String
    .PARAMETER AsString
      Return as a String
    .EXAMPLE
      Compress-MyData -Data "String"
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "CompressText")]
  Param (
    [parameter(Mandatory = $False, ParameterSetName = "CompressText")]
    [parameter(Mandatory = $True, ParameterSetName = "Decompress")]
    [String]$Data,
    [parameter(Mandatory = $False, ParameterSetName = "CompressText")]
    [String]$DataName = "CompressedText",
    [parameter(Mandatory = $True, ParameterSetName = "CompressFile")]
    [String]$Path,
    [parameter(Mandatory = $False, ParameterSetName = "CompressFile")]
    [parameter(Mandatory = $False, ParameterSetName = "CompressText")]
    [Int]$LineLength = 160,
    [parameter(Mandatory = $True, ParameterSetName = "CompressFile")]
    [parameter(Mandatory = $False, ParameterSetName = "CompressText")]
    [Switch]$Encode,
    [parameter(Mandatory = $False, ParameterSetName = "Decompress")]
    [Switch]$Decompress,
    [parameter(Mandatory = $False, ParameterSetName = "Decompress")]
    [Switch]$AsString
  )
  Write-Verbose -Message "Enter Function Compress-MyData"
  
  If ($PSCmdLet.ParameterSetName -eq "Decompress")
  {
    $CompressedData = [System.Convert]::FromBase64String($Data)
    $MemoryStream = [System.IO.MemoryStream]::New()
    $MemoryStream.Write($CompressedData, 0, $CompressedData.Length)
    [Void]$MemoryStream.Seek(0, 0)
    $GZipStream = [System.IO.Compression.GZipStream]::New($MemoryStream, [System.IO.Compression.CompressionMode]::Decompress)
    $StreamReader = [System.IO.StreamReader]::New($GZipStream, [System.Text.Encoding]::UTF8)
    If ($AsString.IsPresent)
    {
      $StreamReader.ReadToEnd()
    }
    Else
    {
      $ArrayList = [System.Collections.ArrayList]::New()
      $Buffer = [System.Char[]]::New(4096)
      While (-not $StreamReader.EndOfStream)
      {
        $Bytes = $StreamReader.Read($Buffer, 0, 4096)
        If ($Bytes -gt 0)
        {
          $ArrayList.AddRange($Buffer[0..($Bytes - 1)])
        }
      }
      $ArrayList
      $ArrayList.Clear()
    }
    # Close Reader
    $StreamReader.Close()
  }
  Else
  {
    If ($PSCmdlet.ParameterSetName -eq "CompressFile")
    {
      $Data = Get-Content -Path $Path -Raw
      $DataName = ([System.IO.Path]::GetFileName($Path) -replace "[\.\-\s]", "")
    }
    $MemoryStream = [System.IO.MemoryStream]::New()
    $GZipStream = [System.IO.Compression.GZipStream]::New($MemoryStream, [System.IO.Compression.CompressionMode]::Compress)
    $StreamWriter = [System.IO.StreamWriter]::New($GZipStream, [System.Text.Encoding]::UTF8)
    $StreamWriter.Write($Data)
    # Close Writer
    $StreamWriter.Close()
    $StringBuilder = [System.Text.StringBuilder]::New()
    If ($Encode.IsPresent)
    {
      [Void]$StringBuilder.AppendLine("#region $($DataName) Data")
      [Void]$StringBuilder.AppendLine("`$$($DataName) = @`"")
    }
    ForEach ($Line In @([System.Convert]::ToBase64String($MemoryStream.ToArray()) -split "(?<=\G.{$LineLength})(?=.)"))
    {
      [Void]$StringBuilder.AppendLine($Line)
    }
    If ($Encode.IsPresent)
    {
      [Void]$StringBuilder.AppendLine("`"@")
      [Void]$StringBuilder.AppendLine("#endregion $($DataName) Data")
    }
    # Return Encryped Data
    $StringBuilder.ToString()
    [void]$StringBuilder.Clear()
  }
  
  # Close Streams
  $GZipStream.Close()
  $MemoryStream.Close()
  
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
  
  Write-Verbose -Message "Exit Function Compress-MyData"
}
#endregion function Compress-MyData

#endregion ******* Encrypt / Encode Data Functions ********

#region function Write-MyLogFile
function Write-MyLogFile()
{
  <#
    .SYNOPSIS
      Writes a log entry to a specified log file with customizable severity, formatting, and output options.
    .DESCRIPTION
      The Write-MyLogFile function writes log messages to a file, with support for log rotation, severity levels, colored host output, and customizable log folder and file names. 
      It is designed for flexible logging in scripts and automation tasks.
    .PARAMETER LogFolder
      Specifies the folder where the log file will be stored. Defaults to the script name if not specified.
    .PARAMETER LogName
      Specifies the name of the log file. Defaults to the script name with a .log extension.
    .PARAMETER SystemLog
      Switch to use the Windows system log folder for storing the log file.
    .PARAMETER Severity
      Specifies the severity of the log entry. Valid values are Text, Info, Good, Warning, and Error. Default is Text.
    .PARAMETER Message
      The message to log. This parameter is mandatory.
    .PARAMETER Component
      Specifies the component or source of the log entry. Defaults to the script name.
    .PARAMETER Context
      Additional context information for the log entry.
    .PARAMETER Thread
      The thread or process ID associated with the log entry. Defaults to the current process ID.
    .PARAMETER MaxSize
      The maximum size (in bytes) of the log file before it is rotated. Default is 52428800 (50 MB).
    .PARAMETER OutHost
      Switch to also write the log message to the host (console) with color.
    .PARAMETER ColorText
      The color used for Text severity messages in the host output. Default is Gray.
    .PARAMETER ColorInfo
      The color used for Info severity messages in the host output. Default is DarkCyan.
    .PARAMETER ColorGood
      The color used for Good severity messages in the host output. Default is DarkGreen.
    .PARAMETER ColorWarn
      The color used for Warning severity messages in the host output. Default is DarkYellow.
    .PARAMETER ColorError
      The color used for Error severity messages in the host output. Default is DarkRed.
    .EXAMPLE
      Write-MyLogFile -LogFolder "MyLogFolder" -Message "This is My Info Log File Message"
    .EXAMPLE
      Write-MyLogFile -LogFolder "MyLogFolder" -Severity "Info" -Message "This is My Info Log File Message"
    .EXAMPLE
      Write-MyLogFile -LogFolder "MyLogFolder" -Severity "Warning" -Message "This is My Warning Log File Message"
    .EXAMPLE
      Write-MyLogFile -LogFolder "MyLogFolder" -Severity "Error" -Message "This is My Error Log File Message"
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [parameter(Mandatory = $False, ParameterSetName = "LogFolder")]
    [String]$LogFolder = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName),
    [String]$LogName = "$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName)).log",
    [parameter(Mandatory = $False, ParameterSetName = "SystemLog")]
    [Switch]$SystemLog,
    [ValidateSet("Text", "Info", "Good", "Warning", "Error")]
    [String]$Severity = "Text",
    [parameter(Mandatory = $True)]
    [String]$Message,
    [String]$Component = "",
    [String]$Context = "",
    [Int]$Thread = $PID,
    [ValidateRange(0, 16777216)]
    [Int]$MaxSize = 52428800,
    [Switch]$OutHost,
    [ConsoleColor]$ColorText = [ConsoleColor]::Gray,
    [ConsoleColor]$ColorInfo = [ConsoleColor]::DarkCyan,
    [ConsoleColor]$ColorGood = [ConsoleColor]::DarkGreen,
    [ConsoleColor]$ColorWarn = [ConsoleColor]::DarkYellow,
    [ConsoleColor]$ColorError = [ConsoleColor]::DarkRed
  )
  Write-Verbose -Message "Enter Function Write-MyLogFile"

  switch ($PSCmdlet.ParameterSetName)
  {
    "LogFolder"
    {
      $LogPath = $LogFolder
      break
    }
    "SystemLog"
    {
      $LogPath = "$($ENV:SystemRoot)\Logs\$($LogFolder)"
      break
    }
    Default
    {
      $LogPath = "$($PSScriptRoot)\Logs"
      break
    }
  }

  if (-not [System.IO.Directory]::Exists($LogPath))
  {
    [Void][System.IO.Directory]::CreateDirectory($LogPath)
  }
  $TempFile = "$($LogPath)\$LogName"

  switch ($Severity)
  {
    "Text"
    {
      $TempSeverity = 1
      $HostColor = $ColorText
      break
    }
    "Info"
    {
      $TempSeverity = 1
      $HostColor = $ColorInfo
      break
    }
    "Good"
    {
      $TempSeverity = 1
      $HostColor = $ColorGood
      break
    }
    "Warning"
    {
      $TempSeverity = 2
      $HostColor = $ColorWarn
      break
    }
    "Error"
    {
      $TempSeverity = 3
      $HostColor = $ColorError
      break
    }
  }

  $TempDate = [DateTime]::Now

  if (-not $PSBoundParameters.ContainsKey("Component"))
  {
    $TempSource = [System.IO.Path]::GetFileName($MyInvocation.ScriptName)
    $Component = [System.IO.Path]::GetFileNameWithoutExtension($TempSource)
  }

  if ([System.IO.File]::Exists($TempFile) -and $MaxSize -gt 0)
  {
    if (([System.IO.FileInfo]$TempFile).Length -gt $MaxSize)
    {
      $TempBackup = [System.IO.Path]::ChangeExtension($TempFile, "lo_")
      if ([System.IO.File]::Exists($TempBackup))
      {
        Remove-Item -Force -Path $TempBackup
      }
      Rename-Item -Force -Path $TempFile -NewName ([System.IO.Path]::GetFileName($TempBackup))
    }
  }

  if ($OutHost.IsPresent)
  {
    Write-Host -Object "$($TempDate.ToString("yy-MM-dd HH:mm:ss")) - $($Message)" -ForegroundColor $HostColor
  }

  Add-Content -Encoding Ascii -Path $TempFile -Value ("<![LOG[{0}]LOG]!><time=`"{1}`" date=`"{2}`" component=`"{3}`" context=`"{4}`" type=`"{5}`" thread=`"{6}`" file=`"{7}`">" -f $Message, $($TempDate.ToString("HH:mm:ss.fff+000")), $($TempDate.ToString("MM-dd-yyyy")), $Component, $Context, $TempSeverity, $Thread, $TempSource)

  Write-Verbose -Message "Exit Function Write-MyLogFile"
}
#endregion function Write-MyLogFile

#region ******* Generic / General Functions ********

#region function Invoke-MyPause
function Invoke-MyPause
{
  <#
    .SYNOPSIS
      Pauses script execution for a specified number of Milliseconds or until a condition is met.
    .DESCRIPTION
      This function pauses the script for the specified number of Milliseconds. Optionally, a ScriptBlock can be provided to determine if the pause should continue. 
      The function processes Windows Forms events during the pause.
    .PARAMETER Milliseconds
      The number of Milliseconds to pause the script. Default is 60.
    .PARAMETER ScriptBlock
      A ScriptBlock that returns $True to continue pausing or $False to stop. Default is { $True }.
    .EXAMPLE
      Invoke-MyPause -Milliseconds 30
      Pauses the script for 30 Milliseconds.
    .EXAMPLE
      Invoke-MyPause -Milliseconds 10 -ScriptBlock { $global:ContinuePause }
      Pauses the script for up to 10 Milliseconds or until $global:ContinuePause is $False.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [UInt16]$Milliseconds = 60,
    [ScriptBlock]$ScriptBlock = { $True }
  )
  Write-Verbose -Message "Enter Function Invoke-MyPause"

  $TmpPause = [System.Diagnostics.Stopwatch]::StartNew()
  do
  {
    [System.Threading.Thread]::Sleep(10)
    $WaitCheck = $ScriptBlock.Invoke()
    [System.Windows.Forms.Application]::DoEvents()
  }
  while (($TmpPause.Elapsed.TotalMilliseconds -lt $Milliseconds) -and $WaitCheck)
  $TmpPause.Stop()

  Write-Verbose -Message "Exit Function Invoke-MyPause"
}
#endregion function Invoke-MyPause

#region function Set-MyClipboard
function Set-MyClipboard()
{
  <#
    .SYNOPSIS
      Copies object data to the clipboard in HTML and CSV format.
    .DESCRIPTION
      This function copies the specified object data to the clipboard, formatting it as both HTML and CSV. 
      The HTML output includes customizable styles for the table title, property headers, and row colors. 
      The function is useful for exporting tabular data from PowerShell scripts for use in other applications.
    .PARAMETER Items
      The array of objects to copy to the clipboard. Each object should contain the properties specified in the Properties parameter.
    .PARAMETER Title
      The title displayed at the top of the HTML table. Default is "My Copied Data from PowerShell".
    .PARAMETER TitleFore
      The foreground (text) color for the table title. Default is "Black".
    .PARAMETER TitleBack
      The background color for the table title. Default is "LightSteelBlue".
    .PARAMETER Properties
      The list of property names to include as columns in the table. This parameter is mandatory.
    .PARAMETER PropertyFore
      The foreground (text) color for the property header row. Default is "Black".
    .PARAMETER PropertyBack
      The background color for the property header row. Default is "PowderBlue".
    .PARAMETER RowFore
      The foreground (text) color for all data rows. Default is "Black".
    .PARAMETER RowEvenBack
      The background color for even-numbered data rows. Default is "White".
    .PARAMETER RowOddBack
      The background color for odd-numbered data rows. Default is "Gainsboro".
    .EXAMPLE
      Set-MyClipboard -Items $Items -Title "This is My Title" -Properties "Property1", "Property2", "Property3"
      Copies the specified properties of $Items to the clipboard with a custom title.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Office")]
  param (
    [parameter(Mandatory = $True, ValueFromPipeline = $True)]
    [Object[]]$Items,
    [String]$Title = "My Copied Data from PowerShell",
    [String]$TitleFore = "Black",
    [String]$TitleBack = "LightSteelBlue",
    [parameter(Mandatory = $True)]
    [String[]]$Properties,
    [String]$PropertyFore = "Black",
    [String]$PropertyBack = "PowderBlue",
    [String]$RowFore = "Black",
    [String]$RowEvenBack = "White",
    [String]$RowOddBack = "Gainsboro"
  )
  begin
  {
    Write-Verbose -Message "Enter Function Set-MyClipboard Begin Block"

    # Init StringBuilding
    $HTMLStringBuilder = [System.Text.StringBuilder]::New()

    # Start HTML ClipBaord Data
    [Void]$HTMLStringBuilder.Append("Version:1.0`r`nStartHTML:000START`r`nEndHTML:00000END`r`nStartFragment:00FSTART`r`nEndFragment:0000FEND`r`n")
    [Void]$HTMLStringBuilder.Replace("000START", ("{0:X8}" -f $HTMLStringBuilder.Length))
    [Void]$HTMLStringBuilder.Append("<html><head><title>My Copied Data</title></head><body><!--StartFragment-->")
    [Void]$HTMLStringBuilder.Replace("00FSTART", ("{0:X8}" -f $HTMLStringBuilder.Length))

    # Table Style
    [Void]$HTMLStringBuilder.Append("<style>`r`n.Title{border: 1px solid black; border-collapse: collapse; font-weight: bold; text-align: center; color: $($TitleFore); background: $($TitleBack);}`r`n.Property{border: 1px solid black; border-collapse: collapse; font-weight: bold; text-align: center; color: $($PropertyFore); background: $($PropertyBack);}`r`n.Row0 {border: 1px solid black; border-collapse: collapse;color: $($RowFore); background: $($RowEvenBack);}`r`n.Row1 {border: 1px solid black; border-collapse: collapse; color: $($RowFore); background: $($RowOddBack);}`r`n</style>")

    # Start Build Table / Set Title
    [Void]$HTMLStringBuilder.Append("<table><tr><th class=Title aligh=center colspan=$($Properties.Count)>&nbsp;$($Title)&nbsp;</th></tr>")

    # Add Table Column / Property Names
    [Void]$HTMLStringBuilder.Append("<tr>$(($Properties | ForEach-Object -Process { "<td class=Property aligh=center>&nbsp;$($PSItem)&nbsp;</td>" }) -join '')</tr>")

    # Start Row Count
    $TmpRowCount = 0

    $TmpItemList = [System.Collections.ArrayList]::New()

    Write-Verbose -Message "Exit Function Set-MyClipboard Begin Block"
  }
  process
  {
    Write-Verbose -Message "Enter Function Set-MyClipboard Process Block"

    foreach ($Item in $Items)
    {
      [Void]$HTMLStringBuilder.Append("<tr>$(((($Properties | ForEach-Object -Process { $Item.($PSItem) }) | ForEach-Object -Process { "<td class=Row$($TmpRowCount)>&nbsp;$($PSItem)&nbsp;</td>" }) -join ''))</tr>")
      [Void]$TmpItemList.Add(($Item | Select-Object -Property $Properties))
      $TmpRowCount = ($TmpRowCount + 1) % 2
    }

    Write-Verbose -Message "Exit Function Set-MyClipboard Process Block"
  }
  end
  {
    Write-Verbose -Message "Enter Function Set-MyClipboard End Block"

    # Close HTML Table
    [Void]$HTMLStringBuilder.Append("</table><br><br>")

    # Set End Clipboard Values
    [Void]$HTMLStringBuilder.Replace("0000FEND", ("{0:X8}" -f $HTMLStringBuilder.Length))
    [Void]$HTMLStringBuilder.Append("<!--EndFragment--></body></html>")
    [Void]$HTMLStringBuilder.Replace("00000END", ("{0:X8}" -f $HTMLStringBuilder.Length))

    [System.Windows.Forms.Clipboard]::Clear()
    $DataObject = [System.Windows.Forms.DataObject]::New("Text", ($TmpItemList | Select-Object -Property $Properties | ConvertTo-Csv -NoTypeInformation | Out-String))
    $DataObject.SetData("HTML Format", $HTMLStringBuilder.ToString())
    [System.Windows.Forms.Clipboard]::SetDataObject($DataObject)

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Function Set-MyClipboard End Block"
  }
}
#endregion function Set-MyClipboard

#region function Send-MyEMail
function Send-MyEMail()
{
  <#
    .SYNOPSIS
      Sends an E-mail message using SMTP.
    .DESCRIPTION
      This function sends an E-mail message using the specified SMTP server and port. You can specify recipients, sender, subject, body, message file, HTML formatting, CC, BCC, attachments, and priority.
    .PARAMETER SMTPServer
      The SMTP server to use for sending the E-mail. Default is [MyConfig]::SMTPServer.
    .PARAMETER SMTPPort
      The port number to use for the SMTP server. Default is [MyConfig]::SMTPPort.
    .PARAMETER To
      One or more recipient E-mail addresses. Mandatory.
    .PARAMETER From
      The sender's E-mail address. Mandatory.
    .PARAMETER Subject
      The subject of the E-mail message. Mandatory.
    .PARAMETER Body
      The body text of the E-mail message. If a file path is provided, the contents of the file will be used as the body.
    .PARAMETER IsHTML
      Indicates whether the body of the E-mail is formatted as HTML.
    .PARAMETER CC
      One or more E-mail addresses to send a carbon copy (CC) of the message.
    .PARAMETER BCC
      One or more E-mail addresses to send a blind carbon copy (BCC) of the message.
    .PARAMETER Attachment
      One or more attachments to include with the E-mail message.
    .PARAMETER Priority
      The priority of the E-mail message. Valid values are "Low", "Normal", or "High". Default is "Normal".
    .EXAMPLE
      Send-MyEMail -To "user@example.com" -From "me@example.com" -Subject "Test" -Body "Hello World" -SMTPServer "smtp.example.com" -SMTPPort 25
      Sends a simple E-mail message.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [String]$SMTPServer = [MyConfig]::SMTPServer,
    [Int]$SMTPPort = [MyConfig]::SMTPPort,
    [parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True, HelpMessage = "Enter To")]
    [System.Net.Mail.MailAddress[]]$To,
    [parameter(Mandatory = $True, HelpMessage = "Enter From")]
    [System.Net.Mail.MailAddress]$From,
    [parameter(Mandatory = $True, HelpMessage = "Enter Subject")]
    [String]$Subject,
    [parameter(Mandatory = $True, HelpMessage = "Enter Message Text")]
    [String]$Body,
    [Switch]$IsHTML,
    [System.Net.Mail.MailAddress[]]$CC,
    [System.Net.Mail.MailAddress[]]$BCC,
    [System.Net.Mail.Attachment[]]$Attachment,
    [ValidateSet("Low", "Normal", "High")]
    [System.Net.Mail.MailPriority]$Priority = "Normal"
  )
  begin
  {
    Write-Verbose -Message "Enter Function Send-MyEMail Begin"

    $MyMessage = [System.Net.Mail.MailMessage]::New()
    $MyMessage.From = $From
    $MyMessage.Subject = $Subject
    $MyMessage.IsBodyHtml = $IsHTML
    $MyMessage.Priority = $Priority

    if ($PSBoundParameters.ContainsKey("CC"))
    {
      foreach ($SendCC in $CC)
      {
        $MyMessage.CC.Add($SendCC)
      }
    }

    if ($PSBoundParameters.ContainsKey("BCC"))
    {
      foreach ($SendBCC in $BCC)
      {
        $MyMessage.BCC.Add($SendBCC)
      }
    }

    if ([System.IO.File]::Exists($Body))
    {
      $MyMessage.Body = $([System.IO.File]::ReadAllText($Body))
    }
    else
    {
      $MyMessage.Body = $Body
    }

    if ($PSBoundParameters.ContainsKey("Attachment"))
    {
      foreach ($AttachedFile in $Attachment)
      {
        $MyMessage.Attachments.Add($AttachedFile)
      }
    }

    Write-Verbose -Message "Exit Function Send-MyEMail Begin"
  }
  process
  {
    Write-Verbose -Message "Enter Function Send-MyEMail Process"

    $MyMessage.To.Clear()
    foreach ($SendTo in $To)
    {
      $MyMessage.To.Add($SendTo)
    }

    $SMTPClient = [System.Net.Mail.SmtpClient]::New($SMTPServer, $SMTPPort)
    $SMTPClient.Send($MyMessage)

    Write-Verbose -Message "Exit Function Send-MyEMail Process"
  }
  end
  {
    Write-Verbose -Message "Enter Function Send-MyEMail End"
    Write-Verbose -Message "Exit Function Send-MyEMail End"
  }
}
#endregion function Send-MyEMail

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

#region class MyConCommand
class MyConCommand
{
  [Int]$ExitCode
  [String]$OutputTxt
  [String]$ErrorMsg

  MyConCommand ([Int]$ExitCode, [String]$OutputTxt, [String]$ErrorMsg)
  {
    $This.ExitCode = $ExitCode
    $This.OutputTxt = $OutputTxt
    $This.ErrorMsg = $ErrorMsg
  }
}
#endregion class MyConCommand

#region function Invoke-MyConCommand
function Invoke-MyConCommand ()
{
  <#
    .SYNOPSIS
      Invokes a console command and returns the exit code, output, and error message.
    .DESCRIPTION
      This function executes a specified console command with optional parameters, captures the exit code, standard output, and standard error, and returns them in a custom object. 
      Useful for running external processes and retrieving their results in PowerShell.
    .PARAMETER Command
      The full path to the executable or command to run. This parameter is mandatory.
    .PARAMETER Parameters
      Optional command line parameters to pass to the executable. Default is $Null.
    .EXAMPLE
      Invoke-MyConCommand -Command "C:\Windows\System32\cmd.exe" -Parameters "/c Exit 1"
      Runs cmd.exe with the specified parameters and returns the exit code and output.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [String]$Command,
    [String]$Parameters = $Null
  )
  Write-Verbose -Message "Enter Function Invoke-MyConCommand"

  if ([System.IO.File]::Exists($Command))
  {
    $PSI = [System.Diagnostics.ProcessStartInfo]::New($Command, $Parameters)
    $PSI.UseShellExecute = $False
    $PSI.RedirectStandardError = $True
    $PSI.RedirectStandardOutput = $True
    try
    {
      $Out = [System.Diagnostics.Process]::Start($PSI)
      $Out.WaitForExit()
      [MyConCommand]::New($Out.ExitCode, $Out.StandardOutput.ReadToEnd(), $Out.StandardError.ReadToEnd())
    }
    catch
    {
      [MyConCommand]::New(-2, $Null, $Error[0].Message)
    }
  }
  else
  {
    [MyConCommand]::New(-1, $Null, "Command was not Found")
  }

  Write-Verbose -Message "Exit Function Invoke-MyConCommand"
}
#endregion function Invoke-MyConCommand

#region function Test-MyClassLoaded
function Test-MyClassLoaded()
{
  <#
    .SYNOPSIS
      Test if Custom Class is Loaded
    .DESCRIPTION
      Test if Custom Class is Loaded
    .PARAMETER Name
      Name of Custom Class
    .EXAMPLE
      $IsLoaded = Test-MyClassLoaded -Name "CustomClass"
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [parameter(Mandatory = $True, ParameterSetName = "Default")]
    [String]$Name
  )
  Write-Verbose -Message "Enter Function Test-MyClassLoaded"

  (-not [String]::IsNullOrEmpty(([Management.Automation.PSTypeName]::New($Name)).Type))

  Write-Verbose -Message "Exit Function Test-MyClassLoaded"
}
#endregion function Test-MyClassLoaded

#region function New-MyComObject
function New-MyComObject()
{
  <#
    .SYNOPSIS
      Creates Local and Remote COMObjects.
    .DESCRIPTION
      This function creates a COM object either locally or on a remote computer using the specified ProgID. 
      It is useful for automating tasks that require COM automation, such as interacting with Office applications or other COM-enabled software.
    .PARAMETER ComputerName
      The name of the computer on which to create the COM object. Defaults to the local computer.
    .PARAMETER COMObject
      The ProgID of the COM object to create. This parameter is mandatory.
    .EXAMPLE
      New-MyComObject -COMObject "Excel.Application"
      Creates an Excel COM object on the local computer.
    .EXAMPLE
      New-MyComObject -ComputerName "RemotePC" -COMObject "Excel.Application"
      Creates an Excel COM object on the remote computer "RemotePC".
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [String]$ComputerName = [System.Environment]::MachineName,
    [parameter(Mandatory = $True)]
    [String]$COMObject
  )
  Write-Verbose -Message "Enter Function New-MyComObject"

  [Activator]::CreateInstance([Type]::GetTypeFromProgID($COMObject, $ComputerName))

  Write-Verbose -Message "Exit Function New-MyComObject"
}
#endregion function New-MyComObject

#region function ConvertTo-MyIconImage
function ConvertTo-MyIconImage()
{
  <#
    .SYNOPSIS
      Converts a Base64-encoded string to an Icon or Image object.
    .DESCRIPTION
      This function takes a Base64-encoded string representing an image or icon and converts it back to a .NET Icon or Image object. 
      Use the -Image switch to specify that the output should be an Image object; otherwise, an Icon object is returned.
    .PARAMETER EncodedImage
      The Base64-encoded string representing the image or icon to convert. This parameter is mandatory.
    .PARAMETER Image
      If specified, the function returns a System.Drawing.Image object. If not specified, a System.Drawing.Icon object is returned.
    .EXAMPLE
      $NewItem = ConvertTo-MyIconImage -EncodedImage $EncodedImage
      Converts the Base64 string to an Icon object.
    .EXAMPLE
      $NewItem = ConvertTo-MyIconImage -EncodedImage $EncodedImage -Image
      Converts the Base64 string to an Image object.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $True)]
    [String]$EncodedImage,
    [Switch]$Image
  )
  Write-Verbose -Message "Enter Function ConvertTo-MyIconImage"

  if ($Image.IsPresent)
  {
    [System.Drawing.Image]::FromStream([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($EncodedImage)))
  }
  else
  {
    [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($EncodedImage)))
  }

  Write-Verbose -Message "Exit Function ConvertTo-MyIconImage"
}
#endregion function ConvertTo-MyIconImage

#region function Send-MyTextMessage
function Send-MyTextMessage ()
{
  <#
    .SYNOPSIS
      Sends a text message to a remote or local computer or IP address using UDP.
    .DESCRIPTION
      This function sends a text message to a specified computer name or IP address using UDP protocol. 
      You can specify the target by computer name or IP address, set the message content, and choose the port. 
      The function is useful for simple network notifications or inter-process communication.
    .PARAMETER ComputerName
      The name of the computer to send the message to. Defaults to the local computer. Used only if IPAddress is not specified.
    .PARAMETER IPAddress
      The IP address to send the message to. Defaults to "127.0.0.1". Use "255.255.255.255" for broadcast.
    .PARAMETER Message
      The text message to send. Defaults to "This is My Message".
    .PARAMETER Port
      The UDP port to use for sending the message. Defaults to 2500.
    .EXAMPLE
      Send-MyTextMessage -Message "Hello World" -IPAddress "192.168.1.100" -Port 2500
      Sends "Hello World" to IP address 192.168.1.100 on port 2500.
    .EXAMPLE
      Send-MyTextMessage -ComputerName "RemotePC" -Message "Test Notification"
      Sends "Test Notification" to the computer named RemotePC.
    .NOTES
      Original function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "IPAddress")]
  param (
    [parameter(Mandatory = $False, ParameterSetName = "ComputerName")]
    [String]$ComputerName = [System.Environment]::MachineName,
    [parameter(Mandatory = $False, ParameterSetName = "IPAddress")]
    [System.Net.IPAddress]$IPAddress = "127.0.0.1",
    [parameter(Mandatory = $False)]
    [String]$Message = "This is My Message",
    [int]$Port = 2500
  )
  Write-Verbose -Message "Enter function Send-MyTextMessage"

  if ($PSCmdlet.ParameterSetName -eq "IPAddress")
  {
    $RemoteClient = [System.Net.IPEndPoint]::New($IPAddress, $Port)
  }
  else
  {
    $RemoteClient = [System.Net.IPEndPoint]::New((([System.Net.Dns]::GetHostByName($ComputerName)).AddressList[0]), $Port)
  }
  $MessageBytes = [Text.Encoding]::ASCII.GetBytes("$($Message)")
  $UDPClient = [System.Net.Sockets.UdpClient]::New()
  $UDPClient.Send($MessageBytes, $MessageBytes.Length, $RemoteClient)
  $UDPClient.Close()
  $UDPClient.Dispose()

  Write-Verbose -Message "Exit function Send-MyTextMessage"
}
#endregion function Send-MyTextMessage

#region function Listen-MyTextMessage
function Listen-MyTextMessage ()
{
  <#
    .SYNOPSIS
      Listens for text messages sent via UDP from remote or local computers.
    .DESCRIPTION
      This function listens for incoming UDP text messages on a specified port and IP address or computer name. 
      It displays received messages and the sender's address. The listener runs until a message with the content "Exit" is received.
    .PARAMETER ComputerName
      The name of the computer to listen for messages from. Defaults to the local computer. Used only if IPAddress is not specified.
    .PARAMETER IPAddress
      The IP address to listen on. Defaults to "127.0.0.1". Use "0.0.0.0" to listen on all interfaces.
    .PARAMETER Port
      The UDP port to listen on. Defaults to 2500.
    .EXAMPLE
      Listen-MyTextMessage
      Listens for UDP messages on 127.0.0.1:2500.
    .EXAMPLE
      Listen-MyTextMessage -IPAddress "0.0.0.0" -Port 3000
      Listens for UDP messages on all interfaces at port 3000.
    .EXAMPLE
      Listen-MyTextMessage -ComputerName "RemotePC"
      Listens for UDP messages from the computer named RemotePC.
    .NOTES
      Original function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "IPAddress")]
  param (
    [parameter(Mandatory = $False, ParameterSetName = "ComputerName")]
    [String]$ComputerName = [System.Environment]::MachineName,
    [parameter(Mandatory = $False, ParameterSetName = "IPAddress")]
    [System.Net.IPAddress]$IPAddress = "127.0.0.1",
    [int]$Port = 2500
  )
  Write-Verbose -Message "Enter function Listen-MyTextMessage"

  if ($PSCmdlet.ParameterSetName -eq "IPAddress")
  {
    $RemoteClient = [System.Net.IPEndPoint]::New($IPAddress, $Port)
  }
  else
  {
    $RemoteClient = [System.Net.IPEndPoint]::New((([System.Net.Dns]::GetHostByName($ComputerName)).AddressList[0]), $Port)
  }
  $UDPClient = [System.Net.Sockets.UdpClient]::New($Port)
  do
  {
    $TempRemoteClient = $RemoteClient
    $Message = $UDPClient.Receive([ref]$TempRemoteClient)
    $DecodedMessage = [Text.Encoding]::ASCII.GetString($Message)
    Write-Host -Object "Message From: $($TempRemoteClient.Address) - $($DecodedMessage)"
  } while ($True -and ($DecodedMessage -ne "Exit"))

  Write-Verbose -Message "Exit function Listen-MyTextMessage"
}
#endregion function Listen-MyTextMessage

#region class MyWorkstationInfo
class MyWorkstationInfo
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
    if ($DomainMember)
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
    return $This
  }

  [TimeSpan] GetRunTime ()
  {
    return ($This.EndTime - $This.StartTime)
  }
}
#endregion class MyWorkstationInfo

#region function Get-MyWorkstationInfo
function Get-MyWorkstationInfo()
{
  <#
    .SYNOPSIS
      Verify Remote Workstation is the Correct One
    .DESCRIPTION
      Verify Remote Workstation is the Correct One
    .PARAMETER ComputerName
      Name of the Computer to Verify
    .PARAMETER Credential
      Credentials to use when connecting to the Remote Computer
    .PARAMETER Serial
      Return Serial Number
    .PARAMETER Mobile
      Check if System is Desktop / Laptop
    .INPUTS
    .OUTPUTS
    .EXAMPLE
      Get-MyWorkstationInfo -ComputerName "MyWorkstation"
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $False, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
    [String[]]$ComputerName = [System.Environment]::MachineName,
    [PSCredential]$Credential,
    [Switch]$Serial,
    [Switch]$Mobile
  )
  begin
  {
    Write-Verbose -Message "Enter Function Get-MyWorkstationInfo"

    # Default Common Get-WmiObject Options
    if ($PSBoundParameters.ContainsKey("Credential"))
    {
      $Params = @{
        "ComputerName" = $Null
        "Credential"   = $Credential
      }
    }
    else
    {
      $Params = @{
        "ComputerName" = $Null
      }
    }
  }
  process
  {
    Write-Verbose -Message "Enter Function Get-MyWorkstationInfo - Process"

    foreach ($Computer in $ComputerName)
    {
      # Start Setting Return Values as they are Found
      $VerifyObject = [MyWorkstationInfo]::New($Computer)

      # Validate ComputerName
      if ($Computer -match "^(([a-zA-Z]|[a-zA-Z][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z]|[A-Za-z][A-Za-z0-9\-]*[A-Za-z0-9])$")
      {
        try
        {
          # Get IP Address from DNS, you want to do all remote checks using IP rather than ComputerName.  If you connect to a computer using the wrong name Get-WmiObject will fail and using the IP Address will not
          $IPAddresses = @([System.Net.Dns]::GetHostAddresses($Computer) | Where-Object -FilterScript { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } | Select-Object -ExpandProperty IPAddressToString)
          :FoundMyWork foreach ($IPAddress in $IPAddresses)
          {
            if ([System.Net.NetworkInformation.Ping]::New().Send($IPAddress).Status -eq [System.Net.NetworkInformation.IPStatus]::Success)
            {
              # Set Default Parms
              $Params.ComputerName = $IPAddress

              # Get ComputerSystem
              [Void]($MyCompData = Get-WmiObject @Params -Class Win32_ComputerSystem)
              $VerifyObject.AddComputerSystem($Computer, $IPAddress, ($MyCompData.Name), ($MyCompData.PartOfDomain), ($MyCompData.Domain), ($MyCompData.Manufacturer), ($MyCompData.Model), ($MyCompData.UserName), ($MyCompData.TotalPhysicalMemory))
              $MyCompData.Dispose()

              # Verify Remote Computer is the Connect Computer, No need to get any more information
              if ($VerifyObject.Found)
              {
                # Start Secondary Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
                [Void]($MyOSData = Get-WmiObject @Params -Class Win32_OperatingSystem)
                $VerifyObject.AddOperatingSystem(($MyOSData.ProductType), ($MyOSData.Caption), ($MyOSData.CSDVersion), ($MyOSData.BuildNumber), ($MyOSData.Version), ($MyOSData.OSArchitecture), ([System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.LocalDateTime)), ([System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.InstallDate)), ([System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.LastBootUpTime)))
                $MyOSData.Dispose()

                # Optional SerialNumber Job
                if ($Serial.IsPresent)
                {
                  # Start Optional Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
                  [Void]($MyBIOSData = Get-WmiObject @Params -Class Win32_Bios)
                  $VerifyObject.AddSerialNumber($MyBIOSData.SerialNumber)
                  $MyBIOSData.Dispose()
                }

                # Optional Mobile / ChassisType Job
                if ($Mobile.IsPresent)
                {
                  # Start Optional Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
                  [Void]($MyChassisData = Get-WmiObject @Params -Class Win32_SystemEnclosure)
                  $VerifyObject.AddIsMobile($MyChassisData.ChassisTypes)
                  $MyChassisData.Dispose()
                }
              }
              else
              {
                $VerifyObject.UpdateStatus("Wrong Workstation Name")
              }
              # Beak out of Loop, Verify was a Success no need to try other IP Address if any
              break FoundMyWork
            }
          }
        }
        catch
        {
          # Workstation Not in DNS
          $VerifyObject.UpdateStatus("Workstation Not in DNS")
        }
      }
      else
      {
        $VerifyObject.UpdateStatus("Invalid Computer Name")
      }

      # Set End Time and Return Results
      $VerifyObject.SetEndTime()
    }
    Write-Verbose -Message "Exit Function Get-MyWorkstationInfo - Process"
  }
  end
  {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Verbose -Message "Exit Function Get-MyWorkstationInfo"
  }
}
#endregion function Get-MyWorkstationInfo

#region function Get-MyNetAdapterConStatus
function Get-MyNetAdapterConStatus ()
{
  <#
    .SYNOPSIS
      Gets the connection status of wired and wireless network adapters on a specified computer.
    .DESCRIPTION
      This function checks the network adapters on the specified computer and determines if there are any active wired or wireless connections. 
      It uses WMI queries to identify the physical medium type and connection status of each adapter.
    .PARAMETER ComputerName
      The name of the computer to query. Defaults to the local computer.
    .PARAMETER Credential
      The credentials to use when connecting to the remote computer. Defaults to an empty credential.
    .EXAMPLE
      Get-MyNetAdapterConStatus -ComputerName "PC01"
      Returns the wired and wireless connection status for computer "PC01".
    .EXAMPLE
      Get-MyNetAdapterConStatus -ComputerName "PC01" -Credential (Get-Credential)
      Returns the connection status using the specified credentials.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [String]$ComputerName = [System.Environment]::MachineName,
    [PSCredential]$Credential = [PSCredential]::Empty
  )
  Write-Verbose -Message "Enter Function Get-MyNetAdapterConStatus"

  $PhysicalMediumTypeList = @(Get-WmiObject -ComputerName $ComputerName -Credential $Credential -Namespace "Root\WMI" -Query "Select InstanceName, NdisPhysicalMediumType From MSNdis_PhysicalMediumType Where Active = 1" | Select-Object -Property InstanceName, NdisPhysicalMediumType)
  $NetworkAdapters = @(Get-WmiObject -ComputerName $ComputerName -Credential $Credential -Namespace "Root\CimV2" -Query "Select Name from Win32_NetworkAdapter Where NetConnectionStatus = 2" | Select-Object -ExpandProperty Name)
  [PSCustomObject][ordered]@{
    "Wired"    = (@($PhysicalMediumTypeList | Where-Object -FilterScript { ($PSItem.NdisPhysicalMediumType -eq 0) -and ($PSItem.InstanceName -in $NetworkAdapters) }).Count -gt 0)
    "Wireless" = (@($PhysicalMediumTypeList | Where-Object -FilterScript { ($PSItem.NdisPhysicalMediumType -eq 9) -and ($PSItem.InstanceName -in $NetworkAdapters) }).Count -gt 0)
  }

  Write-Verbose -Message "Exit Function Get-MyNetAdapterConStatus"
}
#endregion function Get-MyNetAdapterConStatus

#endregion ******* Generic / General Functions ********

#region ******* Registry / Environement Variable Functions ********

#region function Reset-MyRegKeyOwner
function Reset-MyRegKeyOwner ()
{
  <#
    .SYNOPSIS
      Take Ownership of a Registry Key and Reset Access Rules.
    .DESCRIPTION
      This function takes ownership of a specified registry key and optionally resets its access rules. It can operate recursively and supports changing ownership to either the Administrators or Users group.
    .PARAMETER Hive
      Specifies the registry hive to operate on. Defaults to LocalMachine.
    .PARAMETER Key
      The path of the registry key to take ownership of. This parameter is mandatory.
    .PARAMETER User
      If specified, sets the owner to the Users group (S-1-5-32-545). Otherwise, sets to Administrators group (S-1-5-32-544).
    .PARAMETER ResetAccess
      If specified, resets the access rules for the registry key to grant full control to the new owner.
    .PARAMETER Recurse
      If specified, applies ownership and access rule changes recursively to all subkeys.
    .EXAMPLE
      Reset-MyRegKeyOwner -Key "SOFTWARE\MyApp"
      Takes ownership of the "SOFTWARE\MyApp" registry key as Administrators.
    .EXAMPLE
      Reset-MyRegKeyOwner -Key "SOFTWARE\MyApp" -User
      Takes ownership of the "SOFTWARE\MyApp" registry key as Users.
    .EXAMPLE
      Reset-MyRegKeyOwner -Key "SOFTWARE\MyApp" -ResetAccess
      Takes ownership and resets access rules for the "SOFTWARE\MyApp" registry key.
    .EXAMPLE
      Reset-MyRegKeyOwner -Key "SOFTWARE\MyApp" -Recurse
      Takes ownership recursively for "SOFTWARE\MyApp" and all its subkeys.
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $False)]
    [Microsoft.Win32.RegistryKey]$Hive = [Microsoft.Win32.Registry]::LocalMachine,
    [parameter(Mandatory = $True)]
    [String]$Key,
    [Switch]$User,
    [Switch]$ResetAccess,
    [Switch]$Recurse
  )
  Write-Verbose -Message "Enter Function Reset-MyRegKeyOwner"

  if ($User.IsPresent)
  {
    $NewOwner = [System.Security.Principal.SecurityIdentifier]::New("S-1-5-32-545")
  }
  else
  {
    $NewOwner = [System.Security.Principal.SecurityIdentifier]::New("S-1-5-32-544")
  }

  Write-Verbose -Message "Key: $($Key)"
  $TempKey = $Hive.OpenSubKey($Key, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, [System.Security.AccessControl.RegistryRights]::TakeOwnership)
  $ACL = [System.Security.AccessControl.RegistrySecurity]::New()
  $ACL.SetOwner($NewOwner)
  $TempKey.SetAccessControl($ACL)
  $ACL.SetAccessRuleProtection($False, $False)
  $TempKey.SetAccessControl($ACL)

  if ($ResetAccess.IsPresent)
  {
    $TempKey = $TempKey.OpenSubKey("", [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, [System.Security.AccessControl.RegistryRights]::ChangePermissions)
    $Rule = [System.Security.AccessControl.RegistryAccessRule]::New($NewOwner, [System.Security.AccessControl.RegistryRights]::FullControl, [System.Security.AccessControl.InheritanceFlags]::ContainerInherit, [System.Security.AccessControl.PropagationFlags]::None, [System.Security.AccessControl.AccessControlType]::Allow)
    $ACL.ResetAccessRule($Rule)
    $TempKey.SetAccessControl($ACL)
  }

  if ($Recurse.IsPresent)
  {
    [Void]$PSBoundParameters.Remove("Key")
    [Void]$PSBoundParameters.Remove("ResetAccess")
    $TempKey = $TempKey.OpenSubKey("")
    foreach ($SubKey in @($TempKey.GetSubKeyNames()))
    {
      Reset-MyRegKeyOwner @PSBoundParameters -Key "$($Key)\$($SubKey)"
    }
  }

  Write-Verbose -Message "Exit Function Reset-MyRegKeyOwner"
}
#endregion function Reset-MyRegKeyOwner

#region function Set-MyISScriptData
function Set-MyISScriptData()
{
  <#
    .SYNOPSIS
      Writes Script Data to the Registry
    .DESCRIPTION
      Writes Script Data to the Registry
    .PARAMETER Script
     Name of the Regsitry Key to write the values under. Defaults to the name of the script.
    .PARAMETER Name
     Name of the Value to write
    .PARAMETER Value
      The Data to write
    .PARAMETER MultiValue
      Write Multiple values to the Registry
    .PARAMETER User
      Write to the HKCU Registry Hive
    .PARAMETER Computer
      Write to the HKLM Registry Hive
    .PARAMETER Bitness
      Specify 32/64 bit HKLM Registry Hive
    .EXAMPLE
      Set-MyISScriptData -Name "Name" -Value "Value"

      Write REG_SZ value to the HKCU Registry Hive under the Default Script Name registry key
    .EXAMPLE
      Set-MyISScriptData -Name "Name" -Value @("This", "That") -User -Script "ScriptName"

      Write REG_MULTI_SZ value to the HKCU Registry Hive under the Specified Script Name registry key

      Single element arrays will be written as REG_SZ. To ensure they are written as REG_MULTI_SZ
      Use @() or (,) when specifing the Value paramter value
    .EXAMPLE
      Set-MyISScriptData -Name "Name" -Value (,8) -Bitness "64" -Computer

      Write REG_MULTI_SZ value to the 64 bit HKLM Registry Hive under the Default Script Name registry key

      Number arrays are written to the registry as strings.
    .EXAMPLE
      Set-MyISScriptData -Name "Name" -Value 0 -Computer

      Write REG_DWORD value to the HKLM Registry Hive under the Default Script Name registry key
    .EXAMPLE
      Set-MyISScriptData -MultiValue @{"Name" = "MyName"; "Number" = 4; "Array" = @("First", 2, "3rd", 4)} -Computer -Bitness "32"

      Write multiple values to the 32 bit HKLM Registry Hive under the Default Script Name registry key
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "User")]
  param (
    [String]$Script = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName),
    [parameter(Mandatory = $True, ParameterSetName = "User")]
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [String]$Name,
    [parameter(Mandatory = $True, ParameterSetName = "User")]
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [Object]$Value,
    [parameter(Mandatory = $True, ParameterSetName = "UserMulti")]
    [parameter(Mandatory = $True, ParameterSetName = "CompMulti")]
    [HashTable]$MultiValue,
    [parameter(Mandatory = $False, ParameterSetName = "User")]
    [parameter(Mandatory = $False, ParameterSetName = "UserMulti")]
    [Switch]$User,
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [parameter(Mandatory = $True, ParameterSetName = "CompMulti")]
    [Switch]$Computer,
    [parameter(Mandatory = $False, ParameterSetName = "Comp")]
    [parameter(Mandatory = $False, ParameterSetName = "CompMulti")]
    [ValidateSet("32", "64", "All")]
    [String]$Bitness = "All"
  )
  Write-Verbose -Message "Enter Function Set-MyISScriptData"

  # Get Default Registry Paths
  $RegPaths = [System.Collections.ArrayList]::New()
  if ($Computer.IsPresent)
  {
    if ($Bitness -match "All|32")
    {
      [Void]$RegPaths.Add("Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node")
    }
    if ($Bitness -match "All|64")
    {
      [Void]$RegPaths.Add("Registry::HKEY_LOCAL_MACHINE\Software")
    }
  }
  else
  {
    [Void]$RegPaths.Add("Registry::HKEY_CURRENT_USER\Software")
  }

  # Create the Registry Keys if Needed.
  foreach ($RegPath in $RegPaths)
  {
    if ([String]::IsNullOrEmpty((Get-Item -Path "$RegPath\MyISScriptData" -ErrorAction "SilentlyContinue")))
    {
      try
      {
        [Void](New-Item -Path $RegPath -Name "MyISScriptData")
      }
      catch
      {
        throw "Error Creating Registry Key : MyISScriptData"
      }
    }
    if ([String]::IsNullOrEmpty((Get-Item -Path "$RegPath\MyISScriptData\$Script" -ErrorAction "SilentlyContinue")))
    {
      try
      {
        [Void](New-Item -Path "$RegPath\MyISScriptData" -Name $Script)
      }
      catch
      {
        throw "Error Creating Registry Key : $Script"
      }
    }
  }

  # Write the values to the registry.
  switch -regex ($PSCmdlet.ParameterSetName)
  {
    "Multi"
    {
      foreach ($Key in $MultiValue.Keys)
      {
        if ($MultiValue[$Key] -is [Array])
        {
          $Data = [String[]]$MultiValue[$Key]
        }
        else
        {
          $Data = $MultiValue[$Key]
        }
        foreach ($RegPath in $RegPaths)
        {
          [Void](Set-ItemProperty -Path "$RegPath\MyISScriptData\$Script" -Name $Key -Value $Data)
        }
      }
    }
    default
    {
      if ($Value -is [Array])
      {
        $Data = [String[]]$Value
      }
      else
      {
        $Data = $Value
      }
      foreach ($RegPath in $RegPaths)
      {
        [Void](Set-ItemProperty -Path "$RegPath\MyISScriptData\$Script" -Name $Name -Value $Data)
      }
    }
  }

  Write-Verbose -Message "Exit Function Set-MyISScriptData"
}
#endregion function Set-MyISScriptData

#region function Get-MyISScriptData
function Get-MyISScriptData()
{
  <#
    .SYNOPSIS
      Reads Script Data from the Registry
    .DESCRIPTION
      Reads Script Data from the Registry
    .PARAMETER Script
     Name of the Regsitry Key to read the values from. Defaults to the name of the script.
    .PARAMETER Name
     Name of the Value to read
    .PARAMETER User
      Read from the HKCU Registry Hive
    .PARAMETER Computer
      Read from the HKLM Registry Hive
    .PARAMETER Bitness
      Specify 32/64 bit HKLM Registry Hive
    .EXAMPLE
      $RegValues = Get-MyISScriptData -Name "Name"

      Read the value from the HKCU Registry Hive under the Default Script Name registry key
    .EXAMPLE
      $RegValues = Get-MyISScriptData -Name "Name" -User -Script "ScriptName"

      Read the value from the HKCU Registry Hive under the Specified Script Name registry key
    .EXAMPLE
      $RegValues = Get-MyISScriptData -Name "Name" -Computer

      Read the value from the 64 bit HKLM Registry Hive under the Default Script Name registry key
    .EXAMPLE
      $RegValues = Get-MyISScriptData -Name "Name" -Bitness "32" -Script "ScriptName" -Computer

      Read the value from the 32 bit HKLM Registry Hive under the Specified Script Name registry key
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "User")]
  param (
    [String]$Script = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName),
    [String[]]$Name = "*",
    [parameter(Mandatory = $False, ParameterSetName = "User")]
    [Switch]$User,
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [Switch]$Computer,
    [parameter(Mandatory = $False, ParameterSetName = "Comp")]
    [ValidateSet("32", "64")]
    [String]$Bitness = "64"
  )
  Write-Verbose -Message "Enter Function Get-MyISScriptData"

  # Get Default Registry Path
  if ($Computer.IsPresent)
  {
    if ($Bitness -eq "64")
    {
      $RegPath = "Registry::HKEY_LOCAL_MACHINE\Software"
    }
    else
    {
      $RegPath = "Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node"
    }
  }
  else
  {
    $RegPath = "Registry::HKEY_CURRENT_USER\Software"
  }

  # Get the values from the registry.
  Get-ItemProperty -Path "$RegPath\MyISScriptData\$Script" -ErrorAction "SilentlyContinue" | Select-Object -Property $Name

  Write-Verbose -Message "Exit Function Get-MyISScriptData"
}
#endregion function Get-MyISScriptData

#region function Remove-MyISScriptData
function Remove-MyISScriptData()
{
  <#
    .SYNOPSIS
      Removes Script Data from the Registry
    .DESCRIPTION
      Removes Script Data from the Registry
    .PARAMETER Script
     Name of the Regsitry Key to remove. Defaults to the name of the script.
    .PARAMETER User
      Remove from the HKCU Registry Hive
    .PARAMETER Computer
      Remove from the HKLM Registry Hive
    .PARAMETER Bitness
      Specify 32/64 bit HKLM Registry Hive
    .EXAMPLE
      Remove-MyISScriptData

      Removes the default script registry key from the HKCU Registry Hive
    .EXAMPLE
      Remove-MyISScriptData -User -Script "ScriptName"

      Removes the Specified Script Name registry key from the HKCU Registry Hive
    .EXAMPLE
      Remove-MyISScriptData -Computer

      Removes the default script registry key from the 32/64 bit HKLM Registry Hive
    .EXAMPLE
      Remove-MyISScriptData -Computer -Script "ScriptName" -Bitness "32"

      Removes the Specified Script Name registry key from the 32 bit HKLM Registry Hive
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding(DefaultParameterSetName = "User")]
  param (
    [String]$Script = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName),
    [parameter(Mandatory = $False, ParameterSetName = "User")]
    [Switch]$User,
    [parameter(Mandatory = $True, ParameterSetName = "Comp")]
    [Switch]$Computer,
    [parameter(Mandatory = $False, ParameterSetName = "Comp")]
    [ValidateSet("32", "64", "All")]
    [String]$Bitness = "All"
  )
  Write-Verbose -Message "Enter Function Remove-MyISScriptData"

  # Get Default Registry Paths
  $RegPaths = [System.Collections.ArrayList]::New()
  if ($Computer.IsPresent)
  {
    if ($Bitness -match "All|32")
    {
      [Void]$RegPaths.Add("Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node")
    }
    if ($Bitness -match "All|64")
    {
      [Void]$RegPaths.Add("Registry::HKEY_LOCAL_MACHINE\Software")
    }
  }
  else
  {
    [Void]$RegPaths.Add("Registry::HKEY_CURRENT_USER\Software")
  }

  # Remove the values from the registry.
  foreach ($RegPath in $RegPaths)
  {
    [Void](Remove-Item -Path "$RegPath\MyISScriptData\$Script")
  }

  Write-Verbose -Message "Exit Function Remove-MyISScriptData"
}
#endregion function Remove-MyISScriptData

#region function Get-EnvironmentVariable
function Get-EnvironmentVariable()
{
  <#
    .SYNOPSIS
      Retrieves environment variables from the local or remote workstation.
    .DESCRIPTION
      This function queries environment variables for a specified user on one or more computers using CIM/WMI.
    .PARAMETER ComputerName
      Specifies one or more computer names to query. Defaults to the local computer.
    .PARAMETER Variable
      Specifies the name of the environment variable to retrieve. Supports wildcards. Defaults to '%' (all variables).
    .PARAMETER UserName
      Specifies the user context for the environment variable. Defaults to '<SYSTEM>' for system-wide variables.
    .PARAMETER Credential
      Specifies a PSCredential object for authentication when connecting to remote computers.
    .EXAMPLE
      Get-EnvironmentVariable -Variable "Path"
      Retrieves the "Path" environment variable for the system on the local computer.
    .EXAMPLE
      Get-EnvironmentVariable -ComputerName "Server01","Server02" -Variable "TEMP" -UserName "DOMAIN\User"
      Retrieves the "TEMP" environment variable for the specified user on multiple remote computers.
    .NOTES
      Original Script By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $False, ValueFromPipeline = $True)]
    [String[]]$ComputerName = [System.Environment]::MachineName,
    [String]$Variable = "%",
    [String]$UserName = "<SYSTEM>",
    [PSCredential]$Credential = [PSCredential]::Empty
  )
  begin
  {
    Write-Verbose -Message "Enter Function Get-EnvironmentVariable Begin Block"

    $Query = "Select * from Win32_Environment Where Name like '$Variable' and UserName = '$UserName'"

    $SessionParams = @{
      "ComputerName" = ""
    }
    if ($PSBoundParameters.ContainsKey("Credential"))
    {
      [Void]$SessionParms.Add("Credential", $Credential)
    }

    Write-Verbose -Message "Exit Function Get-EnvironmentVariable Begin Block"
  }
  process
  {
    Write-Verbose -Message "Enter Function Get-EnvironmentVariable Process Block"

    foreach ($Computer in $ComputerName)
    {
      $SessionParams.ComputerName = $Computer
      Get-CimInstance -CimSession (New-CimSession @SessionParams) -Query $Query
    }

    Write-Verbose -Message "Exit Function Get-EnvironmentVariable Process Block"
  }
}
#endregion function Get-EnvironmentVariable

#region function Set-EnvironmentVariable
function Set-EnvironmentVariable()
{
  <#
    .SYNOPSIS
      Sets or creates an environment variable on the local or remote workstation.
    .DESCRIPTION
      This function sets the value of an environment variable for a specified user on one or more computers using CIM/WMI. If the variable does not exist, it will be created.
    .PARAMETER ComputerName
      Specifies one or more computer names to target. Defaults to the local computer.
    .PARAMETER Variable
      Specifies the name of the environment variable to set. This parameter is mandatory.
    .PARAMETER Value
      Specifies the value to assign to the environment variable.
    .PARAMETER UserName
      Specifies the user context for the environment variable. Defaults to '<SYSTEM>' for system-wide variables.
    .PARAMETER Credential
      Specifies a PSCredential object for authentication when connecting to remote computers.
    .EXAMPLE
      Set-EnvironmentVariable -Variable "Path" -Value "C:\MyPath"
      Sets the "Path" environment variable for the system on the local computer.
    .EXAMPLE
      Set-EnvironmentVariable -ComputerName "Server01","Server02" -Variable "TEMP" -Value "C:\Temp" -UserName "DOMAIN\User"
      Sets the "TEMP" environment variable for the specified user on multiple remote computers.
    .NOTES
      Original Script By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $False, ValueFromPipeline = $True)]
    [String[]]$ComputerName = [System.Environment]::MachineName,
    [parameter(Mandatory = $True)]
    [String]$Variable,
    [String]$Value,
    [String]$UserName = "<SYSTEM>",
    [PSCredential]$Credential = [PSCredential]::Empty
  )
  begin
  {
    Write-Verbose -Message "Enter Function Set-EnvironmentVariable Begin Block"

    $Query = "Select * from Win32_Environment Where Name = '$Variable' and UserName = '$UserName'"

    $SessionParams = @{
      "ComputerName" = ""
    }
    if ($PSBoundParameters.ContainsKey("Credential"))
    {
      [Void]$SessionParms.Add("Credential", $Credential)
    }

    Write-Verbose -Message "Exit Function Set-EnvironmentVariable Begin Block"
  }
  process
  {
    Write-Verbose -Message "Enter Function Set-EnvironmentVariable Process Block"

    foreach ($Computer in $ComputerName)
    {
      $SessionParams.ComputerName = $Computer
      $CimSession = New-CimSession @SessionParams
      if ([String]::IsNullOrEmpty(($Instance = Get-CimInstance -CimSession $CimSession -Query $Query)))
      {
        New-CimInstance -CimSession $CimSession -ClassName Win32_Environment -Property @{ "Name" = $Variable; "VariableValue" = $Value; "UserName" = $UserName }
      }
      else
      {
        Set-CimInstance -InputObject $Instance -Property @{ "Name" = $Variable; "VariableValue" = $Value } -PassThru
      }
      $CimSession.Close()
    }

    Write-Verbose -Message "Exit Function Set-EnvironmentVariable Process Block"
  }
}
#endregion function Set-EnvironmentVariable

#region function Remove-EnvironmentVariable
function Remove-EnvironmentVariable()
{
  <#
    .SYNOPSIS
      Removes an environment variable from the local or remote workstation.
    .DESCRIPTION
      This function deletes a specified environment variable for a given user on one or more computers using CIM/WMI. 
      It supports system-wide and user-specific variables and can authenticate to remote computers.
    .PARAMETER ComputerName
      Specifies one or more computer names to target. Defaults to the local computer.
    .PARAMETER Variable
      Specifies the name of the environment variable to remove. This parameter is mandatory.
    .PARAMETER UserName
      Specifies the user context for the environment variable. Defaults to '<SYSTEM>' for system-wide variables.
    .PARAMETER Credential
      Specifies a PSCredential object for authentication when connecting to remote computers.
    .EXAMPLE
      Remove-EnvironmentVariable -Variable "TEMP"
      Removes the "TEMP" environment variable for the system on the local computer.
    .EXAMPLE
      Remove-EnvironmentVariable -ComputerName "Server01","Server02" -Variable "Path" -UserName "DOMAIN\User"
      Removes the "Path" environment variable for the specified user on multiple remote computers.
    .NOTES
      Original Script By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [parameter(Mandatory = $False, ValueFromPipeline = $True)]
    [String[]]$ComputerName = [System.Environment]::MachineName,
    [parameter(Mandatory = $True)]
    [String]$Variable,
    [String]$UserName = "<SYSTEM>",
    [PSCredential]$Credential = [PSCredential]::Empty
  )
  begin
  {
    Write-Verbose -Message "Enter Function Remove-EnvironmentVariable Begin Block"

    $Query = "Select * from Win32_Environment Where Name = '$Variable' and UserName = '$UserName'"

    $SessionParams = @{
      "ComputerName" = ""
    }
    if ($PSBoundParameters.ContainsKey("Credential"))
    {
      [Void]$SessionParms.Add("Credential", $Credential)
    }

    Write-Verbose -Message "Exit Function Remove-EnvironmentVariable Begin Block"
  }
  process
  {
    Write-Verbose -Message "Enter Function Remove-EnvironmentVariable Process Block"

    foreach ($Computer in $ComputerName)
    {
      $SessionParams.ComputerName = $Computer
      $CimSession = New-CimSession @SessionParams
      if (-not [String]::IsNullOrEmpty(($Instance = Get-CimInstance -CimSession $CimSession -Query $Query)))
      {
        Remove-CimInstance -InputObject $Instance
      }
      $CimSession.Close()
    }

    Write-Verbose -Message "Exit Function Remove-EnvironmentVariable Process Block"
  }
}
#endregion function Remove-EnvironmentVariable

#endregion ******* Registry / Environement Variable Functions ********

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
function Start-MyRSJob()
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
  param (
    [parameter(Mandatory = $True, ParameterSetName = "RSPool")]
    [MyRSPool]$RSPool,
    [parameter(Mandatory = $False, ParameterSetName = "PoolName")]
    [String]$PoolName = "MyDefaultRSPool",
    [parameter(Mandatory = $True, ParameterSetName = "PoolID")]
    [Guid]$PoolID,
    [parameter(Mandatory = $False, ValueFromPipeline = $True)]
    [Object[]]$InputObject,
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

    if ($PSBoundParameters.ContainsKey("InputObject"))
    {
      ForEach ($Object in $InputObject)
      {
        # Create New PowerShell Instance with ScriptBlock
        $PowerShell = ([Management.Automation.PowerShell]::Create()).AddScript($ScriptBlock)
        # Set RunspacePool
        $PowerShell.RunspacePool = $TempPool.RunspacePool
        # Add Parameters
        [Void]$PowerShell.AddParameter($InputParam, $Object)
        if ($PSBoundParameters.ContainsKey("Parameters"))
        {
          [Void]$PowerShell.AddParameters($Parameters)
        }
        # set Job Name
        if (($Object -is [String]) -or ($Object -is [ValueType]))
        {
          $TempJobName = "$JobName - $($Object)"
        }
        else
        {
          $TempJobName = $($Object.$JobName)
        }
        [Void]$NewJobs.Add(([MyRSjob]::New($TempJobName, $PowerShell, $PowerShell.BeginInvoke(), $Object, $TempPool.Name, $TempPool.InstanceID)))
      }
    }
    else
    {
      # Create New PowerShell Instance with ScriptBlock
      $PowerShell = ([Management.Automation.PowerShell]::Create()).AddScript($ScriptBlock)
      # Set RunspacePool
      $PowerShell.RunspacePool = $TempPool.RunspacePool
      # Add Parameters
      if ($PSBoundParameters.ContainsKey("Parameters"))
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

    if ($NewJobs.Count)
    {
      $TempPool.Jobs.AddRange($NewJobs)
      # Return Jobs only if New RunspacePool
      if ($PassThru.IsPresent)
      {
        [MyRSJob[]]($NewJobs)
      }
      $NewJobs.Clear()
    }

    Write-Verbose -Message "Exit Function Start-MyRSJob End Block"
  }
}
#endregion function Start-MyRSJob

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

#region ******** RSPools Sample Code ********

#region function Test-Function
Function Test-Function
{
  <#
    .SYNOPSIS
      Test Function for RunspacePool ScriptBlock
    .DESCRIPTION
      Test Function for RunspacePool ScriptBlock
    .PARAMETER Value
      Value Command Line Parameter
    .EXAMPLE
      Test-Function -Value "String"
    .NOTES
      Original Function By Ken Sweet
    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param (
    [parameter(Mandatory = $False, HelpMessage = "Enter Value", ParameterSetName = "Default")]
    [Object[]]$Value = "Default Value"
  )
  Write-Verbose -Message "Enter Function Test-Function"

  Start-Sleep -Milliseconds (1000 * 5)
  ForEach ($Item in $Value)
  {
    "Return Value: `$Item = $Item"
  }

  Write-Verbose -Message "Exit Function Test-Function"
}
#endregion function Test-Function

#region Job $ScriptBlock
$ScriptBlock = {
  <#
    .SYNOPSIS
      Test RunspacePool ScriptBlock
    .DESCRIPTION
      Test RunspacePool ScriptBlock
    .PARAMETER InputObject
      InputObject passed to script
    .EXAMPLE
      Test-Script.ps1 -InputObject $InputObject
    .NOTES
      Original Script By Ken Sweet on 10/15/2017
      Updated Script By Ken Sweet on 02/04/2019

      Thread Script Variables
        [String]$Mutex - Exist only if -Mutex was specified on the Start-MyRSPool command line
        [HashTable]$SyncedHash - Always Exists, Default values $SyncedHash.Enabled = $True

    .LINK
  #>
  [CmdletBinding(DefaultParameterSetName = "ByValue")]
  Param (
    [parameter(Mandatory = $False, ParameterSetName = "ByValue")]
    [Object[]]$InputObject
  )

  # Generate Error Message to show in Error Buffer
  $ErrorActionPreference = "Continue"
  GenerateErrorMessage
  $ErrorActionPreference = "Stop"

  # Enable Verbose logging
  $VerbosePreference = "Continue"

  # Check is Thread is Enabled to Run
  if ($SyncedHash.Enabled)
  {
    # Call Imported Test Function
    Test-Function -Value $InputObject

    # Check if a Mutex exist
    if ([String]::IsNullOrEmpty($Mutex))
    {
      $HasMutex = $False
    }
    else
    {
      # Open and wait for Mutex
      $MyMutex = [System.Threading.Mutex]::OpenExisting($Mutex)
      [Void]($MyMutex.WaitOne())
      $HasMutex = $True
    }

    # Write Data to the Screen
    For ($Count = 0; $Count -le 8; $Count++)
    {
      Write-Host -Object "`$InputObject = $InputObject"
    }

    # Release the Mutex if it Exists
    if ($HasMutex)
    {
      $MyMutex.ReleaseMutex()
    }
  }
  else
  {
    "Return Value: RSJob was Canceled"
  }
}
#endregion

#region $WaitScript
$WaitScript = {
  Write-Host -Object "Completed $(@(Get-MyRSJob | Where-Object -FilterScript { $PSItem.State -eq 'Completed' }).Count) Jobs"
  Start-Sleep -Milliseconds 1000
}
#endregion

<#
$TestFunction = @{}
$TestFunction.Add("Test-Function", (Get-Command -Type Function -Name Test-Function).ScriptBlock)

# Start and Get RSPool
$RSPool = Start-MyRSPool -MaxJobs 8 -Functions $TestFunction -PassThru #-Mutex "TestMutex"

# Create new RunspacePool and start 5 Jobs
1..10 | Start-MyRSJob -ScriptBlock $ScriptBlock -PassThru | Out-String

# Add 5 new Jobs to an existing RunspacePool
11..20 | Start-MyRSJob -ScriptBlock $ScriptBlock -PassThru | Out-String

# Disable Thread Script
#$RSPool.SyncedHash.Enabled = $False

# Wait for all Jobs to Complete or Fail
Get-MyRSJob | Wait-MyRSJob -SciptBlock $WaitScript -PassThru | Out-String

# Receive Completed Jobs and Remove them
Get-MyRSJob | Receive-MyRSJob -AutoRemove

# Close RunspacePool
Close-MyRSPool
#>

#endregion ******** RSPools Sample Code ********

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
# Show AlertMessage Function
# --------------------------
#region function Show-AlertMessage
Function Show-AlertMessage ()
{
  <#
    .SYNOPSIS
      Shows Show-AlertMessage
    .DESCRIPTION
      Shows Show-AlertMessage
    .PARAMETER Title
      Title of the Show-AlertMessage Dialog Window
    .PARAMETER Message
      Alert Message to Display
    .PARAMETER Width
      Width of Show-AlertMessage Dialog Window
    .PARAMETER MsgType
      Type of Alert Message to SHow
    .EXAMPLE
      Show-AlertMessage -Title "Example Alert" -Message "Show Success, Warning, Error, and Information Alert Messages"
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  Param (
    [parameter(Mandatory = $True)]
    [String]$Title,
    [parameter(Mandatory = $True)]
    [String]$Message,
    [Int]$Width = 25,
    [ValidateSet("Success", "Warning", "Error", "Info")]
    [String]$MsgType = "Info"
  )
  Write-Verbose -Message "Enter Function Show-AlertMessage"

  #region ******** Begin **** $AlertMessage **** Begin ********

  # ************************************************
  # $AlertMessage Form
  # ************************************************
  #region $AlertMessageForm = [System.Windows.Forms.Form]::New()
  $AlertMessageForm = [System.Windows.Forms.Form]::New()
  $AlertMessageForm.BackColor = [MyConfig]::Colors.TextBack
  $AlertMessageForm.Font = [MyConfig]::Font.Regular
  $AlertMessageForm.ForeColor = [MyConfig]::Colors.TextFore
  $AlertMessageForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D

  switch ($MsgType)
  {
    "Success"
    {
      $AlertMessageForm.Icon = [System.Drawing.SystemIcons]::Shield
      Break
    }
    "Warning"
    {
      $AlertMessageForm.Icon = [System.Drawing.SystemIcons]::Warning
      Break
    }
    "Error"
    {
      $AlertMessageForm.Icon = [System.Drawing.SystemIcons]::Error
      Break
    }
    "Info"
    {
      $AlertMessageForm.Icon = [System.Drawing.SystemIcons]::Information
      Break
    }
  }
  $AlertMessageForm.KeyPreview = $True
  $AlertMessageForm.MaximizeBox = $False
  $AlertMessageForm.MinimizeBox = $False
  $AlertMessageForm.Name = "AlertMessageForm"
  $AlertMessageForm.Owner = $PILForm
  $AlertMessageForm.ShowInTaskbar = $False
  $AlertMessageForm.Size = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * 25))
  $AlertMessageForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $AlertMessageForm.Tag = @{ "Cancel" = $False; "Pause" = $False }
  $AlertMessageForm.Text = $Title
  #endregion $AlertMessageForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-AlertMessageFormKeyDown ********
  Function Start-AlertMessageFormKeyDown
  {
  <#
    .SYNOPSIS
      KeyDown Event for the AlertMessage Form Control
    .DESCRIPTION
      KeyDown Event for the AlertMessage Form Control
    .PARAMETER Sender
       The Form Control that fired the KeyDown Event
    .PARAMETER EventArg
       The Event Arguments for the Form KeyDown Event
    .EXAMPLE
       Start-AlertMessageFormKeyDown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter KeyDown Event for `$AlertMessageForm"

    [MyConfig]::AutoExit = 0

    If ($EventArg.KeyCode -in ([System.Windows.Forms.Keys]::Enter, [System.Windows.Forms.Keys]::Space, [System.Windows.Forms.Keys]::Escape))
    {
      $AlertMessageForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$AlertMessageForm"
  }
  #endregion ******** Function Start-AlertMessageFormKeyDown ********
  $AlertMessageForm.add_KeyDown({ Start-AlertMessageFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-AlertMessageFormShown ********
  Function Start-AlertMessageFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the $AlertMessage Form Control
      .DESCRIPTION
        Shown Event for the $AlertMessage Form Control
      .PARAMETER Sender
         The Form Control that fired the Shown Event
      .PARAMETER EventArg
         The Event Arguments for the Form Shown Event
      .EXAMPLE
         Start-AlertMessageFormShown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Shown Event for `$AlertMessageForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Shown Event for `$AlertMessageForm"
  }
  #endregion ******** Function Start-AlertMessageFormShown ********
  $AlertMessageForm.add_Shown({ Start-AlertMessageFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for $AlertMessage Form ********

  # ************************************************
  # $AlertMessage Panel
  # ************************************************
  #region $AlertMessagePanel = [System.Windows.Forms.Panel]::New()
  $AlertMessagePanel = [System.Windows.Forms.Panel]::New()
  $AlertMessageForm.Controls.Add($AlertMessagePanel)
  $AlertMessagePanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
  $AlertMessagePanel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $AlertMessagePanel.Name = "AlertMessagePanel"
  $AlertMessagePanel.Size = [System.Drawing.Size]::New(($AlertMessageForm.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($AlertMessageForm.ClientSize.Height - ([MyConfig]::FormSpacer * 2)))
  #endregion $AlertMessagePanel = [System.Windows.Forms.Panel]::New()

  #region ******** $AlertMessagePanel Controls ********

  #region $AlertMessageTopLabel = [System.Windows.Forms.Label]::New()
  $AlertMessageTopLabel = [System.Windows.Forms.Label]::New()
  $AlertMessagePanel.Controls.Add($AlertMessageTopLabel)
  $AlertMessageTopLabel.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")

  Switch ($MsgType)
  {
    "Info"
    {
      $AlertMessageTopLabel.BackColor = [MyConfig]::Colors.TextInfo
      Break
    }
    "Success"
    {
      $AlertMessageTopLabel.BackColor = [MyConfig]::Colors.TextGood
      Break
    }
    "Warning"
    {
      $AlertMessageTopLabel.BackColor = [MyConfig]::Colors.TextWarn
      Break
    }
    "Error"
    {
      $AlertMessageTopLabel.BackColor = [MyConfig]::Colors.TextBad
      Break
    }
  }
  $AlertMessageTopLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
  $AlertMessageTopLabel.Font = [MyConfig]::Font.Title
  $AlertMessageTopLabel.ForeColor = [MyConfig]::Colors.TextBack
  $AlertMessageTopLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $AlertMessageTopLabel.Name = "AlertMessageTopLabel"
  $AlertMessageTopLabel.Size = [System.Drawing.Size]::New(($AlertMessagePanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), $AlertMessageTopLabel.PreferredHeight)
  $AlertMessageTopLabel.Text = $Title
  $AlertMessageTopLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
  #endregion $AlertMessageTopLabel = [System.Windows.Forms.Label]::New()

  #region $AlertMessageBtmLabel = [System.Windows.Forms.Label]::New()
  $AlertMessageBtmLabel = [System.Windows.Forms.Label]::New()
  $AlertMessagePanel.Controls.Add($AlertMessageBtmLabel)
  $AlertMessageBtmLabel.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $AlertMessageBtmLabel.BackColor = [MyConfig]::Colors.TextBack
  $AlertMessageBtmLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
  $AlertMessageBtmLabel.Font = [MyConfig]::Font.Bold
  $AlertMessageBtmLabel.ForeColor = [MyConfig]::Colors.TextFore
  $AlertMessageBtmLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($AlertMessageTopLabel.Bottom + [MyConfig]::FormSpacer))
  $AlertMessageBtmLabel.Name = "AlertMessageBtmLabel"
  $AlertMessageBtmLabel.Size = [System.Drawing.Size]::New($AlertMessageTopLabel.Width, ($AlertMessageTopLabel.Width - ($AlertMessageBtmLabel.Top * 3)))
  $AlertMessageBtmLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
  $AlertMessageBtmLabel.Text = $Message
  #endregion $AlertMessageBtmLabel = [System.Windows.Forms.Label]::New()

  $AlertMessagePanel.ClientSize = [System.Drawing.Size]::New($AlertMessagePanel.ClientSize.Width, ($AlertMessageBtmLabel.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $AlertMessagePanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  # ************************************************
  # $AlertMessageBtm Panel
  # ************************************************
  #region $AlertMessageBtmPanel = [System.Windows.Forms.Panel]::New()
  $AlertMessageBtmPanel = [System.Windows.Forms.Panel]::New()
  $AlertMessageForm.Controls.Add($AlertMessageBtmPanel)
  $AlertMessageBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $AlertMessageBtmPanel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, $AlertMessagePanel.Bottom)
  $AlertMessageBtmPanel.Name = "AlertMessageBtmPanel"
  #endregion $AlertMessageBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $AlertMessageBtmPanel Controls ********

  $NumButtons = 3
  $TempSpace = [Math]::Floor($AlertMessageBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $AlertMessageBtmMidButton = [System.Windows.Forms.Button]::New()
  $AlertMessageBtmMidButton = [System.Windows.Forms.Button]::New()
  $AlertMessageBtmPanel.Controls.Add($AlertMessageBtmMidButton)
  $AlertMessageBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $AlertMessageBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $AlertMessageBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $AlertMessageBtmMidButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
  $AlertMessageBtmMidButton.Font = [MyConfig]::Font.Bold
  $AlertMessageBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $AlertMessageBtmMidButton.Location = [System.Drawing.Point]::New(($TempWidth + ([MyConfig]::FormSpacer * 2)), [MyConfig]::FormSpacer)
  $AlertMessageBtmMidButton.Name = "AlertMessageBtmMidButton"
  $AlertMessageBtmMidButton.TabStop = $True
  $AlertMessageBtmMidButton.Text = "OK"
  $AlertMessageBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $AlertMessageBtmMidButton.PreferredSize.Height)
  #endregion $AlertMessageBtmMidButton = [System.Windows.Forms.Button]::New()

  $AlertMessageBtmPanel.ClientSize = [System.Drawing.Size]::New($AlertMessagePanel.ClientSize.Width, (($AlertMessageBtmPanel.Controls[$AlertMessageBtmPanel.Controls.Count - 1]).Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $AlertMessageBtmPanel Controls ********

  $AlertMessageForm.ClientSize = [System.Drawing.Size]::New($AlertMessageForm.ClientSize.Width, $AlertMessageBtmPanel.Bottom)

  #endregion ******** Controls for $AlertMessage Form ********

  #endregion ******** End **** $Show-AlertMessage **** End ********

  $DialogResult = $AlertMessageForm.ShowDialog($PILForm)
  $AlertMessageForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Show-AlertMessage"
}
#endregion function Show-AlertMessage

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
      $DialogResult = Get-UserResponse -Title "Get User Text - Single" -Message "Show this Sample Message Prompt to the User"
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
    $MultiTextBoxInputLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
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
    $MultiTextBoxInputTextBox.MaxLength = 25
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

# ------------------------------
# Get RadioButtonOption Function
# ------------------------------
#region RadioButtonOption Result Class
Class RadioButtonOption
{
  [Bool]$Success
  [Object]$DialogResult
  [HashTable]$Item = @{}

  RadioButtonOption ([Bool]$Success, [Object]$DialogResult)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
  }

  RadioButtonOption ([Bool]$Success, [Object]$DialogResult, [HashTable]$Item)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.Item = $Item
  }
}
#endregion RadioButtonOption Result Class

#region function Get-RadioButtonOption
Function Get-RadioButtonOption ()
{
  <#
    .SYNOPSIS
      Shows Get-RadioButtonOption
    .DESCRIPTION
      Shows Get-RadioButtonOption
    .PARAMETER Title
      Title of the Get-RadioButtonOption Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER Selected
      Selected RadioButtonOption
    .PARAMETER OrderedItems
      Ordered List (HashTable) if Names and Values
    .PARAMETER Width
      With if Get-RadioButtonOption Dialog Window
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $OrderedItems = [Ordered]@{ "First Choice in the List." = "1"; "Pick this Item!" = "2"; "No, Pick this one!!" = "3"; "Never Pick this Option." = "4"}
      $DialogResult = Get-RadioButtonOption -Title "RadioButton Option" -Message "Show this Sample Message Prompt to the User" -OrderedItems $OrderedItems -Selected "4"
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
    [String]$Message = "Status Message",
    [Object]$Selected = "",
    [parameter(Mandatory = $True)]
    [System.Collections.Specialized.OrderedDictionary]$OrderedItems,
    [Int]$Width = 35,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel"
  )
  Write-Verbose -Message "Enter Function Get-RadioButtonOption"

  #region ******** Begin **** RadioButtonOption **** Begin ********

  # ************************************************
  # RadioButtonOption Form
  # ************************************************
  #region $RadioButtonOptionForm = [System.Windows.Forms.Form]::New()
  $RadioButtonOptionForm = [System.Windows.Forms.Form]::New()
  $RadioButtonOptionForm.BackColor = [MyConfig]::Colors.Back
  $RadioButtonOptionForm.Font = [MyConfig]::Font.Regular
  $RadioButtonOptionForm.ForeColor = [MyConfig]::Colors.Fore
  $RadioButtonOptionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $RadioButtonOptionForm.Icon = $PILForm.Icon
  $RadioButtonOptionForm.KeyPreview = $True
  $RadioButtonOptionForm.MaximizeBox = $False
  $RadioButtonOptionForm.MinimizeBox = $False
  $RadioButtonOptionForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), 0)
  $RadioButtonOptionForm.Name = "RadioButtonOptionForm"
  $RadioButtonOptionForm.Owner = $PILForm
  $RadioButtonOptionForm.ShowInTaskbar = $False
  $RadioButtonOptionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $RadioButtonOptionForm.Text = $Title
  #endregion $RadioButtonOptionForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-RadioButtonOptionFormKeyDown ********
  Function Start-RadioButtonOptionFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the RadioButtonOption Form Control
      .DESCRIPTION
        KeyDown Event for the RadioButtonOption Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-RadioButtonOptionFormKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By CDUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$RadioButtonOptionForm"

    [MyConfig]::AutoExit = 0

    If ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $RadioButtonOptionForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$RadioButtonOptionForm"
  }
  #endregion ******** Function Start-RadioButtonOptionFormKeyDown ********
  $RadioButtonOptionForm.add_KeyDown({ Start-RadioButtonOptionFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-RadioButtonOptionFormShown ********
  Function Start-RadioButtonOptionFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the RadioButtonOption Form Control
      .DESCRIPTION
        Shown Event for the RadioButtonOption Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-RadioButtonOptionFormShown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Shown Event for `$RadioButtonOptionForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    Write-Verbose -Message "Exit Shown Event for `$RadioButtonOptionForm"
  }
  #endregion ******** Function Start-RadioButtonOptionFormShown ********
  $RadioButtonOptionForm.add_Shown({ Start-RadioButtonOptionFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for RadioButtonOption Form ********

  # ************************************************
  # RadioButtonOption Panel
  # ************************************************
  #region $RadioButtonOptionPanel = [System.Windows.Forms.Panel]::New()
  $RadioButtonOptionPanel = [System.Windows.Forms.Panel]::New()
  $RadioButtonOptionForm.Controls.Add($RadioButtonOptionPanel)
  $RadioButtonOptionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $RadioButtonOptionPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $RadioButtonOptionPanel.Name = "RadioButtonOptionPanel"
  #$RadioButtonOptionPanel.Padding = [System.Windows.Forms.Padding]::New(([MyConfig]::FormSpacer * [MyConfig]::FormSpacer), 0, ([MyConfig]::FormSpacer * [MyConfig]::FormSpacer), 0)
  #endregion $RadioButtonOptionPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $RadioButtonOptionPanel Controls ********

  If ($PSBoundParameters.ContainsKey("Message"))
  {
    #region $RadioButtonOptionLabel = [System.Windows.Forms.Label]::New()
    $RadioButtonOptionLabel = [System.Windows.Forms.Label]::New()
    $RadioButtonOptionPanel.Controls.Add($RadioButtonOptionLabel)
    $RadioButtonOptionLabel.ForeColor = [MyConfig]::Colors.LabelFore
    $RadioButtonOptionLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
    $RadioButtonOptionLabel.Name = "RadioButtonOptionLabel"
    $RadioButtonOptionLabel.Size = [System.Drawing.Size]::New(($RadioButtonOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
    $RadioButtonOptionLabel.Text = $Message
    $RadioButtonOptionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #endregion $RadioButtonOptionLabel = [System.Windows.Forms.Label]::New()

    # Returns the minimum size required to display the text
    $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($RadioButtonOptionLabel.Text, [MyConfig]::Font.Regular, $RadioButtonOptionLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
    $RadioButtonOptionLabel.Size = [System.Drawing.Size]::New(($RadioButtonOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

    $TempBottom = $RadioButtonOptionLabel.Bottom + [MyConfig]::FormSpacer
  }
  Else
  {
    $TempBottom = [MyConfig]::FormSpacer
  }

  # ************************************************
  # RadioButtonOption GroupBox
  # ************************************************
  #region $RadioButtonOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  $RadioButtonOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $RadioButtonOptionPanel.Controls.Add($RadioButtonOptionGroupBox)
  $RadioButtonOptionGroupBox.BackColor = [MyConfig]::Colors.Back
  $RadioButtonOptionGroupBox.Font = [MyConfig]::Font.Regular
  $RadioButtonOptionGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $RadioButtonOptionGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($TempBottom + [MyConfig]::FormSpacer))
  $RadioButtonOptionGroupBox.Name = "RadioButtonOptionGroupBox"
  $RadioButtonOptionGroupBox.Size = [System.Drawing.Size]::New(($RadioButtonOptionPanel.Width - ([MyConfig]::FormSpacer * 2)), 23)
  $RadioButtonOptionGroupBox.Text = $Null
  #endregion $RadioButtonOptionGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $RadioButtonOptionGroupBox Controls ********

  $RadioButtonOptionNumber = 0
  $GroupBottom = [MyConfig]::Font.Height
  ForEach ($Key In $OrderedItems.Keys)
  {
    #region $RadioButtonOptionRadioButton = [System.Windows.Forms.RadioButton]::New()
    $RadioButtonOptionRadioButton = [System.Windows.Forms.RadioButton]::New()
    $RadioButtonOptionGroupBox.Controls.Add($RadioButtonOptionRadioButton)
    #$RadioButtonOptionRadioButton.AutoCheck = $True
    $RadioButtonOptionRadioButton.AutoSize = $True
    $RadioButtonOptionRadioButton.BackColor = [MyConfig]::Colors.Back
    $RadioButtonOptionRadioButton.Checked = ($OrderedItems[$Key] -eq $Selected)
    $RadioButtonOptionRadioButton.Font = [MyConfig]::Font.Regular
    $RadioButtonOptionRadioButton.ForeColor = [MyConfig]::Colors.LabelFore
    $RadioButtonOptionRadioButton.Location = [System.Drawing.Point]::New(([MyConfig]::FormSpacer * [MyConfig]::FormSpacer), $GroupBottom)
    $RadioButtonOptionRadioButton.Name = "RadioChoice$($RadioButtonOptionNumber)"
    $RadioButtonOptionRadioButton.Tag = $OrderedItems[$Key]
    $RadioButtonOptionRadioButton.Text = $Key
    #endregion $RadioButtonOptionRadioButton = [System.Windows.Forms.RadioButton]::New()

    $GroupBottom = ($RadioButtonOptionRadioButton.Bottom + [MyConfig]::FormSpacer)
    $RadioButtonOptionNumber += 1
  }

  $RadioButtonOptionGroupBox.ClientSize = [System.Drawing.Size]::New($RadioButtonOptionGroupBox.ClientSize.Width, ($GroupBottom + [MyConfig]::FormSpacer))

  #endregion ******** $RadioButtonOptionGroupBox Controls ********

  #endregion ******** $RadioButtonOptionPanel Controls ********

  # ************************************************
  # RadioButtonOptionBtm Panel
  # ************************************************
  #region $RadioButtonOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $RadioButtonOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $RadioButtonOptionForm.Controls.Add($RadioButtonOptionBtmPanel)
  $RadioButtonOptionBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $RadioButtonOptionBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $RadioButtonOptionBtmPanel.Name = "RadioButtonOptionBtmPanel"
  #endregion $RadioButtonOptionBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $RadioButtonOptionBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($RadioButtonOptionBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $RadioButtonOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $RadioButtonOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $RadioButtonOptionBtmPanel.Controls.Add($RadioButtonOptionBtmLeftButton)
  $RadioButtonOptionBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $RadioButtonOptionBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $RadioButtonOptionBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $RadioButtonOptionBtmLeftButton.Font = [MyConfig]::Font.Bold
  $RadioButtonOptionBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $RadioButtonOptionBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $RadioButtonOptionBtmLeftButton.Name = "RadioButtonOptionBtmLeftButton"
  $RadioButtonOptionBtmLeftButton.TabIndex = 1
  $RadioButtonOptionBtmLeftButton.TabStop = $True
  $RadioButtonOptionBtmLeftButton.Text = $ButtonLeft
  $RadioButtonOptionBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $RadioButtonOptionBtmLeftButton.PreferredSize.Height)
  #endregion $RadioButtonOptionBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-RadioButtonOptionBtmLeftButtonClick ********
  Function Start-RadioButtonOptionBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the RadioButtonOptionBtmLeft Button Control
      .DESCRIPTION
        Click Event for the RadioButtonOptionBtmLeft Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-RadioButtonOptionBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By CDUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$RadioButtonOptionBtmLeftButton"

    [MyConfig]::AutoExit = 0

    If (@($RadioButtonOptionGroupBox.Controls | Where-Object -FilterScript { ($PSItem.GetType().Name -eq "RadioButton") -and $PSItem.Checked }).Count -eq 1)
    {
      $RadioButtonOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    Else
    {
      [Void][System.Windows.Forms.MessageBox]::Show($RadioButtonOptionForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }

    Write-Verbose -Message "Exit Click Event for `$RadioButtonOptionBtmLeftButton"
  }
  #endregion ******** Function Start-RadioButtonOptionBtmLeftButtonClick ********
  $RadioButtonOptionBtmLeftButton.add_Click({ Start-RadioButtonOptionBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $RadioButtonOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $RadioButtonOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $RadioButtonOptionBtmPanel.Controls.Add($RadioButtonOptionBtmMidButton)
  $RadioButtonOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $RadioButtonOptionBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $RadioButtonOptionBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $RadioButtonOptionBtmMidButton.Font = [MyConfig]::Font.Bold
  $RadioButtonOptionBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $RadioButtonOptionBtmMidButton.Location = [System.Drawing.Point]::New(($RadioButtonOptionBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $RadioButtonOptionBtmMidButton.Name = "RadioButtonOptionBtmMidButton"
  $RadioButtonOptionBtmMidButton.TabIndex = 2
  $RadioButtonOptionBtmMidButton.TabStop = $True
  $RadioButtonOptionBtmMidButton.Text = $ButtonMid
  $RadioButtonOptionBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $RadioButtonOptionBtmMidButton.PreferredSize.Height)
  #endregion $RadioButtonOptionBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-RadioButtonOptionBtmMidButtonClick ********
  Function Start-RadioButtonOptionBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the RadioButtonOptionBtmMid Button Control
      .DESCRIPTION
        Click Event for the RadioButtonOptionBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-RadioButtonOptionBtmMidButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By CDUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$RadioButtonOptionBtmMidButton"

    [MyConfig]::AutoExit = 0

    ForEach ($RadioButton In @($RadioButtonOptionGroupBox.Controls | Where-Object -FilterScript { $PSItem.Name -Like "RadioChoice*" }))
    {
      $RadioButton.Checked = ($RadioButton.Tag -eq $Selected)
    }

    Write-Verbose -Message "Exit Click Event for `$RadioButtonOptionBtmMidButton"
  }
  #endregion ******** Function Start-RadioButtonOptionBtmMidButtonClick ********
  $RadioButtonOptionBtmMidButton.add_Click({ Start-RadioButtonOptionBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $RadioButtonOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $RadioButtonOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $RadioButtonOptionBtmPanel.Controls.Add($RadioButtonOptionBtmRightButton)
  $RadioButtonOptionBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $RadioButtonOptionBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $RadioButtonOptionBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $RadioButtonOptionBtmRightButton.Font = [MyConfig]::Font.Bold
  $RadioButtonOptionBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $RadioButtonOptionBtmRightButton.Location = [System.Drawing.Point]::New(($RadioButtonOptionBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $RadioButtonOptionBtmRightButton.Name = "RadioButtonOptionBtmRightButton"
  $RadioButtonOptionBtmRightButton.TabIndex = 3
  $RadioButtonOptionBtmRightButton.TabStop = $True
  $RadioButtonOptionBtmRightButton.Text = $ButtonRight
  $RadioButtonOptionBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $RadioButtonOptionBtmRightButton.PreferredSize.Height)
  #endregion $RadioButtonOptionBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-RadioButtonOptionBtmRightButtonClick ********
  Function Start-RadioButtonOptionBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the RadioButtonOptionBtmRight Button Control
      .DESCRIPTION
        Click Event for the RadioButtonOptionBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-RadioButtonOptionBtmRightButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By CDUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$RadioButtonOptionBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here

    $RadioButtonOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$RadioButtonOptionBtmRightButton"
  }
  #endregion ******** Function Start-RadioButtonOptionBtmRightButtonClick ********
  $RadioButtonOptionBtmRightButton.add_Click({ Start-RadioButtonOptionBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $RadioButtonOptionBtmPanel.ClientSize = [System.Drawing.Size]::New(($RadioButtonOptionBtmRightButton.Right + [MyConfig]::FormSpacer), ($RadioButtonOptionBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $RadioButtonOptionBtmPanel Controls ********

  $RadioButtonOptionForm.ClientSize = [System.Drawing.Size]::New($RadioButtonOptionForm.ClientSize.Width, ($RadioButtonOptionGroupBox.Bottom + [MyConfig]::FormSpacer + $RadioButtonOptionBtmPanel.Height))

  #endregion ******** Controls for RadioButtonOption Form ********

  #endregion ******** End **** RadioButtonOption **** End ********

  $DialogResult = $RadioButtonOptionForm.ShowDialog($PILForm)
  If ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
  {
    $TempItem = @{}
    $RadioButtonOptionGroupBox.Controls | Where-Object -FilterScript { $PSItem.Name -Like "RadioChoice*" -and $PSItem.Checked } | ForEach-Object -Process { $TempItem.Add($PSItem.Text, $PSItem.Tag) }
    [RadioButtonOption]::New($True, $DialogResult, $TempItem)
  }
  Else
  {
    [RadioButtonOption]::New($False, $DialogResult)
  }

  $RadioButtonOptionForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-RadioButtonOption"
}
#endregion function Get-RadioButtonOption

# ------------------------------
# Get CheckBoxOption Function
# ------------------------------
#region CheckBoxOption Result Class
Class CheckBoxOption
{
  [Bool]$Success
  [Object]$DialogResult
  [HashTable]$Items = @{}

  CheckBoxOption ([Bool]$Success, [Object]$DialogResult)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
  }

  CheckBoxOption ([Bool]$Success, [Object]$DialogResult, [HashTable]$Items)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.Items = $Items
  }
}
#endregion CheckBoxOption Result Class

#region function Get-CheckBoxOption
Function Get-CheckBoxOption ()
{
  <#
    .SYNOPSIS
      Shows Get-CheckBoxOption
    .DESCRIPTION
      Shows Get-CheckBoxOption
    .PARAMETER Title
      Title of the Get-CheckBoxOption Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER Selected
      Selected Items
    .PARAMETER OrderedItems
      Ordered List (HashTable) if Names and Values
    .PARAMETER Width
      With of Get-CheckBoxOption Dialog Window
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $OrderedItems = [Ordered]@{ "First Choice in the List." = "1"; "Pick this Item!" = "2"; "No, Pick this one!!" = "3"; "Never Pick this Option." = "4" }
      $DialogResult = Get-CheckBoxOption -Title "Get CheckBox Option" -Message "Show this Sample Message Prompt to the User" -OrderedItems $OrderedItems -Selected @("1", "4") -Required
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
    [String]$Message = "Status Message",
    [Object[]]$Selected = @(),
    [parameter(Mandatory = $True)]
    [System.Collections.Specialized.OrderedDictionary]$OrderedItems,
    [Int]$Width = 35,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel",
    [Switch]$Required
  )
  Write-Verbose -Message "Enter Function Get-CheckBoxOption"

  #region ******** Begin **** CheckBoxOption **** Begin ********

  # ************************************************
  # CheckBoxOption Form
  # ************************************************
  #region $CheckBoxOptionForm = [System.Windows.Forms.Form]::New()
  $CheckBoxOptionForm = [System.Windows.Forms.Form]::New()
  $CheckBoxOptionForm.BackColor = [MyConfig]::Colors.Back
  $CheckBoxOptionForm.Font = [MyConfig]::Font.Regular
  $CheckBoxOptionForm.ForeColor = [MyConfig]::Colors.Fore
  $CheckBoxOptionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $CheckBoxOptionForm.Icon = $PILForm.Icon
  $CheckBoxOptionForm.KeyPreview = $True
  $CheckBoxOptionForm.MaximizeBox = $False
  $CheckBoxOptionForm.MinimizeBox = $False
  $CheckBoxOptionForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), 0)
  $CheckBoxOptionForm.Name = "CheckBoxOptionForm"
  $CheckBoxOptionForm.Owner = $PILForm
  $CheckBoxOptionForm.ShowInTaskbar = $False
  $CheckBoxOptionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $CheckBoxOptionForm.Text = $Title
  #endregion $CheckBoxOptionForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-CheckBoxOptionFormKeyDown ********
  Function Start-CheckBoxOptionFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the CheckBoxOption Form Control
      .DESCRIPTION
        KeyDown Event for the CheckBoxOption Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-CheckBoxOptionFormKeyDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By CDUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter KeyDown Event for `$CheckBoxOptionForm"

    [MyConfig]::AutoExit = 0

    If ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $CheckBoxOptionForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$CheckBoxOptionForm"
  }
  #endregion ******** Function Start-CheckBoxOptionFormKeyDown ********
  $CheckBoxOptionForm.add_KeyDown({ Start-CheckBoxOptionFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-CheckBoxOptionFormShown ********
  Function Start-CheckBoxOptionFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the CheckBoxOption Form Control
      .DESCRIPTION
        Shown Event for the CheckBoxOption Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-CheckBoxOptionFormShown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Shown Event for `$CheckBoxOptionForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    Write-Verbose -Message "Exit Shown Event for `$CheckBoxOptionForm"
  }
  #endregion ******** Function Start-CheckBoxOptionFormShown ********
  $CheckBoxOptionForm.add_Shown({ Start-CheckBoxOptionFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for CheckBoxOption Form ********

  # ************************************************
  # CheckBoxOption Panel
  # ************************************************
  #region $CheckBoxOptionPanel = [System.Windows.Forms.Panel]::New()
  $CheckBoxOptionPanel = [System.Windows.Forms.Panel]::New()
  $CheckBoxOptionForm.Controls.Add($CheckBoxOptionPanel)
  $CheckBoxOptionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $CheckBoxOptionPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $CheckBoxOptionPanel.Name = "CheckBoxOptionPanel"
  #$CheckBoxOptionPanel.Padding = [System.Windows.Forms.Padding]::New(([MyConfig]::FormSpacer * [MyConfig]::FormSpacer), 0, ([MyConfig]::FormSpacer * [MyConfig]::FormSpacer), 0)
  #endregion $CheckBoxOptionPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $CheckBoxOptionPanel Controls ********

  If ($PSBoundParameters.ContainsKey("Message"))
  {
    #region $CheckBoxOptionLabel = [System.Windows.Forms.Label]::New()
    $CheckBoxOptionLabel = [System.Windows.Forms.Label]::New()
    $CheckBoxOptionPanel.Controls.Add($CheckBoxOptionLabel)
    $CheckBoxOptionLabel.ForeColor = [MyConfig]::Colors.LabelFore
    $CheckBoxOptionLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
    $CheckBoxOptionLabel.Name = "CheckBoxOptionLabel"
    $CheckBoxOptionLabel.Size = [System.Drawing.Size]::New(($CheckBoxOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
    $CheckBoxOptionLabel.Text = $Message
    $CheckBoxOptionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #endregion $CheckBoxOptionLabel = [System.Windows.Forms.Label]::New()

    # Returns the minimum size required to display the text
    $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($CheckBoxOptionLabel.Text, [MyConfig]::Font.Regular, $CheckBoxOptionLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
    $CheckBoxOptionLabel.Size = [System.Drawing.Size]::New(($CheckBoxOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

    $TempBottom = $CheckBoxOptionLabel.Bottom + [MyConfig]::FormSpacer
  }
  Else
  {
    $TempBottom = [MyConfig]::FormSpacer
  }

  # ************************************************
  # CheckBoxOption GroupBox
  # ************************************************
  #region $CheckBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  $CheckBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $CheckBoxOptionPanel.Controls.Add($CheckBoxOptionGroupBox)
  $CheckBoxOptionGroupBox.BackColor = [MyConfig]::Colors.Back
  $CheckBoxOptionGroupBox.Font = [MyConfig]::Font.Regular
  $CheckBoxOptionGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $CheckBoxOptionGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($TempBottom + [MyConfig]::FormSpacer))
  $CheckBoxOptionGroupBox.Name = "CheckBoxOptionGroupBox"
  $CheckBoxOptionGroupBox.Size = [System.Drawing.Size]::New(($CheckBoxOptionPanel.Width - ([MyConfig]::FormSpacer * 2)), 23)
  $CheckBoxOptionGroupBox.Text = $Null
  #endregion $CheckBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $CheckBoxOptionGroupBox Controls ********

  $CheckBoxOptionNumber = 0
  $GroupBottom = [MyConfig]::Font.Height
  ForEach ($Key In $OrderedItems.Keys)
  {
    #region $CheckBoxOptionCheckBox = [System.Windows.Forms.CheckBox]::New()
    $CheckBoxOptionCheckBox = [System.Windows.Forms.CheckBox]::New()
    $CheckBoxOptionGroupBox.Controls.Add($CheckBoxOptionCheckBox)
    #$CheckBoxOptionCheckBox.AutoCheck = $True
    $CheckBoxOptionCheckBox.AutoSize = $True
    $CheckBoxOptionCheckBox.BackColor = [MyConfig]::Colors.Back
    $CheckBoxOptionCheckBox.Checked = ($OrderedItems[$Key] -in $Selected)
    $CheckBoxOptionCheckBox.Font = [MyConfig]::Font.Regular
    $CheckBoxOptionCheckBox.ForeColor = [MyConfig]::Colors.LabelFore
    $CheckBoxOptionCheckBox.Location = [System.Drawing.Point]::New(([MyConfig]::FormSpacer * [MyConfig]::FormSpacer), $GroupBottom)
    $CheckBoxOptionCheckBox.Name = "CheckBox$($CheckBoxOptionNumber)"
    $CheckBoxOptionCheckBox.Tag = $OrderedItems[$Key]
    $CheckBoxOptionCheckBox.Text = $Key
    #endregion $CheckBoxOptionCheckBox = [System.Windows.Forms.CheckBox]::New()

    $GroupBottom = ($CheckBoxOptionCheckBox.Bottom + [MyConfig]::FormSpacer)
    $CheckBoxOptionNumber += 1
  }

  $CheckBoxOptionGroupBox.ClientSize = [System.Drawing.Size]::New($CheckBoxOptionGroupBox.ClientSize.Width, ($GroupBottom + [MyConfig]::FormSpacer))

  #endregion ******** $CheckBoxOptionGroupBox Controls ********

  #endregion ******** $CheckBoxOptionPanel Controls ********

  # ************************************************
  # CheckBoxOptionBtm Panel
  # ************************************************
  #region $CheckBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $CheckBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $CheckBoxOptionForm.Controls.Add($CheckBoxOptionBtmPanel)
  $CheckBoxOptionBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $CheckBoxOptionBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $CheckBoxOptionBtmPanel.Name = "CheckBoxOptionBtmPanel"
  #endregion $CheckBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $CheckBoxOptionBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($CheckBoxOptionBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $CheckBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $CheckBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $CheckBoxOptionBtmPanel.Controls.Add($CheckBoxOptionBtmLeftButton)
  $CheckBoxOptionBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $CheckBoxOptionBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $CheckBoxOptionBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $CheckBoxOptionBtmLeftButton.Font = [MyConfig]::Font.Bold
  $CheckBoxOptionBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $CheckBoxOptionBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $CheckBoxOptionBtmLeftButton.Name = "CheckBoxOptionBtmLeftButton"
  $CheckBoxOptionBtmLeftButton.TabIndex = 1
  $CheckBoxOptionBtmLeftButton.TabStop = $True
  $CheckBoxOptionBtmLeftButton.Text = $ButtonLeft
  $CheckBoxOptionBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $CheckBoxOptionBtmLeftButton.PreferredSize.Height)
  #endregion $CheckBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-CheckBoxOptionBtmLeftButtonClick ********
  Function Start-CheckBoxOptionBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckBoxOptionBtmLeft Button Control
      .DESCRIPTION
        Click Event for the CheckBoxOptionBtmLeft Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-CheckBoxOptionBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By CDUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$CheckBoxOptionBtmLeftButton"

    [MyConfig]::AutoExit = 0

    if ($Required.IsPresent)
    {
      If (@($CheckBoxOptionGroupBox.Controls | Where-Object -FilterScript { ($PSItem.GetType().Name -eq "CheckBox") -and $PSItem.Checked }).Count -gt 0)
      {
        $CheckBoxOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
      }
      Else
      {
        [Void][System.Windows.Forms.MessageBox]::Show($CheckBoxOptionForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
      }
    }
    else
    {
      $CheckBoxOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }

    Write-Verbose -Message "Exit Click Event for `$CheckBoxOptionBtmLeftButton"
  }
  #endregion ******** Function Start-CheckBoxOptionBtmLeftButtonClick ********
  $CheckBoxOptionBtmLeftButton.add_Click({ Start-CheckBoxOptionBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $CheckBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $CheckBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $CheckBoxOptionBtmPanel.Controls.Add($CheckBoxOptionBtmMidButton)
  $CheckBoxOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $CheckBoxOptionBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $CheckBoxOptionBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $CheckBoxOptionBtmMidButton.Font = [MyConfig]::Font.Bold
  $CheckBoxOptionBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $CheckBoxOptionBtmMidButton.Location = [System.Drawing.Point]::New(($CheckBoxOptionBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $CheckBoxOptionBtmMidButton.Name = "CheckBoxOptionBtmMidButton"
  $CheckBoxOptionBtmMidButton.TabIndex = 2
  $CheckBoxOptionBtmMidButton.TabStop = $True
  $CheckBoxOptionBtmMidButton.Text = $ButtonMid
  $CheckBoxOptionBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $CheckBoxOptionBtmMidButton.PreferredSize.Height)
  #endregion $CheckBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-CheckBoxOptionBtmMidButtonClick ********
  Function Start-CheckBoxOptionBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckBoxOptionBtmMid Button Control
      .DESCRIPTION
        Click Event for the CheckBoxOptionBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-CheckBoxOptionBtmMidButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By CDUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$CheckBoxOptionBtmMidButton"

    [MyConfig]::AutoExit = 0

    ForEach ($CheckBox In @($CheckBoxOptionGroupBox.Controls | Where-Object -FilterScript { $PSItem.Name -Like "RadioChoice*" }))
    {
      $CheckBox.Checked = ($CheckBox.Tag -in $Selected)
    }

    Write-Verbose -Message "Exit Click Event for `$CheckBoxOptionBtmMidButton"
  }
  #endregion ******** Function Start-CheckBoxOptionBtmMidButtonClick ********
  $CheckBoxOptionBtmMidButton.add_Click({ Start-CheckBoxOptionBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $CheckBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $CheckBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $CheckBoxOptionBtmPanel.Controls.Add($CheckBoxOptionBtmRightButton)
  $CheckBoxOptionBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $CheckBoxOptionBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $CheckBoxOptionBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $CheckBoxOptionBtmRightButton.Font = [MyConfig]::Font.Bold
  $CheckBoxOptionBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $CheckBoxOptionBtmRightButton.Location = [System.Drawing.Point]::New(($CheckBoxOptionBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $CheckBoxOptionBtmRightButton.Name = "CheckBoxOptionBtmRightButton"
  $CheckBoxOptionBtmRightButton.TabIndex = 3
  $CheckBoxOptionBtmRightButton.TabStop = $True
  $CheckBoxOptionBtmRightButton.Text = $ButtonRight
  $CheckBoxOptionBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $CheckBoxOptionBtmRightButton.PreferredSize.Height)
  #endregion $CheckBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-CheckBoxOptionBtmRightButtonClick ********
  Function Start-CheckBoxOptionBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckBoxOptionBtmRight Button Control
      .DESCRIPTION
        Click Event for the CheckBoxOptionBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-CheckBoxOptionBtmRightButtonClick -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By CDUserName)
    #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Button]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$CheckBoxOptionBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here

    $CheckBoxOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$CheckBoxOptionBtmRightButton"
  }
  #endregion ******** Function Start-CheckBoxOptionBtmRightButtonClick ********
  $CheckBoxOptionBtmRightButton.add_Click({ Start-CheckBoxOptionBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $CheckBoxOptionBtmPanel.ClientSize = [System.Drawing.Size]::New(($CheckBoxOptionBtmRightButton.Right + [MyConfig]::FormSpacer), ($CheckBoxOptionBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $CheckBoxOptionBtmPanel Controls ********

  $CheckBoxOptionForm.ClientSize = [System.Drawing.Size]::New($CheckBoxOptionForm.ClientSize.Width, ($CheckBoxOptionGroupBox.Bottom + [MyConfig]::FormSpacer + $CheckBoxOptionBtmPanel.Height))

  #endregion ******** Controls for CheckBoxOption Form ********

  #endregion ******** End **** CheckBoxOption **** End ********

  $DialogResult = $CheckBoxOptionForm.ShowDialog($PILForm)
  If ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
  {
    $TempItems = @{}
    $CheckBoxOptionGroupBox.Controls | Where-Object -FilterScript { $PSItem.Name -Like "CheckBox*" -and $PSItem.Checked } | ForEach-Object -Process { $TempItems.Add($PSItem.Text, $PSItem.Tag) }
    [CheckBoxOption]::New($True, $DialogResult, $TempItems)
  }
  Else
  {
    [CheckBoxOption]::New($False, $DialogResult)
  }

  $CheckBoxOptionForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-CheckBoxOption"
}
#endregion function Get-CheckBoxOption

# ------------------------------
# Get ListBoxOption Function
# ------------------------------
#region ListBoxOption Result Class
Class ListBoxOption
{
  [Bool]$Success
  [Object]$DialogResult
  [Object[]]$Items

  ListBoxOption ([Bool]$Success, [Object]$DialogResult, [Object[]]$Items)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.Items = $Items
  }
}
#endregion ListBoxOption Result Class

#region function Get-ListBoxOption
function Get-ListBoxOption ()
{
  <#
    .SYNOPSIS
      Shows Get-ListBoxOption
    .DESCRIPTION
      Shows Get-ListBoxOption
    .PARAMETER Title
      Title of the Get-ListBoxOption Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER Items
      Items to show in the ListView
    .PARAMETER Sorted
      Sort ListView
    .PARAMETER Multi
      Allow Selecting Multiple Items
    .PARAMETER DisplayMember
      Name of the Property to Display in the ListBox
    .PARAMETER ValueMember
      Name of the Property for the Value
    .PARAMETER Selected
      Default Selected ListBox Items
    .PARAMETER Width
      Width of Get-ListBoxOption Dialog Window
    .PARAMETER Height
      Height of Get-ListBoxOption Dialog Window
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Middle Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $Items = Get-Service
      $DialogResult = Get-ListBoxOption -Title "Get ListBox Option" -Message "Show this Sample Message Prompt to the User" -DisplayMember "DisplayName" -ValueMember "Name" -Items $Items -Selected $Items[1, 3, 5, 7] -Multi
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
    [Switch]$Multi,
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
  Write-Verbose -Message "Enter Function Get-ListBoxOption"

  #region ******** Begin **** ListBoxOption **** Begin ********

  # ************************************************
  # ListBoxOption Form
  # ************************************************
  #region $ListBoxOptionForm = [System.Windows.Forms.Form]::New()
  $ListBoxOptionForm = [System.Windows.Forms.Form]::New()
  $ListBoxOptionForm.BackColor = [MyConfig]::Colors.Back
  $ListBoxOptionForm.Font = [MyConfig]::Font.Regular
  $ListBoxOptionForm.ForeColor = [MyConfig]::Colors.Fore
  $ListBoxOptionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $ListBoxOptionForm.Icon = $PILForm.Icon
  $ListBoxOptionForm.KeyPreview = $True
  $ListBoxOptionForm.MaximizeBox = $False
  $ListBoxOptionForm.MinimizeBox = $False
  $ListBoxOptionForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * $Height))
  $ListBoxOptionForm.Name = "ListBoxOptionForm"
  $ListBoxOptionForm.Owner = $PILForm
  $ListBoxOptionForm.ShowInTaskbar = $False
  $ListBoxOptionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $ListBoxOptionForm.Text = $Title
  #endregion $ListBoxOptionForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-ListBoxOptionFormKeyDown ********
  function Start-ListBoxOptionFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the ListBoxOption Form Control
      .DESCRIPTION
        KeyDown Event for the ListBoxOption Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-ListBoxOptionFormKeyDown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter KeyDown Event for `$ListBoxOptionForm"

    [MyConfig]::AutoExit = 0

    if ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $ListBoxOptionForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$ListBoxOptionForm"
  }
  #endregion ******** Function Start-ListBoxOptionFormKeyDown ********
  $ListBoxOptionForm.add_KeyDown({ Start-ListBoxOptionFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ListBoxOptionFormShown ********
  function Start-ListBoxOptionFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the ListBoxOption Form Control
      .DESCRIPTION
        Shown Event for the ListBoxOption Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-ListBoxOptionFormShown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Shown Event for `$ListBoxOptionForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    Write-Verbose -Message "Exit Shown Event for `$ListBoxOptionForm"
  }
  #endregion ******** Function Start-ListBoxOptionFormShown ********
  $ListBoxOptionForm.add_Shown({ Start-ListBoxOptionFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for ListBoxOption Form ********

  # ************************************************
  # ListBoxOption Panel
  # ************************************************
  #region $ListBoxOptionPanel = [System.Windows.Forms.Panel]::New()
  $ListBoxOptionPanel = [System.Windows.Forms.Panel]::New()
  $ListBoxOptionForm.Controls.Add($ListBoxOptionPanel)
  $ListBoxOptionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ListBoxOptionPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $ListBoxOptionPanel.Name = "ListBoxOptionPanel"
  #endregion $ListBoxOptionPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ListBoxOptionPanel Controls ********

  if ($PSBoundParameters.ContainsKey("Message"))
  {
    #region $ListBoxOptionLabel = [System.Windows.Forms.Label]::New()
    $ListBoxOptionLabel = [System.Windows.Forms.Label]::New()
    $ListBoxOptionPanel.Controls.Add($ListBoxOptionLabel)
    $ListBoxOptionLabel.ForeColor = [MyConfig]::Colors.LabelFore
    $ListBoxOptionLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
    $ListBoxOptionLabel.Name = "ListBoxOptionLabel"
    $ListBoxOptionLabel.Size = [System.Drawing.Size]::New(($ListBoxOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
    $ListBoxOptionLabel.Text = $Message
    $ListBoxOptionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #endregion $ListBoxOptionLabel = [System.Windows.Forms.Label]::New()

    # Returns the minimum size required to display the text
    $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($ListBoxOptionLabel.Text, [MyConfig]::Font.Regular, $ListBoxOptionLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
    $ListBoxOptionLabel.Size = [System.Drawing.Size]::New(($ListBoxOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

    $TmpBottom = $ListBoxOptionLabel.Bottom + [MyConfig]::FormSpacer
  }
  else
  {
    $TmpBottom = 0
  }

  # ************************************************
  # ListBoxOption GroupBox
  # ************************************************
  #region $ListBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  $ListBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $ListBoxOptionPanel.Controls.Add($ListBoxOptionGroupBox)
  $ListBoxOptionGroupBox.BackColor = [MyConfig]::Colors.Back
  $ListBoxOptionGroupBox.Font = [MyConfig]::Font.Regular
  $ListBoxOptionGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $ListBoxOptionGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($TmpBottom + [MyConfig]::FormSpacer))
  $ListBoxOptionGroupBox.Name = "ListBoxOptionGroupBox"
  $ListBoxOptionGroupBox.Size = [System.Drawing.Size]::New(($ListBoxOptionPanel.Width - ([MyConfig]::FormSpacer * 2)), ($ListBoxOptionPanel.Height - ($ListBoxOptionGroupBox.Top + [MyConfig]::FormSpacer)))
  #endregion $ListBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $ListBoxOptionGroupBox Controls ********

  #region $ListBoxOptionListBox = [System.Windows.Forms.ListBox]::New()
  $ListBoxOptionListBox = [System.Windows.Forms.ListBox]::New()
  $ListBoxOptionGroupBox.Controls.Add($ListBoxOptionListBox)
  $ListBoxOptionListBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
  $ListBoxOptionListBox.AutoSize = $True
  $ListBoxOptionListBox.BackColor = [MyConfig]::Colors.TextBack
  $ListBoxOptionListBox.DisplayMember = $DisplayMember
  $ListBoxOptionListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $ListBoxOptionListBox.Font = [MyConfig]::Font.Regular
  $ListBoxOptionListBox.ForeColor = [MyConfig]::Colors.TextFore
  $ListBoxOptionListBox.Name = "ListBoxOptionListBox"
  if ($Multi.IsPresent)
  {
    $ListBoxOptionListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
  }
  else
  {
    $ListBoxOptionListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::One
  }
  $ListBoxOptionListBox.Sorted = $Sorted.IsPresent
  $ListBoxOptionListBox.TabIndex = 0
  $ListBoxOptionListBox.TabStop = $True
  $ListBoxOptionListBox.Tag = $Null
  $ListBoxOptionListBox.ValueMember = $ValueMember
  #endregion $ListBoxOptionListBox = [System.Windows.Forms.ListBox]::New()

  $ListBoxOptionListBox.Items.AddRange($Items)
  if ($PSBoundParameters.ContainsKey("Selected"))
  {
    if ($Multi.IsPresent)
    {
      $ListBoxOptionListBox.Tag = @($Items | Where-Object -FilterScript { $PSItem -in $Selected} )
    }
    else
    {
      $ListBoxOptionListBox.Tag = @($Items | Select-Object -First 1 )
    }
    $ListBoxOptionListBox.SelectedItems.Clear()
    $ListBoxOptionListBox.Tag | ForEach-Object -Process { $ListBoxOptionListBox.SelectedItems.Add($PSItem) }
  }
  else
  {
    $ListBoxOptionListBox.Tag = @()
  }
  
  #region ******** Function Start-ListBoxOptionListBoxMouseDown ********
  function Start-ListBoxOptionListBoxMouseDown
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
         Start-ListBoxOptionListBoxMouseDown -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By ken.sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ListBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter MouseDown Event for `$ListBoxOptionListBox"

    [MyConfig]::AutoExit = 0

    If ($EventArg.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
      if ($ListBoxOptionListBox.Items.Count -gt 0)
      {
        $ListBoxOptionContextMenuStrip.Show($ListBoxOptionListBox, $EventArg.Location)
      }
    }

    Write-Verbose -Message "Exit MouseDown Event for `$ListBoxOptionListBox"
  }
  #endregion ******** Function Start-ListBoxOptionListBoxMouseDown ********
  if ($Multi.IsPresent)
  {
    $ListBoxOptionListBox.add_MouseDown({ Start-ListBoxOptionListBoxMouseDown -Sender $This -EventArg $PSItem })
  }
  
  $ListBoxOptionGroupBox.ClientSize = [System.Drawing.Size]::New($ListBoxOptionGroupBox.ClientSize.Width, ($ListBoxOptionListBox.Bottom + ([MyConfig]::FormSpacer * 2)))
  
  # ************************************************
  # ListBoxOption ContextMenuStrip
  # ************************************************
  #region $ListBoxOptionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  $ListBoxOptionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  #$ListBoxOptionListView.Controls.Add($ListBoxOptionContextMenuStrip)
  $ListBoxOptionContextMenuStrip.BackColor = [MyConfig]::Colors.Back
  #$ListBoxOptionContextMenuStrip.Enabled = $True
  $ListBoxOptionContextMenuStrip.Font = [MyConfig]::Font.Regular
  $ListBoxOptionContextMenuStrip.ForeColor = [MyConfig]::Colors.Fore
  $ListBoxOptionContextMenuStrip.ImageList = $PILSmallImageList
  $ListBoxOptionContextMenuStrip.Name = "ListBoxOptionContextMenuStrip"
  #endregion $ListBoxOptionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()

  #region ******** Function Start-ListBoxOptionContextMenuStripOpening ********
  function Start-ListBoxOptionContextMenuStripOpening
  {
    <#
      .SYNOPSIS
        Opening Event for the ListBoxOption ContextMenuStrip Control
      .DESCRIPTION
        Opening Event for the ListBoxOption ContextMenuStrip Control
      .PARAMETER Sender
         The ContextMenuStrip Control that fired the Opening Event
      .PARAMETER EventArg
         The Event Arguments for the ContextMenuStrip Opening Event
      .EXAMPLE
         Start-ListBoxOptionContextMenuStripOpening -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Opening Event for `$ListBoxOptionContextMenuStrip"

    [MyConfig]::AutoExit = 0

    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

    Write-Verbose -Message "Exit Opening Event for `$ListBoxOptionContextMenuStrip"
  }
  #endregion ******** Function Start-ListBoxOptionContextMenuStripOpening ********
  $ListBoxOptionContextMenuStrip.add_Opening({Start-ListBoxOptionContextMenuStripOpening -Sender $This -EventArg $PSItem})

  #region ******** Function Start-ListBoxOptionContextMenuStripItemClick ********
  function Start-ListBoxOptionContextMenuStripItemClick
  {
    <#
      .SYNOPSIS
        Click Event for the ListBoxOption ToolStripItem Control
      .DESCRIPTION
        Click Event for the ListBoxOption ToolStripItem Control
      .PARAMETER Sender
         The ToolStripItem Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the ToolStripItem Click Event
      .EXAMPLE
         Start-ListBoxOptionContextMenuStripItemClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$ListBoxOptionContextMenuStripItem"

    [MyConfig]::AutoExit = 0
    
    switch ($Sender.Name)
    {
      "SelectAll"
      {
        @($ListBoxOptionListBox.Items) | ForEach-Object -Process { $ListBoxOptionListBox.SelectedItems.Add($PSItem) }
        Break
      }
      "UnSelectAll"
      {
        $ListBoxOptionListBox.SelectedItems.Clear()
        Break
      }
    }

    Write-Verbose -Message "Exit Click Event for `$ListBoxOptionContextMenuStripItem"
  }
  #endregion ******** Function Start-ListBoxOptionContextMenuStripItemClick ********

  (New-MenuItem -Menu $ListBoxOptionContextMenuStrip -Text "Select All" -Name "SelectAll" -Tag "SelectAll" -DisplayStyle "ImageAndText" -ImageKey "CheckIcon" -PassThru).add_Click({Start-ListBoxOptionContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $ListBoxOptionContextMenuStrip -Text "Unselect All" -Name "UnSelectAll" -Tag "UnSelectAll" -DisplayStyle "ImageAndText" -ImageKey "UncheckIcon" -PassThru).add_Click({Start-ListBoxOptionContextMenuStripItemClick -Sender $This -EventArg $PSItem})

  #endregion ******** $ListBoxOptionGroupBox Controls ********

  $TempClientSize = [System.Drawing.Size]::New(($ListBoxOptionGroupBox.Right + [MyConfig]::FormSpacer), ($ListBoxOptionGroupBox.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ListBoxOptionPanel Controls ********

  # ************************************************
  # ListBoxOptionBtm Panel
  # ************************************************
  #region $ListBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $ListBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $ListBoxOptionForm.Controls.Add($ListBoxOptionBtmPanel)
  $ListBoxOptionBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ListBoxOptionBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $ListBoxOptionBtmPanel.Name = "ListBoxOptionBtmPanel"
  #endregion $ListBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ListBoxOptionBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($ListBoxOptionBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $ListBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ListBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ListBoxOptionBtmPanel.Controls.Add($ListBoxOptionBtmLeftButton)
  $ListBoxOptionBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $ListBoxOptionBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ListBoxOptionBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ListBoxOptionBtmLeftButton.Font = [MyConfig]::Font.Bold
  $ListBoxOptionBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ListBoxOptionBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $ListBoxOptionBtmLeftButton.Name = "ListBoxOptionBtmLeftButton"
  $ListBoxOptionBtmLeftButton.TabIndex = 1
  $ListBoxOptionBtmLeftButton.TabStop = $True
  $ListBoxOptionBtmLeftButton.Text = $ButtonLeft
  $ListBoxOptionBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $ListBoxOptionBtmLeftButton.PreferredSize.Height)
  #endregion $ListBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ListBoxOptionBtmLeftButtonClick ********
  function Start-ListBoxOptionBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ListBoxOptionBtmLeft Button Control
      .DESCRIPTION
        Click Event for the ListBoxOptionBtmLeft Button Control
      .PARAMETER Sender
         The Button Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the Button Click Event
      .EXAMPLE
         Start-ListBoxOptionBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$ListBoxOptionBtmLeftButton"

    [MyConfig]::AutoExit = 0

    if ($ListBoxOptionListBox.SelectedIndex -gt 0)
    {
      $ListBoxOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    else
    {
      [Void][System.Windows.Forms.MessageBox]::Show($ListBoxOptionForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }

    Write-Verbose -Message "Exit Click Event for `$ListBoxOptionBtmLeftButton"
  }
  #endregion ******** Function Start-ListBoxOptionBtmLeftButtonClick ********
  $ListBoxOptionBtmLeftButton.add_Click({ Start-ListBoxOptionBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $ListBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $ListBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $ListBoxOptionBtmPanel.Controls.Add($ListBoxOptionBtmMidButton)
  $ListBoxOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $ListBoxOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $ListBoxOptionBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ListBoxOptionBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ListBoxOptionBtmMidButton.Font = [MyConfig]::Font.Bold
  $ListBoxOptionBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ListBoxOptionBtmMidButton.Location = [System.Drawing.Point]::New(($ListBoxOptionBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ListBoxOptionBtmMidButton.Name = "ListBoxOptionBtmMidButton"
  $ListBoxOptionBtmMidButton.TabIndex = 2
  $ListBoxOptionBtmMidButton.TabStop = $True
  $ListBoxOptionBtmMidButton.Text = $ButtonMid
  $ListBoxOptionBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $ListBoxOptionBtmMidButton.PreferredSize.Height)
  #endregion $ListBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ListBoxOptionBtmMidButtonClick ********
  function Start-ListBoxOptionBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ListBoxOptionBtmMid Button Control
      .DESCRIPTION
        Click Event for the ListBoxOptionBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ListBoxOptionBtmMidButtonClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$ListBoxOptionBtmMidButton"

    [MyConfig]::AutoExit = 0

    $ListBoxOptionListBox.SelectedItems.Clear()
    if ($ListBoxOptionListBox.Tag.Count -gt 0)
    {
      $ListBoxOptionListBox.Tag | ForEach-Object -Process { $ListBoxOptionListBox.SelectedItems.Add($PSItem) }
    }

    Write-Verbose -Message "Exit Click Event for `$ListBoxOptionBtmMidButton"
  }
  #endregion ******** Function Start-ListBoxOptionBtmMidButtonClick ********
  $ListBoxOptionBtmMidButton.add_Click({ Start-ListBoxOptionBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $ListBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $ListBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $ListBoxOptionBtmPanel.Controls.Add($ListBoxOptionBtmRightButton)
  $ListBoxOptionBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $ListBoxOptionBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ListBoxOptionBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ListBoxOptionBtmRightButton.Font = [MyConfig]::Font.Bold
  $ListBoxOptionBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ListBoxOptionBtmRightButton.Location = [System.Drawing.Point]::New(($ListBoxOptionBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ListBoxOptionBtmRightButton.Name = "ListBoxOptionBtmRightButton"
  $ListBoxOptionBtmRightButton.TabIndex = 3
  $ListBoxOptionBtmRightButton.TabStop = $True
  $ListBoxOptionBtmRightButton.Text = $ButtonRight
  $ListBoxOptionBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $ListBoxOptionBtmRightButton.PreferredSize.Height)
  #endregion $ListBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ListBoxOptionBtmRightButtonClick ********
  function Start-ListBoxOptionBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ListBoxOptionBtmRight Button Control
      .DESCRIPTION
        Click Event for the ListBoxOptionBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ListBoxOptionBtmRightButtonClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$ListBoxOptionBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here

    $ListBoxOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$ListBoxOptionBtmRightButton"
  }
  #endregion ******** Function Start-ListBoxOptionBtmRightButtonClick ********
  $ListBoxOptionBtmRightButton.add_Click({ Start-ListBoxOptionBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $ListBoxOptionBtmPanel.ClientSize = [System.Drawing.Size]::New(($ListBoxOptionBtmRightButton.Right + [MyConfig]::FormSpacer), ($ListBoxOptionBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ListBoxOptionBtmPanel Controls ********

  $ListBoxOptionForm.ClientSize = [System.Drawing.Size]::New($ListBoxOptionForm.ClientSize.Width, ($TempClientSize.Height + $ListBoxOptionBtmPanel.Height))

  #endregion ******** Controls for ListBoxOption Form ********

  #endregion ******** End **** ListBoxOption **** End ********

  $DialogResult = $ListBoxOptionForm.ShowDialog()
  if ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
  {
    [ListBoxOption]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, $ListBoxOptionListBox.SelectedItems)
  }
  else
  {
    [ListBoxOption]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, @())
  }

  $ListBoxOptionForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-ListBoxOption"
}
#endregion function Get-ListBoxOption

# --------------------------------
# Get CheckedListBoxOption Function
# --------------------------------
#region CheckedListBoxOption Result Class
Class CheckedListBoxOption
{
  [Bool]$Success
  [Object]$DialogResult
  [Object[]]$Items

  CheckedListBoxOption ([Bool]$Success, [Object]$DialogResult, [Object[]]$Items)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.Items = $Items
  }
}
#endregion CheckedListBoxOption Result Class

#region function Get-CheckedListBoxOption
function Get-CheckedListBoxOption ()
{
  <#
    .SYNOPSIS
      Shows Get-CheckedListBoxOption
    .DESCRIPTION
      Shows Get-CheckedListBoxOption
    .PARAMETER Title
      Title of the Get-CheckedListBoxOption Dialog Window
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
      Width of Get-CheckedListBoxOption Dialog Window
    .PARAMETER Height
      Height of Get-CheckedListBoxOption Dialog Window
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $Items = Get-Service
      $DialogResult = CheckedGet-CheckedListBoxOption -Title "Get CheckListBox Option" -Message "Show this Sample Message Prompt to the User" -DisplayMember "DisplayName" -ValueMember "Name" -Items $Items -Selected $Items[1, 3, 5, 7]
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
  Write-Verbose -Message "Enter Function Get-CheckedListBoxOption"

  #region ******** Begin **** CheckedListBoxOption **** Begin ********

  # ************************************************
  # CheckedListBoxOption Form
  # ************************************************
  #region $CheckedListBoxOptionForm = [System.Windows.Forms.Form]::New()
  $CheckedListBoxOptionForm = [System.Windows.Forms.Form]::New()
  $CheckedListBoxOptionForm.BackColor = [MyConfig]::Colors.Back
  $CheckedListBoxOptionForm.Font = [MyConfig]::Font.Regular
  $CheckedListBoxOptionForm.ForeColor = [MyConfig]::Colors.Fore
  $CheckedListBoxOptionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $CheckedListBoxOptionForm.Icon = $PILForm.Icon
  $CheckedListBoxOptionForm.KeyPreview = $True
  $CheckedListBoxOptionForm.MaximizeBox = $False
  $CheckedListBoxOptionForm.MinimizeBox = $False
  $CheckedListBoxOptionForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * $Height))
  $CheckedListBoxOptionForm.Name = "CheckedListBoxOptionForm"
  $CheckedListBoxOptionForm.Owner = $PILForm
  $CheckedListBoxOptionForm.ShowInTaskbar = $False
  $CheckedListBoxOptionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $CheckedListBoxOptionForm.Text = $Title
  #endregion $CheckedListBoxOptionForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-CheckedListBoxOptionFormKeyDown ********
  function Start-CheckedListBoxOptionFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the CheckedListBoxOption Form Control
      .DESCRIPTION
        KeyDown Event for the CheckedListBoxOption Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-CheckedListBoxOptionFormKeyDown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter KeyDown Event for `$CheckedListBoxOptionForm"

    [MyConfig]::AutoExit = 0

    if ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $CheckedListBoxOptionForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$CheckedListBoxOptionForm"
  }
  #endregion ******** Function Start-CheckedListBoxOptionFormKeyDown ********
  $CheckedListBoxOptionForm.add_KeyDown({ Start-CheckedListBoxOptionFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-CheckedListBoxOptionFormShown ********
  function Start-CheckedListBoxOptionFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the CheckedListBoxOption Form Control
      .DESCRIPTION
        Shown Event for the CheckedListBoxOption Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-CheckedListBoxOptionFormShown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Shown Event for `$CheckedListBoxOptionForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    Write-Verbose -Message "Exit Shown Event for `$CheckedListBoxOptionForm"
  }
  #endregion ******** Function Start-CheckedListBoxOptionFormShown ********
  $CheckedListBoxOptionForm.add_Shown({ Start-CheckedListBoxOptionFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for CheckedListBoxOption Form ********

  # ************************************************
  # CheckedListBoxOption Panel
  # ************************************************
  #region $CheckedListBoxOptionPanel = [System.Windows.Forms.Panel]::New()
  $CheckedListBoxOptionPanel = [System.Windows.Forms.Panel]::New()
  $CheckedListBoxOptionForm.Controls.Add($CheckedListBoxOptionPanel)
  $CheckedListBoxOptionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $CheckedListBoxOptionPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $CheckedListBoxOptionPanel.Name = "CheckedListBoxOptionPanel"
  #endregion $CheckedListBoxOptionPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $CheckedListBoxOptionPanel Controls ********

  if ($PSBoundParameters.ContainsKey("Message"))
  {
    #region $CheckedListBoxOptionLabel = [System.Windows.Forms.Label]::New()
    $CheckedListBoxOptionLabel = [System.Windows.Forms.Label]::New()
    $CheckedListBoxOptionPanel.Controls.Add($CheckedListBoxOptionLabel)
    $CheckedListBoxOptionLabel.ForeColor = [MyConfig]::Colors.LabelFore
    $CheckedListBoxOptionLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
    $CheckedListBoxOptionLabel.Name = "CheckedListBoxOptionLabel"
    $CheckedListBoxOptionLabel.Size = [System.Drawing.Size]::New(($CheckedListBoxOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
    $CheckedListBoxOptionLabel.Text = $Message
    $CheckedListBoxOptionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #endregion $CheckedListBoxOptionLabel = [System.Windows.Forms.Label]::New()

    # Returns the minimum size required to display the text
    $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($CheckedListBoxOptionLabel.Text, [MyConfig]::Font.Regular, $CheckedListBoxOptionLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
    $CheckedListBoxOptionLabel.Size = [System.Drawing.Size]::New(($CheckedListBoxOptionPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

    $TmpBottom = $CheckedListBoxOptionLabel.Bottom + [MyConfig]::FormSpacer
  }
  else
  {
    $TmpBottom = 0
  }

  # ************************************************
  # CheckedListBoxOption GroupBox
  # ************************************************
  #region $CheckedListBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  $CheckedListBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $CheckedListBoxOptionPanel.Controls.Add($CheckedListBoxOptionGroupBox)
  $CheckedListBoxOptionGroupBox.BackColor = [MyConfig]::Colors.Back
  $CheckedListBoxOptionGroupBox.Font = [MyConfig]::Font.Regular
  $CheckedListBoxOptionGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $CheckedListBoxOptionGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($TmpBottom + [MyConfig]::FormSpacer))
  $CheckedListBoxOptionGroupBox.Name = "CheckedListBoxOptionGroupBox"
  $CheckedListBoxOptionGroupBox.Size = [System.Drawing.Size]::New(($CheckedListBoxOptionPanel.Width - ([MyConfig]::FormSpacer * 2)), ($CheckedListBoxOptionPanel.Height - ($CheckedListBoxOptionGroupBox.Top + [MyConfig]::FormSpacer)))
  #endregion $CheckedListBoxOptionGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $CheckedListBoxOptionGroupBox Controls ********

  #region $CheckedListBoxOptionCheckedListBox = [System.Windows.Forms.CheckedListBox]::New()
  $CheckedListBoxOptionCheckedListBox = [System.Windows.Forms.CheckedListBox]::New()
  $CheckedListBoxOptionGroupBox.Controls.Add($CheckedListBoxOptionCheckedListBox)
  $CheckedListBoxOptionCheckedListBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
  $CheckedListBoxOptionCheckedListBox.AutoSize = $True
  $CheckedListBoxOptionCheckedListBox.BackColor = [MyConfig]::Colors.TextBack
  $CheckedListBoxOptionCheckedListBox.CheckOnClick = $True
  $CheckedListBoxOptionCheckedListBox.DisplayMember = $DisplayMember
  $CheckedListBoxOptionCheckedListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $CheckedListBoxOptionCheckedListBox.Font = [MyConfig]::Font.Regular
  $CheckedListBoxOptionCheckedListBox.ForeColor = [MyConfig]::Colors.TextFore
  $CheckedListBoxOptionCheckedListBox.Name = "CheckedListBoxOptionCheckedListBox"
  $CheckedListBoxOptionCheckedListBox.Sorted = $Sorted.IsPresent
  $CheckedListBoxOptionCheckedListBox.TabIndex = 0
  $CheckedListBoxOptionCheckedListBox.TabStop = $True
  $CheckedListBoxOptionCheckedListBox.Tag = $Null
  $CheckedListBoxOptionCheckedListBox.ValueMember = $ValueMember
  #endregion $CheckedListBoxOptionCheckedListBox = [System.Windows.Forms.CheckedListBox]::New()

  $CheckedListBoxOptionCheckedListBox.Items.AddRange($Items)

  if ($PSBoundParameters.ContainsKey("Selected"))
  {
    $CheckedListBoxOptionCheckedListBox.Tag = @($Items | Where-Object -FilterScript { $PSItem -in $Selected})
    $CheckedListBoxOptionCheckedListBox.Tag | ForEach-Object -Process { $CheckedListBoxOptionCheckedListBox.SetItemChecked($CheckedListBoxOptionCheckedListBox.Items.IndexOf($PSItem), $True) }
  }
  else
  {
    $CheckedListBoxOptionCheckedListBox.Tag = @()
  }

  #region ******** Function Start-CheckedListBoxOptionCheckedListBoxMouseDown ********
  function Start-CheckedListBoxOptionCheckedListBoxMouseDown
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
         Start-CheckedListBoxOptionCheckedListBoxMouseDown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter MouseDown Event for `$CheckedListBoxOptionCheckedListBox"

    [MyConfig]::AutoExit = 0

    If ($EventArg.Button -eq [System.Windows.Forms.MouseButtons]::Right)
    {
      if ($CheckedListBoxOptionCheckedListBox.Items.Count -gt 0)
      {
        $CheckedListBoxOptionContextMenuStrip.Show($CheckedListBoxOptionCheckedListBox, $EventArg.Location)
      }
    }

    Write-Verbose -Message "Exit MouseDown Event for `$CheckedListBoxOptionCheckedListBox"
  }
  #endregion ******** Function Start-CheckedListBoxOptionCheckedListBoxMouseDown ********
  $CheckedListBoxOptionCheckedListBox.add_MouseDown({ Start-CheckedListBoxOptionCheckedListBoxMouseDown -Sender $This -EventArg $PSItem })

  $CheckedListBoxOptionGroupBox.ClientSize = [System.Drawing.Size]::New($CheckedListBoxOptionGroupBox.ClientSize.Width, ($CheckedListBoxOptionCheckedListBox.Bottom + ([MyConfig]::FormSpacer * 2)))

  # ************************************************
  # CheckedListBoxOption ContextMenuStrip
  # ************************************************
  #region $CheckedListBoxOptionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  $CheckedListBoxOptionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()
  #$CheckedListBoxOptionListView.Controls.Add($CheckedListBoxOptionContextMenuStrip)
  $CheckedListBoxOptionContextMenuStrip.BackColor = [MyConfig]::Colors.Back
  #$CheckedListBoxOptionContextMenuStrip.Enabled = $True
  $CheckedListBoxOptionContextMenuStrip.Font = [MyConfig]::Font.Regular
  $CheckedListBoxOptionContextMenuStrip.ForeColor = [MyConfig]::Colors.Fore
  $CheckedListBoxOptionContextMenuStrip.ImageList = $PILSmallImageList
  $CheckedListBoxOptionContextMenuStrip.Name = "CheckedListBoxOptionContextMenuStrip"
  #endregion $CheckedListBoxOptionContextMenuStrip = [System.Windows.Forms.ContextMenuStrip]::New()

  #region ******** Function Start-CheckedListBoxOptionContextMenuStripOpening ********
  function Start-CheckedListBoxOptionContextMenuStripOpening
  {
    <#
      .SYNOPSIS
        Opening Event for the CheckedListBoxOption ContextMenuStrip Control
      .DESCRIPTION
        Opening Event for the CheckedListBoxOption ContextMenuStrip Control
      .PARAMETER Sender
         The ContextMenuStrip Control that fired the Opening Event
      .PARAMETER EventArg
         The Event Arguments for the ContextMenuStrip Opening Event
      .EXAMPLE
         Start-CheckedListBoxOptionContextMenuStripOpening -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Opening Event for `$CheckedListBoxOptionContextMenuStrip"

    [MyConfig]::AutoExit = 0

    #$PILBtmStatusStrip.Items["Status"].Text = "$($Sender.Name)"

    Write-Verbose -Message "Exit Opening Event for `$CheckedListBoxOptionContextMenuStrip"
  }
  #endregion ******** Function Start-CheckedListBoxOptionContextMenuStripOpening ********
  $CheckedListBoxOptionContextMenuStrip.add_Opening({Start-CheckedListBoxOptionContextMenuStripOpening -Sender $This -EventArg $PSItem})

  #region ******** Function Start-CheckedListBoxOptionContextMenuStripItemClick ********
  function Start-CheckedListBoxOptionContextMenuStripItemClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckedListBoxOption ToolStripItem Control
      .DESCRIPTION
        Click Event for the CheckedListBoxOption ToolStripItem Control
      .PARAMETER Sender
         The ToolStripItem Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the ToolStripItem Click Event
      .EXAMPLE
         Start-CheckedListBoxOptionContextMenuStripItemClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$CheckedListBoxOptionContextMenuStripItem"

    [MyConfig]::AutoExit = 0

    switch ($Sender.Name)
    {
      "CheckAll"
      {
        $TmpCheckedItems = @($CheckedListBoxOptionCheckedListBox.CheckedIndices)
        (0..$($CheckedListBoxOptionCheckedListBox.Items.Count - 1)) | Where-Object -FilterScript { $PSItem -notin $TmpCheckedItems } | ForEach-Object -Process { $CheckedListBoxOptionCheckedListBox.SetItemChecked($PSItem, $True) }
        Break
      }
      "UnCheckAll"
      {
        $TmpCheckedItems = @($CheckedListBoxOptionCheckedListBox.CheckedIndices)
        $TmpCheckedItems | ForEach-Object -Process { $CheckedListBoxOptionCheckedListBox.SetItemChecked($PSItem, $False) }
        Break
      }
    }

    Write-Verbose -Message "Exit Click Event for `$CheckedListBoxOptionContextMenuStripItem"
  }
  #endregion ******** Function Start-CheckedListBoxOptionContextMenuStripItemClick ********

  (New-MenuItem -Menu $CheckedListBoxOptionContextMenuStrip -Text "Check All" -Name "CheckAll" -Tag "CheckAll" -DisplayStyle "ImageAndText" -ImageKey "CheckIcon" -PassThru).add_Click({Start-CheckedListBoxOptionContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  (New-MenuItem -Menu $CheckedListBoxOptionContextMenuStrip -Text "Uncheck All" -Name "UnCheckAll" -Tag "UnCheckAll" -DisplayStyle "ImageAndText" -ImageKey "UncheckIcon" -PassThru).add_Click({Start-CheckedListBoxOptionContextMenuStripItemClick -Sender $This -EventArg $PSItem})

  #endregion ******** $CheckedListBoxOptionGroupBox Controls ********

  $TempClientSize = [System.Drawing.Size]::New(($CheckedListBoxOptionGroupBox.Right + [MyConfig]::FormSpacer), ($CheckedListBoxOptionGroupBox.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $CheckedListBoxOptionPanel Controls ********

  # ************************************************
  # CheckedListBoxOptionBtm Panel
  # ************************************************
  #region $CheckedListBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $CheckedListBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()
  $CheckedListBoxOptionForm.Controls.Add($CheckedListBoxOptionBtmPanel)
  $CheckedListBoxOptionBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $CheckedListBoxOptionBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $CheckedListBoxOptionBtmPanel.Name = "CheckedListBoxOptionBtmPanel"
  #endregion $CheckedListBoxOptionBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $CheckedListBoxOptionBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($CheckedListBoxOptionBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $CheckedListBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOptionBtmPanel.Controls.Add($CheckedListBoxOptionBtmLeftButton)
  $CheckedListBoxOptionBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $CheckedListBoxOptionBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $CheckedListBoxOptionBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $CheckedListBoxOptionBtmLeftButton.Font = [MyConfig]::Font.Bold
  $CheckedListBoxOptionBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $CheckedListBoxOptionBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $CheckedListBoxOptionBtmLeftButton.Name = "CheckedListBoxOptionBtmLeftButton"
  $CheckedListBoxOptionBtmLeftButton.TabIndex = 1
  $CheckedListBoxOptionBtmLeftButton.TabStop = $True
  $CheckedListBoxOptionBtmLeftButton.Text = $ButtonLeft
  $CheckedListBoxOptionBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $CheckedListBoxOptionBtmLeftButton.PreferredSize.Height)
  #endregion $CheckedListBoxOptionBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-CheckedListBoxOptionBtmLeftButtonClick ********
  function Start-CheckedListBoxOptionBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckedListBoxOptionBtmLeft Button Control
      .DESCRIPTION
        Click Event for the CheckedListBoxOptionBtmLeft Button Control
      .PARAMETER Sender
         The Button Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the Button Click Event
      .EXAMPLE
         Start-CheckedListBoxOptionBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$CheckedListBoxOptionBtmLeftButton"

    [MyConfig]::AutoExit = 0

    if ($CheckedListBoxOptionCheckedListBox.CheckedItems.Count -gt 0)
    {
      $CheckedListBoxOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
    else
    {
      [Void][System.Windows.Forms.MessageBox]::Show($CheckedListBoxOptionForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }

    Write-Verbose -Message "Exit Click Event for `$CheckedListBoxOptionBtmLeftButton"
  }
  #endregion ******** Function Start-CheckedListBoxOptionBtmLeftButtonClick ********
  $CheckedListBoxOptionBtmLeftButton.add_Click({ Start-CheckedListBoxOptionBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $CheckedListBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOptionBtmPanel.Controls.Add($CheckedListBoxOptionBtmMidButton)
  $CheckedListBoxOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $CheckedListBoxOptionBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $CheckedListBoxOptionBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $CheckedListBoxOptionBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $CheckedListBoxOptionBtmMidButton.Font = [MyConfig]::Font.Bold
  $CheckedListBoxOptionBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $CheckedListBoxOptionBtmMidButton.Location = [System.Drawing.Point]::New(($CheckedListBoxOptionBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $CheckedListBoxOptionBtmMidButton.Name = "CheckedListBoxOptionBtmMidButton"
  $CheckedListBoxOptionBtmMidButton.TabIndex = 2
  $CheckedListBoxOptionBtmMidButton.TabStop = $True
  $CheckedListBoxOptionBtmMidButton.Text = $ButtonMid
  $CheckedListBoxOptionBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $CheckedListBoxOptionBtmMidButton.PreferredSize.Height)
  #endregion $CheckedListBoxOptionBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-CheckedListBoxOptionBtmMidButtonClick ********
  function Start-CheckedListBoxOptionBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckedListBoxOptionBtmMid Button Control
      .DESCRIPTION
        Click Event for the CheckedListBoxOptionBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-CheckedListBoxOptionBtmMidButtonClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$CheckedListBoxOptionBtmMidButton"

    [MyConfig]::AutoExit = 0

    $TmpCheckedItems = @($CheckedListBoxOptionCheckedListBox.CheckedIndices)
    $TmpCheckedItems | ForEach-Object -Process { $CheckedListBoxOptionCheckedListBox.SetItemChecked($PSItem, $False) }
    if ($CheckedListBoxOptionCheckedListBox.Tag.Count -gt 0)
    {
      $CheckedListBoxOptionCheckedListBox.Tag | ForEach-Object -Process { $CheckedListBoxOptionCheckedListBox.SetItemChecked($CheckedListBoxOptionCheckedListBox.Items.IndexOf($PSItem), $True) }
    }

    Write-Verbose -Message "Exit Click Event for `$CheckedListBoxOptionBtmMidButton"
  }
  #endregion ******** Function Start-CheckedListBoxOptionBtmMidButtonClick ********
  $CheckedListBoxOptionBtmMidButton.add_Click({ Start-CheckedListBoxOptionBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $CheckedListBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()
  $CheckedListBoxOptionBtmPanel.Controls.Add($CheckedListBoxOptionBtmRightButton)
  $CheckedListBoxOptionBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $CheckedListBoxOptionBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $CheckedListBoxOptionBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $CheckedListBoxOptionBtmRightButton.Font = [MyConfig]::Font.Bold
  $CheckedListBoxOptionBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $CheckedListBoxOptionBtmRightButton.Location = [System.Drawing.Point]::New(($CheckedListBoxOptionBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $CheckedListBoxOptionBtmRightButton.Name = "CheckedListBoxOptionBtmRightButton"
  $CheckedListBoxOptionBtmRightButton.TabIndex = 3
  $CheckedListBoxOptionBtmRightButton.TabStop = $True
  $CheckedListBoxOptionBtmRightButton.Text = $ButtonRight
  $CheckedListBoxOptionBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $CheckedListBoxOptionBtmRightButton.PreferredSize.Height)
  #endregion $CheckedListBoxOptionBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-CheckedListBoxOptionBtmRightButtonClick ********
  function Start-CheckedListBoxOptionBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the CheckedListBoxOptionBtmRight Button Control
      .DESCRIPTION
        Click Event for the CheckedListBoxOptionBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-CheckedListBoxOptionBtmRightButtonClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$CheckedListBoxOptionBtmRightButton"

    [MyConfig]::AutoExit = 0

    # Cancel Code Goes here

    $CheckedListBoxOptionForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$CheckedListBoxOptionBtmRightButton"
  }
  #endregion ******** Function Start-CheckedListBoxOptionBtmRightButtonClick ********
  $CheckedListBoxOptionBtmRightButton.add_Click({ Start-CheckedListBoxOptionBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $CheckedListBoxOptionBtmPanel.ClientSize = [System.Drawing.Size]::New(($CheckedListBoxOptionBtmRightButton.Right + [MyConfig]::FormSpacer), ($CheckedListBoxOptionBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $CheckedListBoxOptionBtmPanel Controls ********

  $CheckedListBoxOptionForm.ClientSize = [System.Drawing.Size]::New($CheckedListBoxOptionForm.ClientSize.Width, ($TempClientSize.Height + $CheckedListBoxOptionBtmPanel.Height))

  #endregion ******** Controls for CheckedListBoxOption Form ********

  #endregion ******** End **** CheckedListBoxOption **** End ********

  $DialogResult = $CheckedListBoxOptionForm.ShowDialog()
  if ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
  {
    [CheckedListBoxOption]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, $CheckedListBoxOptionCheckedListBox.CheckedItems)
  }
  else
  {
    [CheckedListBoxOption]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, @())
  }

  $CheckedListBoxOptionForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-CheckedListBoxOption"
}
#endregion function Get-CheckedListBoxOption

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
# Get ComboBoxFilter Function
# --------------------------------
#region ComboBoxFilterItem Class
Class ComboBoxFilterItem
{
  [String]$Text
  [Object]$Value
  
  ComboBoxFilterItem ([String]$Text, [Object]$Value)
  {
    $This.Text = $Text
    $This.Value = $Value
  }
}
#endregion ComboBoxFilterItem Class

#region ComboBoxFilter Result Class
Class ComboBoxFilter
{
  [Bool]$Success
  [Object]$DialogResult
  [HashTable]$Values

  ComboBoxFilter ([Bool]$Success, [Object]$DialogResult, [HashTable]$Values)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
    $This.Values = $Values
  }
}
#endregion ComboBoxFilter Result Class

#region function Get-ComboBoxFilter
Function Get-ComboBoxFilter ()
{
  <#
    .SYNOPSIS
      Shows Get-ComboBoxFilter
    .DESCRIPTION
      Shows Get-ComboBoxFilter
    .PARAMETER Title
      Title of the Get-ComboBoxFilter Dialog Window
    .PARAMETER Message
      Message to Show
    .PARAMETER Items
      Items to show in the ComboBox
    .PARAMETER Properties
      Name of the Properties to Filter On
    .PARAMETER Selected
      Default Selected ComboBox Values
    .PARAMETER Width
      Width of Get-ComboBoxFilter Dialog Window
    .PARAMETER NoFilter
      Do Not Filter ComBox Items from other Selected ComboBox Items
    .PARAMETER ButtonLeft
      Left Button DaialogResult
    .PARAMETER ButtonMid
      Missing Button DaialogResult
    .PARAMETER ButtonRight
      Right Button DaialogResult
    .EXAMPLE
      $ServiceList = @(Get-Service | Select-Object -Property Status, Name, StartType)
      $DialogResult = Get-ComboBoxFilter -Title "Combo Filter Dialog 01" -Message "Show this Sample Message Prompt to the User" -Items $ServiceList -Properties Status, Name, StartType
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
  Param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [parameter(Mandatory = $True)]
    [String]$Message = "Status Message",
    [Object[]]$Items = @(),
    [String[]]$Properties,
    [HashTable]$Selected = @{},
    [Int]$Width = 35,
    [Switch]$NoFilter,
    [String]$ButtonLeft = "&OK",
    [String]$ButtonMid = "&Reset",
    [String]$ButtonRight = "&Cancel"
  )
  Write-Verbose -Message "Enter Function Get-ComboBoxFilter"

  #region ******** Begin **** ComboBoxFilter **** Begin ********

  # ************************************************
  # ComboBoxFilter Form
  # ************************************************
  #region $ComboBoxFilterForm = [System.Windows.Forms.Form]::New()
  $ComboBoxFilterForm = [System.Windows.Forms.Form]::New()
  $ComboBoxFilterForm.BackColor = [MyConfig]::Colors.Back
  $ComboBoxFilterForm.Font = [MyConfig]::Font.Regular
  $ComboBoxFilterForm.ForeColor = [MyConfig]::Colors.Fore
  $ComboBoxFilterForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $ComboBoxFilterForm.Icon = $PILForm.Icon
  $ComboBoxFilterForm.KeyPreview = $True
  $ComboBoxFilterForm.MaximizeBox = $False
  $ComboBoxFilterForm.MinimizeBox = $False
  $ComboBoxFilterForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), 0)
  $ComboBoxFilterForm.Name = "ComboBoxFilterForm"
  $ComboBoxFilterForm.Owner = $PILForm
  $ComboBoxFilterForm.ShowInTaskbar = $False
  $ComboBoxFilterForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $ComboBoxFilterForm.Text = $Title
  #endregion $ComboBoxFilterForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-ComboBoxFilterFormKeyDown ********
  Function Start-ComboBoxFilterFormKeyDown
  {
    <#
      .SYNOPSIS
        KeyDown Event for the ComboBoxFilter Form Control
      .DESCRIPTION
        KeyDown Event for the ComboBoxFilter Form Control
      .PARAMETER Sender
        The Form Control that fired the KeyDown Event
      .PARAMETER EventArg
        The Event Arguments for the Form KeyDown Event
      .EXAMPLE
        Start-ComboBoxFilterFormKeyDown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter KeyDown Event for `$ComboBoxFilterForm"

    [MyConfig]::AutoExit = 0

    If ($EventArg.KeyCode -eq [System.Windows.Forms.Keys]::Escape)
    {
      $ComboBoxFilterForm.Close()
    }

    Write-Verbose -Message "Exit KeyDown Event for `$ComboBoxFilterForm"
  }
  #endregion ******** Function Start-ComboBoxFilterFormKeyDown ********
  $ComboBoxFilterForm.add_KeyDown({ Start-ComboBoxFilterFormKeyDown -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ComboBoxFilterFormShown ********
  Function Start-ComboBoxFilterFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the ComboBoxFilter Form Control
      .DESCRIPTION
        Shown Event for the ComboBoxFilter Form Control
      .PARAMETER Sender
        The Form Control that fired the Shown Event
      .PARAMETER EventArg
        The Event Arguments for the Form Shown Event
      .EXAMPLE
        Start-ComboBoxFilterFormShown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Shown Event for `$ComboBoxFilterForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Shown Event for `$ComboBoxFilterForm"
  }
  #endregion ******** Function Start-ComboBoxFilterFormShown ********
  $ComboBoxFilterForm.add_Shown({ Start-ComboBoxFilterFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for ComboBoxFilter Form ********

  # ************************************************
  # ComboBoxFilter Panel
  # ************************************************
  #region $ComboBoxFilterPanel = [System.Windows.Forms.Panel]::New()
  $ComboBoxFilterPanel = [System.Windows.Forms.Panel]::New()
  $ComboBoxFilterForm.Controls.Add($ComboBoxFilterPanel)
  $ComboBoxFilterPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ComboBoxFilterPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $ComboBoxFilterPanel.Name = "ComboBoxFilterPanel"
  #endregion $ComboBoxFilterPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ComboBoxFilterPanel Controls ********

  #region $ComboBoxFilterLabel = [System.Windows.Forms.Label]::New()
  $ComboBoxFilterLabel = [System.Windows.Forms.Label]::New()
  $ComboBoxFilterPanel.Controls.Add($ComboBoxFilterLabel)
  $ComboBoxFilterLabel.ForeColor = [MyConfig]::Colors.LabelFore
  $ComboBoxFilterLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ([MyConfig]::FormSpacer * 2))
  $ComboBoxFilterLabel.Name = "ComboBoxFilterLabel"
  $ComboBoxFilterLabel.Size = [System.Drawing.Size]::New(($ComboBoxFilterPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), 23)
  $ComboBoxFilterLabel.Text = $Message
  $ComboBoxFilterLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  #endregion $ComboBoxFilterLabel = [System.Windows.Forms.Label]::New()

  # Returns the minimum size required to display the text
  $TmpSize = [System.Windows.Forms.TextRenderer]::MeasureText($ComboBoxFilterLabel.Text, [MyConfig]::Font.Regular, $ComboBoxFilterLabel.Size, ([System.Windows.Forms.TextFormatFlags]("Top", "Left", "WordBreak")))
  $ComboBoxFilterLabel.Size = [System.Drawing.Size]::New(($ComboBoxFilterPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($TmpSize.Height + [MyConfig]::Font.Height))

  If ($PSBoundParameters.ContainsKey("Properties"))
  {
    $FilterOptionNames = $Properties
  }
  Else
  {
    $FilterOptionNames = ($Items | Select-Object -First 1).PSObject.Properties | Select-Object -ExpandProperty Name
  }

  # ************************************************
  # ComboBoxFilter GroupBox
  # ************************************************
  #region $ComboBoxFilterGroupBox = [System.Windows.Forms.GroupBox]::New()
  $ComboBoxFilterGroupBox = [System.Windows.Forms.GroupBox]::New()
  # Location of First Control = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::Font.Height)
  $ComboBoxFilterPanel.Controls.Add($ComboBoxFilterGroupBox)
  $ComboBoxFilterGroupBox.BackColor = [MyConfig]::Colors.Back
  $ComboBoxFilterGroupBox.Font = [MyConfig]::Font.Regular
  $ComboBoxFilterGroupBox.ForeColor = [MyConfig]::Colors.GroupFore
  $ComboBoxFilterGroupBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($ComboBoxFilterLabel.Bottom + ([MyConfig]::FormSpacer * 2)))
  $ComboBoxFilterGroupBox.Name = "ComboBoxFilterGroupBox"
  $ComboBoxFilterGroupBox.Size = [System.Drawing.Size]::New(($ComboBoxFilterPanel.Width - ([MyConfig]::FormSpacer * 2)), 50)
  #endregion $ComboBoxFilterGroupBox = [System.Windows.Forms.GroupBox]::New()

  #region ******** $ComboBoxFilterGroupBox Controls ********

  #region ******** Function Start-GetComboFilterComboBoxSelectedIndexChanged ********
  Function Start-GetComboFilterComboBoxSelectedIndexChanged
  {
  <#
    .SYNOPSIS
      SelectedIndexChanged Event for the GetSiteComboChoice ComboBox Control
    .DESCRIPTION
      SelectedIndexChanged Event for the GetSiteComboChoice ComboBox Control
    .PARAMETER Sender
       The ComboBox Control that fired the SelectedIndexChanged Event
    .PARAMETER EventArg
       The Event Arguments for the ComboBox SelectedIndexChanged Event
    .EXAMPLE
       Start-GetComboFilterComboBoxSelectedIndexChanged -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By ken.sweet
  #>
    [CmdletBinding()]
    Param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ComboBox]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter SelectedIndexChanged Event for `$GetSiteComboChoiceComboBox"

    [MyConfig]::AutoExit = 0

    $ValidItems = @($Items)
    ForEach ($FilterOptionName In $FilterOptionNames)
    {
      $ValidItems = @($ValidItems | Where-Object -FilterScript { $PSItem.($FilterOptionName) -like $ComboBoxFilterGroupBox.Controls[$FilterOptionName].SelectedItem.Value })
    }

    ForEach ($FilterOptionName In $FilterOptionNames)
    {
      $ValidItemNames = @($ValidItems | Select-Object -ExpandProperty $FilterOptionName -Unique)
      If ($FilterOptionName -ne $Sender.Name)
      {
        $RemoveList = @($ComboBoxFilterGroupBox.Controls[$FilterOptionName].Items | Where-Object -FilterScript { ($PSItem.Text -notin $ValidItemNames) -and ($PSItem.Value -ne "*") })
        ForEach ($RemoveItem In $RemoveList)
        {
          $ComboBoxFilterGroupBox.Controls[$FilterOptionName].Items.Remove($RemoveItem)
        }
      }
      $HaveItemNames = @($ComboBoxFilterGroupBox.Controls[$FilterOptionName].Items | Select-Object -ExpandProperty Text -Unique)
      $AddList = @($ComboBoxFilterGroupBox.Controls[$FilterOptionName].Tag.Items | Where-Object -FilterScript { ($PSItem.Text -in $ValidItemNames) -and ($PSItem.Text -notin $HaveItemNames) })
      $ComboBoxFilterGroupBox.Controls[$FilterOptionName].Items.AddRange($AddList)
    }

    Write-Verbose -Message "Exit SelectedIndexChanged Event for `$GetSiteComboChoiceComboBox"
  }
  #endregion ******** Function Start-GetComboFilterComboBoxSelectedIndexChanged ********

  $GroupBottom = [MyConfig]::Font.Height
  ForEach ($FilterOptionName In $FilterOptionNames)
  {
    #region $TmpFilterComboBox = [System.Windows.Forms.ComboBox]::New()
    $TmpFilterComboBox = [System.Windows.Forms.ComboBox]::New()
    $ComboBoxFilterGroupBox.Controls.Add($TmpFilterComboBox)
    $TmpFilterComboBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom")
    $TmpFilterComboBox.AutoSize = $True
    $TmpFilterComboBox.BackColor = [MyConfig]::Colors.TextBack
    $TmpFilterComboBox.DisplayMember = "Text"
    $TmpFilterComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $TmpFilterComboBox.Font = [MyConfig]::Font.Regular
    $TmpFilterComboBox.ForeColor = [MyConfig]::Colors.TextFore
    [void]$TmpFilterComboBox.Items.Add([PSCustomObject]@{ "Text" = " - Return All $($FilterOptionName) Values - "; "Value" = "*" })
    $TmpFilterComboBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, $GroupBottom)
    $TmpFilterComboBox.Name = $FilterOptionName
    $TmpFilterComboBox.SelectedIndex = 0
    $TmpFilterComboBox.Size = [System.Drawing.Size]::New(($ComboBoxFilterGroupBox.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), $TmpFilterComboBox.PreferredHeight)
    $TmpFilterComboBox.Sorted = $True
    $TmpFilterComboBox.TabIndex = 0
    $TmpFilterComboBox.TabStop = $True
    $TmpFilterComboBox.Tag = $Null
    $TmpFilterComboBox.ValueMember = "Value"
    #endregion $TmpFilterComboBox = [System.Windows.Forms.ComboBox]::New()

    $TmpFilterComboBox.SelectedIndex = 0
    $TmpFilterComboBox.Items.AddRange(@($Items | Where-Object -FilterScript { -not [String]::IsNullOrEmpty($PSITem.($FilterOptionName)) } | Sort-Object -Property $FilterOptionName -Unique | ForEach-Object -Process { [ComboBoxFilterItem]::New($PSITem.($FilterOptionName), $PSITem.($FilterOptionName)) }))
    $TmpFilterComboBox.Tag = @{ "Items" = @($TmpFilterComboBox.Items); "SelectedItem" = $Null }

    if (-not $NoFilter.IsPresent)
    {
      $TmpFilterComboBox.add_SelectedIndexChanged({ Start-GetComboFilterComboBoxSelectedIndexChanged -Sender $This -EventArg $PSItem })
    }

    $GroupBottom = ($TmpFilterComboBox.Bottom + [MyConfig]::FormSpacer)
  }

  $ComboBoxFilterGroupBox.ClientSize = [System.Drawing.Size]::New($ComboBoxFilterGroupBox.ClientSize.Width, ($GroupBottom + [MyConfig]::FormSpacer))

  #endregion ******** $ComboBoxFilterGroupBox Controls ********

  ForEach ($FilterOptionName In $FilterOptionNames)
  {
    # $Sender
    If ($Selected.ContainsKey($FilterOptionName))
    {
      $TmpItem = $ComboBoxFilterGroupBox.Controls[$FilterOptionName].Items | Where-Object -FilterScript { $PSItem.Value -eq $Selected.($FilterOptionName) }
      If (-not [String]::IsNullOrEmpty($TmpItem.Text))
      {
        $ComboBoxFilterGroupBox.Controls[$FilterOptionName].SelectedItem = $TmpItem
      }
    }
    $ComboBoxFilterGroupBox.Controls[$FilterOptionName].Tag.SelectedItem = $ComboBoxFilterGroupBox.Controls[$FilterOptionName].SelectedItem
  }

  $TempClientSize = [System.Drawing.Size]::New(($ComboBoxFilterGroupBox.Right + [MyConfig]::FormSpacer), ($ComboBoxFilterGroupBox.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ComboBoxFilterPanel Controls ********

  # ************************************************
  # ComboBoxFilterBtm Panel
  # ************************************************
  #region $ComboBoxFilterBtmPanel = [System.Windows.Forms.Panel]::New()
  $ComboBoxFilterBtmPanel = [System.Windows.Forms.Panel]::New()
  $ComboBoxFilterForm.Controls.Add($ComboBoxFilterBtmPanel)
  $ComboBoxFilterBtmPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ComboBoxFilterBtmPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $ComboBoxFilterBtmPanel.Name = "ComboBoxFilterBtmPanel"
  #endregion $ComboBoxFilterBtmPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ComboBoxFilterBtmPanel Controls ********

  # Evenly Space Buttons - Move Size to after Text
  $NumButtons = 3
  $TempSpace = [Math]::Floor($ComboBoxFilterBtmPanel.ClientSize.Width - ([MyConfig]::FormSpacer * ($NumButtons + 1)))
  $TempWidth = [Math]::Floor($TempSpace / $NumButtons)
  $TempMod = $TempSpace % $NumButtons

  #region $ComboBoxFilterBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ComboBoxFilterBtmLeftButton = [System.Windows.Forms.Button]::New()
  $ComboBoxFilterBtmPanel.Controls.Add($ComboBoxFilterBtmLeftButton)
  $ComboBoxFilterBtmLeftButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left")
  $ComboBoxFilterBtmLeftButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ComboBoxFilterBtmLeftButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ComboBoxFilterBtmLeftButton.Font = [MyConfig]::Font.Bold
  $ComboBoxFilterBtmLeftButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ComboBoxFilterBtmLeftButton.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $ComboBoxFilterBtmLeftButton.Name = "ComboBoxFilterBtmLeftButton"
  $ComboBoxFilterBtmLeftButton.TabIndex = 1
  $ComboBoxFilterBtmLeftButton.TabStop = $True
  $ComboBoxFilterBtmLeftButton.Text = $ButtonLeft
  $ComboBoxFilterBtmLeftButton.Size = [System.Drawing.Size]::New($TempWidth, $ComboBoxFilterBtmLeftButton.PreferredSize.Height)
  #endregion $ComboBoxFilterBtmLeftButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ComboBoxFilterBtmLeftButtonClick ********
  Function Start-ComboBoxFilterBtmLeftButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ComboBoxFilterBtmLeft Button Control
      .DESCRIPTION
        Click Event for the ComboBoxFilterBtmLeft Button Control
      .PARAMETER Sender
         The Button Control that fired the Click Event
      .PARAMETER EventArg
         The Event Arguments for the Button Click Event
      .EXAMPLE
         Start-ComboBoxFilterBtmLeftButtonClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$ComboBoxFilterBtmLeftButton"

    [MyConfig]::AutoExit = 0

    $ValidateClick = 0
    ForEach ($FilterOptionName In $FilterOptionNames)
    {
      $ValidateClick = $ValidateClick + $ComboBoxFilterGroupBox.Controls[$FilterOptionName].SelectedIndex
    }
    If ($ValidateClick -eq 0)
    {
      [Void][System.Windows.Forms.MessageBox]::Show($ComboBoxFilterForm, "Missing or Invalid Value.", [MyConfig]::ScriptName, "OK", "Warning")
    }
    Else
    {
      $ComboBoxFilterForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }

    Write-Verbose -Message "Exit Click Event for `$ComboBoxFilterBtmLeftButton"
  }
  #endregion ******** Function Start-ComboBoxFilterBtmLeftButtonClick ********
  $ComboBoxFilterBtmLeftButton.add_Click({ Start-ComboBoxFilterBtmLeftButtonClick -Sender $This -EventArg $PSItem })

  #region $ComboBoxFilterBtmMidButton = [System.Windows.Forms.Button]::New()
  $ComboBoxFilterBtmMidButton = [System.Windows.Forms.Button]::New()
  $ComboBoxFilterBtmPanel.Controls.Add($ComboBoxFilterBtmMidButton)
  $ComboBoxFilterBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $ComboBoxFilterBtmMidButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
  $ComboBoxFilterBtmMidButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ComboBoxFilterBtmMidButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ComboBoxFilterBtmMidButton.Font = [MyConfig]::Font.Bold
  $ComboBoxFilterBtmMidButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ComboBoxFilterBtmMidButton.Location = [System.Drawing.Point]::New(($ComboBoxFilterBtmLeftButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ComboBoxFilterBtmMidButton.Name = "ComboBoxFilterBtmMidButton"
  $ComboBoxFilterBtmMidButton.TabIndex = 2
  $ComboBoxFilterBtmMidButton.TabStop = $True
  $ComboBoxFilterBtmMidButton.Text = $ButtonMid
  $ComboBoxFilterBtmMidButton.Size = [System.Drawing.Size]::New(($TempWidth + $TempMod), $ComboBoxFilterBtmMidButton.PreferredSize.Height)
  #endregion $ComboBoxFilterBtmMidButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ComboBoxFilterBtmMidButtonClick ********
  Function Start-ComboBoxFilterBtmMidButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ComboBoxFilterBtmMid Button Control
      .DESCRIPTION
        Click Event for the ComboBoxFilterBtmMid Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ComboBoxFilterBtmMidButtonClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$ComboBoxFilterBtmMidButton"

    [MyConfig]::AutoExit = 0

    ForEach ($FilterOptionName In $FilterOptionNames)
    {
      $ComboBoxFilterGroupBox.Controls[$FilterOptionName].SelectedIndex = 0
    }

    ForEach ($FilterOptionName In $FilterOptionNames)
    {
      $ComboBoxFilterGroupBox.Controls[$FilterOptionName].SelectedItem = $ComboBoxFilterGroupBox.Controls[$FilterOptionName].Tag.SelectedItem
    }

    Write-Verbose -Message "Exit Click Event for `$ComboBoxFilterBtmMidButton"
  }
  #endregion ******** Function Start-ComboBoxFilterBtmMidButtonClick ********
  $ComboBoxFilterBtmMidButton.add_Click({ Start-ComboBoxFilterBtmMidButtonClick -Sender $This -EventArg $PSItem })

  #region $ComboBoxFilterBtmRightButton = [System.Windows.Forms.Button]::New()
  $ComboBoxFilterBtmRightButton = [System.Windows.Forms.Button]::New()
  $ComboBoxFilterBtmPanel.Controls.Add($ComboBoxFilterBtmRightButton)
  $ComboBoxFilterBtmRightButton.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Right")
  $ComboBoxFilterBtmRightButton.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $ComboBoxFilterBtmRightButton.BackColor = [MyConfig]::Colors.ButtonBack
  $ComboBoxFilterBtmRightButton.Font = [MyConfig]::Font.Bold
  $ComboBoxFilterBtmRightButton.ForeColor = [MyConfig]::Colors.ButtonFore
  $ComboBoxFilterBtmRightButton.Location = [System.Drawing.Point]::New(($ComboBoxFilterBtmMidButton.Right + [MyConfig]::FormSpacer), [MyConfig]::FormSpacer)
  $ComboBoxFilterBtmRightButton.Name = "ComboBoxFilterBtmRightButton"
  $ComboBoxFilterBtmRightButton.TabIndex = 3
  $ComboBoxFilterBtmRightButton.TabStop = $True
  $ComboBoxFilterBtmRightButton.Text = $ButtonRight
  $ComboBoxFilterBtmRightButton.Size = [System.Drawing.Size]::New($TempWidth, $ComboBoxFilterBtmRightButton.PreferredSize.Height)
  #endregion $ComboBoxFilterBtmRightButton = [System.Windows.Forms.Button]::New()

  #region ******** Function Start-ComboBoxFilterBtmRightButtonClick ********
  Function Start-ComboBoxFilterBtmRightButtonClick
  {
    <#
      .SYNOPSIS
        Click Event for the ComboBoxFilterBtmRight Button Control
      .DESCRIPTION
        Click Event for the ComboBoxFilterBtmRight Button Control
      .PARAMETER Sender
        The Button Control that fired the Click Event
      .PARAMETER EventArg
        The Event Arguments for the Button Click Event
      .EXAMPLE
        Start-ComboBoxFilterBtmRightButtonClick -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Click Event for `$ComboBoxFilterBtmRightButton"

    [MyConfig]::AutoExit = 0

    $ComboBoxFilterForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    Write-Verbose -Message "Exit Click Event for `$ComboBoxFilterBtmRightButton"
  }
  #endregion ******** Function Start-ComboBoxFilterBtmRightButtonClick ********
  $ComboBoxFilterBtmRightButton.add_Click({ Start-ComboBoxFilterBtmRightButtonClick -Sender $This -EventArg $PSItem })

  $ComboBoxFilterBtmPanel.ClientSize = [System.Drawing.Size]::New(($ComboBoxFilterBtmRightButton.Right + [MyConfig]::FormSpacer), ($ComboBoxFilterBtmRightButton.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ComboBoxFilterBtmPanel Controls ********

  $ComboBoxFilterForm.ClientSize = [System.Drawing.Size]::New($ComboBoxFilterForm.ClientSize.Width, ($TempClientSize.Height + $ComboBoxFilterBtmPanel.Height))

  #endregion ******** Controls for ComboBoxFilter Form ********

  #endregion ******** End **** ComboBoxFilter **** End ********

  $DialogResult = $ComboBoxFilterForm.ShowDialog()
  If ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
  {
    $TmpHash = [HashTable]::New()
    ForEach ($FilterOptionName In $FilterOptionNames)
    {
      [Void]$TmpHash.Add($FilterOptionName, $ComboBoxFilterGroupBox.Controls[$FilterOptionName].SelectedItem.Value)
    }
    [ComboBoxFilter]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, $TmpHash)
  }
  Else
  {
    [ComboBoxFilter]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult, @{ })
  }

  $ComboBoxFilterForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Get-ComboBoxFilter"
}
#endregion function Get-ComboBoxFilter

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

# --------------------------------
# Show ScriptInfo Function
# --------------------------------
#region ScriptInfo Info Topics

#region $InfoIntro Compressed RTF
$InfoIntro = @"
77u/e1xydGYxXGFuc2lcYW5zaWNwZzEyNTJcZGVmZjBcbm91aWNvbXBhdFxkZWZsYW5nMTAzM3tcZm9udHRibHtcZjBcZm5pbCBWZXJkYW5hO317XGYxXGZuaWxcZmNoYXJzZXQwIFZlcmRhbmE7fXtcZjJcZm5p
bFxmY2hhcnNldDAgQ2FsaWJyaTt9fQ0Ke1wqXGdlbmVyYXRvciBSaWNoZWQyMCAxMC4wLjE5MDQxfVx2aWV3a2luZDRcdWMxIA0KXHBhcmRccWNcYlxmMFxmczMwIEhlbHAgSW50b2R1Y3Rpb25ccGFyDQpcYjBc
ZjFcZnMyMFxwYXINClRoaXMgaXMgTXkgSGVscCBJbnRvZHVjdGlvbiFccGFyDQoNClxwYXJkXHNhMjAwXHNsMjc2XHNsbXVsdDFcZjJcZnMyMlxsYW5nOVxwYXINCn0NCgA=
"@
#endregion $InfoIntro Compressed RTF

#region $Info01 Compressed RTF
$Info01 = @"
77u/e1xydGYxXGFuc2lcYW5zaWNwZzEyNTJcZGVmZjBcbm91aWNvbXBhdFxkZWZsYW5nMTAzM3tcZm9udHRibHtcZjBcZm5pbCBWZXJkYW5hO317XGYxXGZuaWxcZmNoYXJzZXQwIFZlcmRhbmE7fXtcZjJcZm5p
bFxmY2hhcnNldDAgQ2FsaWJyaTt9fQ0Ke1wqXGdlbmVyYXRvciBSaWNoZWQyMCAxMC4wLjE5MDQxfVx2aWV3a2luZDRcdWMxIA0KXHBhcmRccWNcYlxmMFxmczMwIEhlbHAgVG9waWMgMDFccGFyDQpcYjBcZjFc
ZnMyMFxwYXINClRoaXMgaXMgTXkgSGVscCBUb3BpYyFccGFyDQoNClxwYXJkXHNhMjAwXHNsMjc2XHNsbXVsdDFcZjJcZnMyMlxsYW5nOVxwYXINCn0NCgA=
"@
#endregion $Info01 Compressed RTF

#region $Info02 Compressed RTF
$Info02 = @"
77u/e1xydGYxXGFuc2lcYW5zaWNwZzEyNTJcZGVmZjBcbm91aWNvbXBhdFxkZWZsYW5nMTAzM3tcZm9udHRibHtcZjBcZm5pbCBWZXJkYW5hO317XGYxXGZuaWxcZmNoYXJzZXQwIFZlcmRhbmE7fXtcZjJcZm5p
bFxmY2hhcnNldDAgQ2FsaWJyaTt9fQ0Ke1wqXGdlbmVyYXRvciBSaWNoZWQyMCAxMC4wLjE5MDQxfVx2aWV3a2luZDRcdWMxIA0KXHBhcmRccWNcYlxmMFxmczMwIEhlbHAgVG9waWMgMFxmMSAyXGYwXHBhcg0K
XGIwXGYxXGZzMjBccGFyDQpUaGlzIGlzIE15IEhlbHAgVG9waWMhXHBhcg0KDQpccGFyZFxzYTIwMFxzbDI3NlxzbG11bHQxXGYyXGZzMjJcbGFuZzlccGFyDQp9DQoA
"@
#endregion $Info02 Compressed RTF

#region $Info03 Compressed RTF
$Info03 = @"
77u/e1xydGYxXGFuc2lcYW5zaWNwZzEyNTJcZGVmZjBcbm91aWNvbXBhdFxkZWZsYW5nMTAzM3tcZm9udHRibHtcZjBcZm5pbCBWZXJkYW5hO317XGYxXGZuaWxcZmNoYXJzZXQwIFZlcmRhbmE7fXtcZjJcZm5p
bFxmY2hhcnNldDAgQ2FsaWJyaTt9fQ0Ke1wqXGdlbmVyYXRvciBSaWNoZWQyMCAxMC4wLjE5MDQxfVx2aWV3a2luZDRcdWMxIA0KXHBhcmRccWNcYlxmMFxmczMwIEhlbHAgVG9waWMgMFxmMSAzXGYwXHBhcg0K
XGIwXGYxXGZzMjBccGFyDQpUaGlzIGlzIE15IEhlbHAgVG9waWMhXHBhcg0KDQpccGFyZFxzYTIwMFxzbDI3NlxzbG11bHQxXGYyXGZzMjJcbGFuZzlccGFyDQp9DQoA
"@
#endregion $Info03 Compressed RTF

$ScriptInfoTopics = [Ordered]@{}
$ScriptInfoTopics.Add("InfoIntro", @{"Name" = "Info Introduction"; "Data" = $InfoIntro; "Type" = "Base64"})
$ScriptInfoTopics.Add("Info01", @{"Name" = "Info Topic 01"; "Data" = $Info01; "Type" = "Base64"})
$ScriptInfoTopics.Add("Info02", @{"Name" = "Info Topic 02"; "Data" = $Info02; "Type" = "Base64"})
$ScriptInfoTopics.Add("Info03", @{"Name" = "Info Topic 03"; "Data" = $Info03; "Type" = "Base64"})

$InfoIntro = $Null
$Info01 = $Null
$Info02 = $Null
$Info03 = $Null

#endregion ScriptInfo Dialog Info Topics

#region function Show-ScriptInfo
function Show-ScriptInfo ()
{
  <#
    .SYNOPSIS
      Shows Show-ScriptInfo
    .DESCRIPTION
      Shows Show-ScriptInfo
    .PARAMETER Title
      Show-ScriptInfo Window Title
    .PARAMETER InfoTitle
      Title of Into Topics
    .PARAMETER Topics
      Orders List of Tpoic to Display
    .PARAMETER DefInfoTopic
      Default Infomration Topic
    .PARAMETER Width
      Width of the Show-ScriptInfo Window
    .PARAMETER Height
      Height of the Show-ScriptInfo Window
    .EXAMPLE
      $Return = Show-ScriptInfo -Topics $Topics
    .NOTES
      Original Function By Ken Sweet
  #>
  [CmdletBinding()]
  param (
    [String]$Title = "$([MyConfig]::ScriptName)",
    [String]$InfoTitle = " << FCG Info Topics >> ",
    [String]$DefInfoTopic = "InfoIntro",
    [System.Collections.Specialized.OrderedDictionary]$Topics = $ScriptInfoTopics,
    [Int]$Width = 60,
    [Int]$Height = 24
  )
  Write-Verbose -Message "Enter Function Show-ScriptInfo"

  #region ******** Begin **** ScriptInfo **** Begin ********

  # ************************************************
  # ScriptInfo Form
  # ************************************************
  #region $ScriptInfoForm = [System.Windows.Forms.Form]::New()
  $ScriptInfoForm = [System.Windows.Forms.Form]::New()
  $ScriptInfoForm.BackColor = [MyConfig]::Colors.Back
  $ScriptInfoForm.Font = [MyConfig]::Font.Regular
  $ScriptInfoForm.ForeColor = [MyConfig]::Colors.Fore
  $ScriptInfoForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
  $ScriptInfoForm.Icon = $FCGForm.Icon
  $ScriptInfoForm.MaximizeBox = $False
  $ScriptInfoForm.MinimizeBox = $False
  $ScriptInfoForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * $Height))
  $ScriptInfoForm.Name = "ScriptInfoForm"
  $ScriptInfoForm.Owner = $FCGForm
  $ScriptInfoForm.ShowInTaskbar = $False
  $ScriptInfoForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $ScriptInfoForm.Text = $Title
  #endregion $ScriptInfoForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-ScriptInfoFormMove ********
  function Start-ScriptInfoFormMove
  {
    <#
      .SYNOPSIS
        Move Event for the ScriptInfo Form Control
      .DESCRIPTION
        Move Event for the ScriptInfo Form Control
      .PARAMETER Sender
        The Form Control that fired the Move Event
      .PARAMETER EventArg
        The Event Arguments for the Form Move Event
      .EXAMPLE
        Start-ScriptInfoFormMove -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Move Event for `$ScriptInfoForm"

    [MyConfig]::AutoExit = 0

    Write-Verbose -Message "Exit Move Event for `$ScriptInfoForm"
  }
  #endregion ******** Function Start-ScriptInfoFormMove ********
  $ScriptInfoForm.add_Move({ Start-ScriptInfoFormMove -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ScriptInfoFormResize ********
  function Start-ScriptInfoFormResize
  {
    <#
      .SYNOPSIS
        Resize Event for the ScriptInfo Form Control
      .DESCRIPTION
        Resize Event for the ScriptInfo Form Control
      .PARAMETER Sender
        The Form Control that fired the Resize Event
      .PARAMETER EventArg
        The Event Arguments for the Form Resize Event
      .EXAMPLE
        Start-ScriptInfoFormResize -Sender $Sender -EventArg $EventArg
      .NOTES
        Original Function By Ken Sweet
    #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.Form]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Resize Event for `$ScriptInfoForm"

    [MyConfig]::AutoExit = 0

    Write-Verbose -Message "Exit Resize Event for `$ScriptInfoForm"
  }
  #endregion ******** Function Start-ScriptInfoFormResize ********
  $ScriptInfoForm.add_Resize({ Start-ScriptInfoFormResize -Sender $This -EventArg $PSItem })

  #region ******** Function Start-ScriptInfoFormShown ********
  function Start-ScriptInfoFormShown
  {
  <#
    .SYNOPSIS
      Shown Event for the ScriptInfo Form Control
    .DESCRIPTION
      Shown Event for the ScriptInfo Form Control
    .PARAMETER Sender
       The Form Control that fired the Shown Event
    .PARAMETER EventArg
       The Event Arguments for the Form Shown Event
    .EXAMPLE
       Start-ScriptInfoFormShown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Shown Event for `$ScriptInfoForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    Start-ScriptInfoLeftToolStripItemClick -Sender ($ScriptInfoLeftMenuStrip.Items[$DefInfoTopic]) -EventArg $EventArg

    Write-Verbose -Message "Exit Shown Event for `$ScriptInfoForm"
  }
  #endregion ******** Function Start-ScriptInfoFormShown ********
  $ScriptInfoForm.add_Shown({ Start-ScriptInfoFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for ScriptInfo Form ********

  # ************************************************
  # ScriptInfo Panel
  # ************************************************
  #region $ScriptInfoPanel = [System.Windows.Forms.Panel]::New()
  $ScriptInfoPanel = [System.Windows.Forms.Panel]::New()
  $ScriptInfoForm.Controls.Add($ScriptInfoPanel)
  $ScriptInfoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ScriptInfoPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $ScriptInfoPanel.Name = "ScriptInfoPanel"
  #endregion $ScriptInfoPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ScriptInfoPanel Controls ********

  #region $ScriptInfoRichTextBox = [System.Windows.Forms.RichTextBox]::New()
  $ScriptInfoRichTextBox = [System.Windows.Forms.RichTextBox]::New()
  $ScriptInfoPanel.Controls.Add($ScriptInfoRichTextBox)
  $ScriptInfoRichTextBox.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Bottom, Right")
  $ScriptInfoRichTextBox.BackColor = [MyConfig]::Colors.TextBack
  $ScriptInfoRichTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
  $ScriptInfoRichTextBox.DetectUrls = $True
  $ScriptInfoRichTextBox.Font = [MyConfig]::Font.Regular
  $ScriptInfoRichTextBox.ForeColor = [MyConfig]::Colors.TextFore
  $ScriptInfoRichTextBox.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $ScriptInfoRichTextBox.MaxLength = [Int]::MaxValue
  $ScriptInfoRichTextBox.Multiline = $True
  $ScriptInfoRichTextBox.Name = "ScriptInfoRichTextBox"
  $ScriptInfoRichTextBox.ReadOnly = $True
  $ScriptInfoRichTextBox.Rtf = ""
  $ScriptInfoRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
  $ScriptInfoRichTextBox.Size = [System.Drawing.Size]::New(($ScriptInfoPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ($ScriptInfoPanel.ClientSize.Height - ($ScriptInfoRichTextBox.Top + [MyConfig]::FormSpacer)))
  $ScriptInfoRichTextBox.TabStop = $False
  $ScriptInfoRichTextBox.Text = ""
  $ScriptInfoRichTextBox.WordWrap = $True
  #endregion $ScriptInfoRichTextBox = [System.Windows.Forms.RichTextBox]::New()

  #endregion ******** $ScriptInfoPanel Controls ********

  # ************************************************
  # ScriptInfoLeft MenuStrip
  # ************************************************
  #region $ScriptInfoLeftMenuStrip = [System.Windows.Forms.MenuStrip]::New()
  $ScriptInfoLeftMenuStrip = [System.Windows.Forms.MenuStrip]::New()
  $ScriptInfoForm.Controls.Add($ScriptInfoLeftMenuStrip)
  $ScriptInfoForm.MainMenuStrip = $ScriptInfoLeftMenuStrip
  $ScriptInfoLeftMenuStrip.BackColor = [MyConfig]::Colors.Back
  $ScriptInfoLeftMenuStrip.Dock = [System.Windows.Forms.DockStyle]::Left
  $ScriptInfoLeftMenuStrip.Font = [MyConfig]::Font.Regular
  $ScriptInfoLeftMenuStrip.ForeColor = [MyConfig]::Colors.Fore
  $ScriptInfoLeftMenuStrip.ImageList = $FCGSmallImageList
  $ScriptInfoLeftMenuStrip.Name = "ScriptInfoLeftMenuStrip"
  $ScriptInfoLeftMenuStrip.ShowItemToolTips = $True
  $ScriptInfoLeftMenuStrip.Text = "ScriptInfoLeftMenuStrip"
  #endregion $ScriptInfoLeftMenuStrip = [System.Windows.Forms.MenuStrip]::New()

  #region ******** Function Start-ScriptInfoLeftToolStripItemClick ********
  function Start-ScriptInfoLeftToolStripItemClick
  {
  <#
    .SYNOPSIS
      Click Event for the ScriptInfoLeft ToolStripItem Control
    .DESCRIPTION
      Click Event for the ScriptInfoLeft ToolStripItem Control
    .PARAMETER Sender
       The ToolStripItem Control that fired the Click Event
    .PARAMETER EventArg
       The Event Arguments for the ToolStripItem Click Event
    .EXAMPLE
       Start-ScriptInfoLeftToolStripItemClick -Sender $Sender -EventArg $EventArg
    .NOTES
      Original Function By Ken Sweet
  #>
    [CmdletBinding()]
    param (
      [parameter(Mandatory = $True)]
      [System.Windows.Forms.ToolStripItem]$Sender,
      [parameter(Mandatory = $True)]
      [Object]$EventArg
    )
    Write-Verbose -Message "Enter Click Event for `$ScriptInfoLeftToolStripItem"

    [MyConfig]::AutoExit = 0

    $ScriptInfoBtmStatusStrip.Items["Status"].Text = "Showing: $($Sender.Text)"

    Switch ($Sender.Name)
    {
      "Exit"
      {
        $ScriptInfoForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        Break
      }
      Default
      {
        $ScriptInfoRichTextBox.Clear()
        $ScriptInfoRichTextBox.Beg
        Switch ($Sender.Tag.Type)
        {
          "None"
          {
            $ScriptInfoRichTextBox.Rtf = $Sender.Tag.Data
            Break
          }
          "Base64"
          {
            $ScriptInfoRichTextBox.Rtf = Encode-MyData -Data ($Sender.Tag.Data) -AsString -Decode
            Break
          }
          "Compress"
          {
            $ScriptInfoRichTextBox.Rtf = Compress-MyData -Data ($Sender.Tag.Data) -Decompress -AsString
            Break
          }
        }
        $ScriptInfoRichTextBox.SelectAll()
        $ScriptInfoRichTextBox.SelectionIndent = 10
        $ScriptInfoRichTextBox.SelectionLength = 0
        Break
      }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Verbose -Message "Exit Click Event for `$ScriptInfoLeftToolStripItem"
  }
  #endregion ******** Function Start-ScriptInfoLeftToolStripItemClick ********

  New-MenuSeparator -Menu $ScriptInfoLeftMenuStrip
  New-MenuLabel -Menu $ScriptInfoLeftMenuStrip -Text $InfoTitle -Name "Info Topics" -Tag "Info Topics" -Font ([MyConfig]::Font.Bold)
  New-MenuSeparator -Menu $ScriptInfoLeftMenuStrip

  forEach ($Key in $Topics.Keys)
  {
    (New-MenuItem -Menu $ScriptInfoLeftMenuStrip -Text ($Topics[$Key].Name) -Name $Key -Tag @{"Data" = $Topics[$Key].Data; "Type" = $Topics[$Key].Type} -Alignment "MiddleLeft" -DisplayStyle "ImageAndText" -ImageKey "HelpIcon" -PassThru).add_Click({ Start-ScriptInfoLeftToolStripItemClick -Sender $This -EventArg $PSItem })
  }

  New-MenuSeparator -Menu $ScriptInfoLeftMenuStrip
  (New-MenuItem -Menu $ScriptInfoLeftMenuStrip -Text "E&xit" -Name "Exit" -Tag "Exit" -Alignment "MiddleLeft" -DisplayStyle "ImageAndText" -ImageKey "ExitIcon" -PassThru).add_Click({ Start-ScriptInfoLeftToolStripItemClick -Sender $This -EventArg $PSItem })
  New-MenuSeparator -Menu $ScriptInfoLeftMenuStrip

  #region $ScriptInfoTopPanel = [System.Windows.Forms.Panel]::New()
  $ScriptInfoTopPanel = [System.Windows.Forms.Panel]::New()
  $ScriptInfoForm.Controls.Add($ScriptInfoTopPanel)
  $ScriptInfoTopPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ScriptInfoTopPanel.Dock = [System.Windows.Forms.DockStyle]::Top
  $ScriptInfoTopPanel.Name = "ScriptInfoTopPanel"
  #endregion $ScriptInfoTopPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ScriptInfoTopPanel Controls ********

  #region $ScriptInfoTopLabel = [System.Windows.Forms.Label]::New()
  $ScriptInfoTopLabel = [System.Windows.Forms.Label]::New()
  $ScriptInfoTopPanel.Controls.Add($ScriptInfoTopLabel)
  $ScriptInfoTopLabel.Anchor = [System.Windows.Forms.AnchorStyles]("Top, Left, Right")
  $ScriptInfoTopLabel.BackColor = [MyConfig]::Colors.TitleBack
  $ScriptInfoTopLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
  $ScriptInfoTopLabel.Font = [MyConfig]::Font.Title
  $ScriptInfoTopLabel.ForeColor = [MyConfig]::Colors.TitleFore
  $ScriptInfoTopLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $ScriptInfoTopLabel.Name = "ScriptInfoTopLabel"
  $ScriptInfoTopLabel.Text = $Title
  $ScriptInfoTopLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
  $ScriptInfoTopLabel.Size = [System.Drawing.Size]::New(($ScriptInfoTopPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), $ScriptInfoTopLabel.PreferredHeight)
  #endregion $ScriptInfoTopLabel = [System.Windows.Forms.Label]::New()

  $ScriptInfoTopPanel.ClientSize = [System.Drawing.Size]::New($ScriptInfoTopPanel.ClientSize.Width, ($ScriptInfoTopLabel.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ScriptInfoTopPanel Controls ********

  # ************************************************
  # ScriptInfoBtm StatusStrip
  # ************************************************
  #region $ScriptInfoBtmStatusStrip = [System.Windows.Forms.StatusStrip]::New()
  $ScriptInfoBtmStatusStrip = [System.Windows.Forms.StatusStrip]::New()
  $ScriptInfoForm.Controls.Add($ScriptInfoBtmStatusStrip)
  $ScriptInfoBtmStatusStrip.BackColor = [MyConfig]::Colors.Back
  $ScriptInfoBtmStatusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $ScriptInfoBtmStatusStrip.Font = [MyConfig]::Font.Regular
  $ScriptInfoBtmStatusStrip.ForeColor = [MyConfig]::Colors.Fore
  $ScriptInfoBtmStatusStrip.ImageList = $FCGSmallImageList
  $ScriptInfoBtmStatusStrip.Name = "ScriptInfoBtmStatusStrip"
  $ScriptInfoBtmStatusStrip.ShowItemToolTips = $True
  $ScriptInfoBtmStatusStrip.Text = "ScriptInfoBtmStatusStrip"
  #endregion $ScriptInfoBtmStatusStrip = [System.Windows.Forms.StatusStrip]::New()

  New-MenuLabel -Menu $ScriptInfoBtmStatusStrip -Text "Status" -Name "Status" -Tag "Status"

  #endregion ******** Controls for ScriptInfo Form ********

  #endregion ******** End **** ScriptInfo **** End ********

  [Void]$ScriptInfoForm.ShowDialog($FCGForm)

  $ScriptInfoForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Show-ScriptInfo"
}
#endregion function Show-ScriptInfo

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

# ---------------------------------------
# Sample Function Display Status Messages
# ---------------------------------------
#region function Sample-RichTextStatus
function Sample-RichTextStatus()
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
      Sample-RichTextStatus -RichTextBox $RichTextBox
    .EXAMPLE
      Sample-RichTextStatus -RichTextBox $RichTextBox -HashTable $HashTable
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $True)]
    [System.Windows.Forms.RichTextBox]$RichTextBox,
    [HashTable]$HashTable
  )
  Write-Verbose -Message "Enter Function Sample-RichTextStatus"

  $DisplayResult = [System.Windows.Forms.DialogResult]::OK
  $RichTextBox.Refresh()

  # Get Passed Values
  If ($HashTable.ContainsKey("ShowHeader"))
  {
    $ShowHeader = $HashTable.ShowHeader
  }
  Else
  {
    $ShowHeader = $True
  }

  # **************
  # RFT Formatting
  # **************
  # Permanate till Changed
  #$RichTextBox.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Left
  #$RichTextBox.SelectionBullet = $True
  #$RichTextBox.SelectionIndent = 10
  # Resets After AppendText
  #$RichTextBox.SelectionBackColor = [MyConfig]::Colors.TextBack
  #$RichTextBox.SelectionCharOffset = 0
  #$RichTextBox.SelectionColor = [MyConfig]::Colors.TextFore
  #$RichTextBox.SelectionFont = [MyConfig]::Font.Bold
  # **********************
  # Update RichTextBox Text...
  # **********************

  $RichTextBox.SelectionIndent = 10
  $RichTextBox.SelectionBullet = $False

  # Write KPI Event
  #Write-KPIEvent -Source "Utility" -EntryType "Information" -EventID 0 -Category 0 -Message "Some Unknown KPI Event"

  if ($ShowHeader)
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
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Started Proccess List Data Here..." -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextTitle)
  $RichTextBox.SelectionIndent = 20
  $RichTextBox.SelectionBullet = $True

  :UserCancel foreach ($Key in $HashTable.Keys)
  {
    Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Found Key" -TextFore ([MyConfig]::Colors.TextInfo) -Value "$($Key) = $($HashTable[$Key])" -ValueFore ([MyConfig]::Colors.TextGood)
    # Check for Fast Exit
    [System.Windows.Forms.Application]::DoEvents()
    If ($RichTextBox.Parent.Parent.Tag.Cancel)
    {
      $RichTextBox.SelectionIndent = 10
      $RichTextBox.SelectionBullet = $False
      Write-RichTextBox -RichTextBox $RichTextBox
      Write-RichTextBox -RichTextBox $RichTextBox -Text "Exiting - User Canceled" -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextBad) -Alignment Center
      $DisplayResult = [System.Windows.Forms.DialogResult]::Abort
      Break UserCancel
    }
    # Pause Processing Loop
    If ($RichTextBox.Parent.Parent.Tag.Pause)
    {
      $TmpPause = $RichTextBox.SelectionBullet
      $TmpTitle = $RichTextBox.Parent.Parent.Text
      $RichTextBox.Parent.Parent.Text = "$($TmpTitle) - PAUSED!"
      $RichTextBox.SelectionBullet = $False
      While ($RichTextBox.Parent.Parent.Tag.Pause)
      {
        [System.Threading.Thread]::Sleep(100)
        [System.Windows.Forms.Application]::DoEvents()
      }
      $RichTextBox.SelectionBullet = $TmpPause
      $RichTextBox.Parent.Parent.Text = $TmpTitle
    }
    Start-Sleep -Milliseconds 100
  }

  # Pause Before Deployment
  $RichTextBox.Parent.Parent.Tag.Pause = $True
  $TmpPause = $RichTextBox.SelectionBullet
  $TmpTitle = $RichTextBox.Parent.Parent.Text
  $RichTextBox.Parent.Parent.Text = "$($TmpTitle) - PAUSED!"
  $RichTextBox.SelectionBullet = $False

  Write-RichTextBox -RichTextBox $RichTextBox
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Pause to Review Status" -Font ([MyConfig]::Font.Bold) -Alignment Center
  Write-RichTextBox -RichTextBox $RichTextBox
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Press 'Pause' to Continue with the Current Deployment" -Alignment Center
  Write-RichTextBox -RichTextBox $RichTextBox -Text "or Ctrl-Alt-Backspace to Exit / Cancel" -Alignment Center
  Write-RichTextBox -RichTextBox $RichTextBox

  While ($RichTextBox.Parent.Parent.Tag.Pause)
  {
    [System.Threading.Thread]::Sleep(100)
    [System.Windows.Forms.Application]::DoEvents()
    If ($RichTextBox.Parent.Parent.Tag.Cancel)
    {
      $RichTextBox.Parent.Parent.Tag.Pause = $False
      $RichTextBox.SelectionIndent = 10
      $RichTextBox.SelectionBullet = $False
      Write-RichTextBox -RichTextBox $RichTextBox
      Write-RichTextBox -RichTextBox $RichTextBox -Text "Exiting - User Canceled" -Font ([MyConfig]::Font.Bold) -TextFore ([MyConfig]::Colors.TextBad) -Alignment Center
      $DisplayResult = [System.Windows.Forms.DialogResult]::Abort
    }
  }
  $RichTextBox.SelectionBullet = $TmpPause
  $RichTextBox.Parent.Parent.Text = $TmpTitle

  # Display an Error Information
  $RichTextBox.SelectionIndent = 10
  $RichTextBox.SelectionBullet = $False
  Write-RichTextBox -RichTextBox $RichTextBox
  Write-RichTextBox -RichTextBox $RichTextBox -Text "Show Fake Error Message" -TextFore ([MyConfig]::Colors.TextWarn) -Font ([MyConfig]::Font.Bold)
  $RichTextBox.SelectionIndent = 20
  $RichTextBox.SelectionBullet = $True
  Try
  {
    Throw "This is a Fake Error!"
  }
  Catch
  {
    # Write Error to Status Dialog
    Write-RichTextBoxError -RichTextBox $RichTextBox
  }

  if ($ShowHeader)
  {
    $RichTextBox.SelectionIndent = 10
    $RichTextBox.SelectionBullet = $False
    Write-RichTextBox -RichTextBox $RichTextBox

    # Set Final Status Message
    Switch ($DisplayResult)
    {
      "OK"
      {
        $FinalMsg = "Add Success Message Here!"
        $FinalClr = [MyConfig]::Colors.TextGood
        Break
      }
      "Cancel"
      {
        $FinalMsg = "Add Error Message Here!"
        $FinalClr = [MyConfig]::Colors.TextBad
        Break
      }
      "Abort"
      {
        $FinalMsg = "Add Abort Message Here!"
        $FinalClr = [MyConfig]::Colors.TextWarn
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

  Write-Verbose -Message "Exit Function Sample-RichTextStatus"
}
#endregion function Sample-RichTextStatus

#$HashTable = @{"ShowHeader" = $True}
#$ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Sample-RichTextStatus -RichTextBox $RichTextBox -HashTable $HashTable }
#$DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title "Initializing $([MyConfig]::ScriptName)" -ButtonMid "OK" -HashTable $HashTable -AllowControl

# ---------------------------
# Show ProgressBarStatus Function
# ---------------------------
#region ProgressBarStatus Result Class
Class ProgressBarStatus
{
  [Bool]$Success
  [Object]$DialogResult
  ProgressBarStatus ([Bool]$Success, [Object]$DialogResult)
  {
    $This.Success = $Success
    $This.DialogResult = $DialogResult
  }
}
#endregion ProgressBarStatus Result Class

#region function Show-ProgressBarStatus
Function Show-ProgressBarStatus ()
{
  <#
    .SYNOPSIS
      Shows Show-ProgressBarStatus
    .DESCRIPTION
      Shows Show-ProgressBarStatus
    .PARAMETER Title
      Title of the Show-ProgressBarStatus Dialog Window
    .PARAMETER ScriptBlock
      Script Block to Execure
    .PARAMETER HashTable
      HashTable of Paramerts to Pass to the ScriptBlock
    .PARAMETER Width
      Width of the Show-ProgressBarStatus Dialog Window
    .PARAMETER AllowControl
      Enable Pause and Break out of Script Block
    .EXAMPLE
      $HashTable = @{"Values" = @(([System.Globalization.DateTimeFormatInfo]::New()).MonthNames)[0..11]}
      $ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.ProgressBar]$ProgressBar, [System.Windows.Forms.Label]$Label) Sample-ProgressBarStatus -ProgressBar $ProgressBar -Label $Label }
      $DialogResult = Show-ProgressBarStatus -ScriptBlock $ScriptBlock -Title "Initializing $([MyConfig]::ScriptName)"
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
    [parameter(Mandatory = $True)]
    [ScriptBlock]$ScriptBlock = { },
    [HashTable]$HashTable = @{ },
    [Int]$Width = 45,
    [Switch]$AllowControl
  )
  Write-Verbose -Message "Enter Function Show-ProgressBarStatus"

  #region ******** Begin **** $ProgressBarStatus **** Begin ********

  # ************************************************
  # $ProgressBarStatus Form
  # ************************************************
  #region $ProgressBarStatusForm = [System.Windows.Forms.Form]::New()
  $ProgressBarStatusForm = [System.Windows.Forms.Form]::New()
  $ProgressBarStatusForm.BackColor = [MyConfig]::Colors.Back
  $ProgressBarStatusForm.Font = [MyConfig]::Font.Regular
  $ProgressBarStatusForm.ForeColor = [MyConfig]::Colors.Fore
  $ProgressBarStatusForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
  $ProgressBarStatusForm.Icon = $PILForm.Icon
  $ProgressBarStatusForm.KeyPreview = $AllowControl.IsPresent
  $ProgressBarStatusForm.MaximizeBox = $False
  $ProgressBarStatusForm.MinimizeBox = $False
  $ProgressBarStatusForm.MinimumSize = [System.Drawing.Size]::New(([MyConfig]::Font.Width * $Width), ([MyConfig]::Font.Height * $Height))
  $ProgressBarStatusForm.Name = "ProgressBarStatusForm"
  $ProgressBarStatusForm.Owner = $PILForm
  $ProgressBarStatusForm.ShowInTaskbar = $False
  $ProgressBarStatusForm.Size = $ProgressBarStatusForm.MinimumSize
  $ProgressBarStatusForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $ProgressBarStatusForm.Tag = @{ "Cancel" = $False; "Pause" = $False; "Finished" = $True }
  $ProgressBarStatusForm.Text = $Title
  #endregion $ProgressBarStatusForm = [System.Windows.Forms.Form]::New()

  #region ******** Function Start-ProgressBarStatusFormKeyDown ********
  Function Start-ProgressBarStatusFormKeyDown
  {
  <#
    .SYNOPSIS
      KeyDown Event for the ProgressBarStatus Form Control
    .DESCRIPTION
      KeyDown Event for the ProgressBarStatus Form Control
    .PARAMETER Sender
       The Form Control that fired the KeyDown Event
    .PARAMETER EventArg
       The Event Arguments for the Form KeyDown Event
    .EXAMPLE
       Start-ProgressBarStatusFormKeyDown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter KeyDown Event for `$ProgressBarStatusForm"

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
            $Sender.DialogResult = [System.Windows.Forms.DialogResult]::OK
          }
          Break
        }
      }
    }

    Write-Verbose -Message "Exit KeyDown Event for `$ProgressBarStatusForm"
  }
  #endregion ******** Function Start-ProgressBarStatusFormKeyDown ********
  If ($AllowControl.IsPresent)
  {
    $ProgressBarStatusForm.add_KeyDown({ Start-ProgressBarStatusFormKeyDown -Sender $This -EventArg $PSItem })
  }

  #region ******** Function Start-ProgressBarStatusFormShown ********
  Function Start-ProgressBarStatusFormShown
  {
    <#
      .SYNOPSIS
        Shown Event for the $ProgressBarStatus Form Control
      .DESCRIPTION
        Shown Event for the $ProgressBarStatus Form Control
      .PARAMETER Sender
         The Form Control that fired the Shown Event
      .PARAMETER EventArg
         The Event Arguments for the Form Shown Event
      .EXAMPLE
         Start-ProgressBarStatusFormShown -Sender $Sender -EventArg $EventArg
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
    Write-Verbose -Message "Enter Shown Event for `$ProgressBarStatusForm"

    [MyConfig]::AutoExit = 0

    $Sender.Refresh()

    If ([MyConfig]::Production)
    {
      # Disable Auto Exit Timer
      $PILTimer.Enabled = $False
    }
    
    if ($PassHashTable)
    {
      $DialogResult = Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $ProgressBarStatusProgressBar, $ProgressBarStatusLabel, $HashTable
    }
    else
    {
      $DialogResult = Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $ProgressBarStatusProgressBar, $ProgressBarStatusLabel
    }
    
    $Sender.Tag.Finished = $True
    
    If ([MyConfig]::Production)
    {
      # Re-enable Auto Exit Timer
      $PILTimer.Enabled = ([MyConfig]::AutoExitMax -gt 0)
    }

    $ProgressBarStatusForm.DialogResult = $DialogResult

    Write-Verbose -Message "Exit Shown Event for `$ProgressBarStatusForm"
  }
  #endregion ******** Function Start-ProgressBarStatusFormShown ********
  $ProgressBarStatusForm.add_Shown({ Start-ProgressBarStatusFormShown -Sender $This -EventArg $PSItem })

  #region ******** Controls for $ProgressBarStatus Form ********

  # ************************************************
  # $ProgressBarStatus Panel
  # ************************************************
  #region $ProgressBarStatusPanel = [System.Windows.Forms.Panel]::New()
  $ProgressBarStatusPanel = [System.Windows.Forms.Panel]::New()
  $ProgressBarStatusForm.Controls.Add($ProgressBarStatusPanel)
  $ProgressBarStatusPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $ProgressBarStatusPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
  $ProgressBarStatusPanel.Name = "ProgressBarStatusPanel"
  #endregion $ProgressBarStatusPanel = [System.Windows.Forms.Panel]::New()

  #region ******** $ProgressBarStatusPanel Controls ********

  #region $ProgressBarStatusLabel = [System.Windows.Forms.Label]::New()
  $ProgressBarStatusLabel = [System.Windows.Forms.Label]::New()
  $ProgressBarStatusPanel.Controls.Add($ProgressBarStatusLabel)
  $ProgressBarStatusLabel.Font = [MyConfig]::Font.Bold
  $ProgressBarStatusLabel.ForeColor = [MyConfig]::Colors.LabelFore
  $ProgressBarStatusLabel.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, [MyConfig]::FormSpacer)
  $ProgressBarStatusLabel.Name = "ProgressBarStatusLabel"
  $ProgressBarStatusLabel.Size = [System.Drawing.Size]::New(($ProgressBarStatusPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2)), ([MyConfig]::Font.Height * 2))
  $ProgressBarStatusLabel.Text = $Null
  $ProgressBarStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
  #endregion $ProgressBarStatusLabel = [System.Windows.Forms.Label]::New()

  #region $ProgressBarStatusProgressBar = [System.Windows.Forms.ProgressBar]::New()
  $ProgressBarStatusProgressBar = [System.Windows.Forms.ProgressBar]::New()
  $ProgressBarStatusPanel.Controls.Add($ProgressBarStatusProgressBar)
  #$ProgressBarStatusProgressBar.AutoSize = $False
  $ProgressBarStatusProgressBar.BackColor = [MyConfig]::Colors.Back
  #$ProgressBarStatusProgressBar.Enabled = $True
  $ProgressBarStatusProgressBar.Font = [MyConfig]::Font.Regular
  $ProgressBarStatusProgressBar.ForeColor = [MyConfig]::Colors.Fore
  $ProgressBarStatusProgressBar.Location = [System.Drawing.Point]::New([MyConfig]::FormSpacer, ($ProgressBarStatusLabel.Bottom + [MyConfig]::FormSpacer))
  $ProgressBarStatusProgressBar.Name = "ProgressBarStatusProgressBar"
  $ProgressBarStatusProgressBar.TabStop = $False
  #$ProgressBarStatusProgressBar.Tag = [System.Object]::New()
  #$ProgressBarStatusProgressBar.Value = 0
  #$ProgressBarStatusProgressBar.Visible = $True
  $ProgressBarStatusProgressBar.Width = ($ProgressBarStatusPanel.ClientSize.Width - ([MyConfig]::FormSpacer * 2))
  #endregion $ProgressBarStatusProgressBar = [System.Windows.Forms.ProgressBar]::New()

  $ProgressBarStatusPanel.ClientSize = [System.Drawing.Size]::New($ProgressBarStatusPanel.ClientSize.Width, ($ProgressBarStatusProgressBar.Bottom + [MyConfig]::FormSpacer))

  #endregion ******** $ProgressBarStatusPanel Controls ********

  $ProgressBarStatusForm.ClientSize = [System.Drawing.Size]::New($ProgressBarStatusForm.ClientSize.Width, $ProgressBarStatusPanel.ClientSize.Height)

  #endregion ******** Controls for $ProgressBarStatus Form ********

  #endregion ******** End **** $Show-ProgressBarStatus **** End ********
  
  $PassHashTable = $PSBoundParameters.ContainsKey("HashTable")
  $DialogResult = $ProgressBarStatusForm.ShowDialog($PILForm)
  [ProgressBarStatus]::New(($DialogResult -eq [System.Windows.Forms.DialogResult]::OK), $DialogResult)

  $ProgressBarStatusForm.Dispose()

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  Write-Verbose -Message "Exit Function Show-ProgressBarStatus"
}
#endregion function Show-ProgressBarStatus

# ---------------------------------------
# Sample Function Display Status Messages
# ---------------------------------------
#region function Sample-ProgressBarStatus
Function Sample-ProgressBarStatus()
{
  <#
    .SYNOPSIS
      Display Utility Status Sample Function
    .DESCRIPTION
      Display Utility Status Sample Function
    .PARAMETER ProgressBar
      The Progress Bar
    .PARAMETER Label
      The Label to Indicate the Current Item being Proccessed
    .PARAMETER HashTable
      Passed Paramters HashTable
    .EXAMPLE
      Sample-ProgressBarStatus -ProgressBar $ProgressBar -Label $Label
    .NOTES
      Original Script By Ken Sweet
    .LINK
  #>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory = $True)]
    [System.Windows.Forms.ProgressBar]$ProgressBar,
    [Parameter(Mandatory = $True)]
    [System.Windows.Forms.Label]$Label,
    [HashTable]$HashTable
  )
  Write-Verbose -Message "Enter Function Sample-ProgressBarStatus"

  $DisplayResult = [System.Windows.Forms.DialogResult]::OK
  $ProgressBar.Refresh()

  # Write KPI Event
  #Write-KPIEvent -Source "Utility" -EntryType "Information" -EventID 0 -Category 0 -Message "Some Unknown KPI Event"

  # Update Status Message
  $PILBtmStatusStrip.Items["Status"].Text = $ProgressBar.Parent.Parent.Text

  # Month Names
  $Values = $HashTable.Values

  # Set Starting ProgresBar Values
  $ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
  $ProgressBar.Maximum = $Values.Count
  $ProgressBar.Minimum = 1
  $ProgressBar.Step = 1
  $ProgressBar.Value = 1

  :UserCancel ForEach ($Value In $Values)
  {
    # Update Progress Information

    $Label.Text = $Value
    $Label.Refresh()

    # Check for Fast Exit
    [System.Windows.Forms.Application]::DoEvents()
    If ($ProgressBar.Parent.Parent.Tag.Cancel)
    {
      $DisplayResult = [System.Windows.Forms.DialogResult]::Abort
      Break UserCancel
    }

    # Pause Processing Loop
    If ($ProgressBar.Parent.Parent.Tag.Pause)
    {
      $TmpTitle = $ProgressBar.Parent.Parent.Text
      $ProgressBar.Parent.Parent.Text = "$($TmpTitle) - PAUSED!"
      While ($ProgressBar.Parent.Parent.Tag.Pause)
      {
        [System.Threading.Thread]::Sleep(100)
        [System.Windows.Forms.Application]::DoEvents()
      }
      $ProgressBar.Parent.Parent.Text = $TmpTitle
    }

    $ProgressBar.Increment(1)
    $ProgressBar.Refresh()
    Start-Sleep -Milliseconds 1000
  }

  # Update Status Message
  $PILBtmStatusStrip.Items["Status"].Text = "Completed $($ProgressBar.Parent.Parent.Text)"

  # Return DialogResult
  $DisplayResult
  $DisplayResult = $Null

  Write-Verbose -Message "Exit Function Sample-ProgressBarStatus"
}
#endregion function Sample-ProgressBarStatus

#$HashTable = @{"Values" = @(([System.Globalization.DateTimeFormatInfo]::New()).MonthNames)[0..11]}
#$ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.ProgressBar]$ProgressBar, [System.Windows.Forms.Label]$Label) Sample-ProgressBarStatus -ProgressBar $ProgressBar -Label $Label }
#$DialogResult = Show-ProgressBarStatus -ScriptBlock $ScriptBlock -Title "Initializing $([MyConfig]::ScriptName)"


#endregion ******** PIL Common Dialogs ********

#region ******** PIL Custom Commands ********

#region function Get-ModuleList
Function Get-ModuleList ()
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
  Param (
    [parameter(Mandatory = $False)]
    [ValidateSet("All Users", "Current User")]
    [String]$Location = "All Users",
    [parameter(Mandatory = $True)]
    [String]$Path
  )
  Write-Verbose -Message "Enter Function $($MyInvocation.MyCommand)"
  
  # Get Installed Modules
  $TmpModList = Get-ChildItem -Path $Path
  ForEach ($TmpModItem In $TmpModList)
  {
    # get Module Versions
    $TmpVersions = @(Get-ChildItem -Path $TmpModItem.FullName | Where-Object -FilterScript { $PSItem.Name -match "\d+\.\d+\.\d+" } | Sort-Object -Property Name -Descending | Select-Object -First 1)
    If ($TmpVersions.Count -eq 0)
    {
      If (-not [MyRuntime]::Modules.ContainsKey($TmpModItem.Name))
      {
        # Custom Module
        [MyRuntime]::Modules.Add($TmpModItem.Name, [PILModule]::New($Location, $TmpModItem.Name, "0.0.0"))
      }
    }
    Else
    {
      If (-not [MyRuntime]::Modules.ContainsKey($TmpModItem.Name))
      {
        # Installed Module
        ForEach ($TmpVersion In $TmpVersions)
        {
          [MyRuntime]::Modules.Add($TmpModItem.Name, [PILModule]::New($Location, $TmpModItem.Name, $TmpVersion.Name))
        }
      }
    }
  }
  
  Write-Verbose -Message "Exit Function $($MyInvocation.MyCommand)"
}
#endregion function Get-ModuleList

#endregion ******** PIL Custom Commands ********

#region ******** PIL Custom Dialogs ********

# ----------------------------
# Sample Initiliaze PILUtility
# ----------------------------
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
  $ExportFile = $HashTable.ExportFile
  
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
    $DialogResult = Load-PILConfigFIle -RichTextBox $RichTextBox -HashTable $HashTable
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


#region function Load-PILConfigFIle
Function Load-PILConfigFIle()
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
      Load-PILConfigFIle -RichTextBox $RichTextBox
    .EXAMPLE
      Load-PILConfigFIle -RichTextBox $RichTextBox -HashTable $HashTable
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
  Write-Verbose -Message "Enter Function Load-PILConfigFIle"
  
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
      
      # Add / Update PIL Columns
      $RichTextBox.SelectionIndent = 20
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Number of Columns" -Value ($TmpConfig.ColumnNames.Count)
      [MyRuntime]::UpdateTotalColumn($TmpConfig.ColumnNames.Count)
      $RichTextBox.SelectionIndent = 30
      $PILItemListListView.BeginUpdate()
      $PILItemListListView.Columns.Clear()
      $PILItemListListView.Items.Clear()
      [MyRuntime]::ThreadConfig.ColumnNames = $TmpConfig.ColumnNames
      For ($I = 0; $I -lt ([MyRuntime]::MaxColumns); $I++)
      {
        New-ColumnHeader -ListView $PILItemListListView -Text ([MyRuntime]::ThreadConfig.ColumnNames[$I]) -Name ([MyRuntime]::ThreadConfig.ColumnNames[$I]) -Tag ([MyRuntime]::ThreadConfig.ColumnNames[$I]) -Width -2
      }
      $PILItemListListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
      New-ColumnHeader -ListView $PILItemListListView -Text " " -Name "Blank" -Tag " " -Width ($PILForm.Width * 4)
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
            $DialogResult = Get-UserResponse -Title "Incorrect Module Version" -Message "The Module $($Module.Name) Version $($Module.Version) was not Found would you like to Install it to $($TmpInallMsg)?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
            If ($DialogResult.Success)
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
          $DialogResult = Get-UserResponse -Title "Module Not Instaled" -Message "The Module $($Module.Name) Version $($Module.Version) was not Found would you like to Install it to $($TmpInallMsg)?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
          If ($DialogResult.Success)
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
    Catch
    {
      Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "ERROR" -TextFore ([MyConfig]::Colors.TextBad) -Value "PIL Config File was not Loaded" -ValueFore ([MyConfig]::Colors.TextFore)
      Write-RichTextBoxError -RichTextBox $RichTextBox
      $DisplayResult = [System.Windows.Forms.DialogResult]::Cancel
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
  
  Write-Verbose -Message "Exit Function Load-PILConfigFIle"
}
#endregion function Load-PILConfigFIle


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
  $ExportFile = $HashTable.ExportFile
  
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
  
  Write-RichTextBoxValue -RichTextBox $RichTextBox -Text "Data Export File File" -Value ([System.IO.Path]::GetFileName($ExportFile)) -Font ([MyConfig]::Font.Bold)
  $RichTextBox.SelectionIndent = 30
  
  If ([System.IO.File]::Exists($ExportFile))
  {
    Try
    {
      # Get Column Names
      $TmpColNames = @($PILItemListListView.Columns | Select-Object -ExpandProperty Text)
      
      # Load Configuration
      $TmpExport = Import-Csv -LiteralPath $ExportFile
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
        ForEach ($TmpDataItem In $TmpDataList)
        {
          $TmpName = $TmpDataItem."$($ChkColumns[0])"
          If (-not $PILItemListListView.Items.ContainsKey($TmpName))
          {
            $TmpDataItem.FakeColumn = ""
            ($PILItemListListView.Items.Add([System.Windows.Forms.ListViewItem]::New(@($TmpDataItem.PSObject.Properties | Select-Object -ExpandProperty Value), "StatusInfo16Icon", [MyConfig]::Colors.TextFore, [MyConfig]::Colors.TextBack, [MyConfig]::Font.Regular))).Name = $TmpName
          }
        }
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
    [String]$Title = "Update PIL Threads Configuration",
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
      }
      Else
      {
        $PILTCFunctionsContextMenuStrip.Items["Remove"].Enabled = $False
      }
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
          $AST = [System.Management.Automation.Language.Parser]::ParseInput($TmpFunctions, [ref]$Null, [ref]$Null)
          $Functions = @($AST.FindAll({ Param ($Node) ($Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and (-not ($node.Parent -is [System.Management.Automation.Language.FunctionMemberAst]))) }, $True))
          If ($Functions.Count -gt 0)
          {
            ForEach ($Function In $Functions)
            {
              [Void]$PILTCFunctionsListBox.Items.Add([PILFunction]::New($Function.Name, $Function.Body.Extent.Text))
            }
          }
        }
        Break
      }
      "Remove"
      {
        $PILTCFunctionsListBox.Items.RemoveAt($PILTCFunctionsListBox.SelectedIndex)
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
  (New-MenuItem -Menu $PILTCFunctionsContextMenuStrip -Text "Clear All" -Name "Clear" -Tag "Clear" -DisplayStyle "ImageAndText" -ImageKey "Clear16Icon" -PassThru).add_Click({ Start-PILTCFunctionsContextMenuStripItemClick -Sender $This -EventArg $PSItem})

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
    
    Switch ($Sender.Name)
    {
      "Add"
      {
        $OrderedItems = [Ordered]@{ "Variable Name"= ""; "Variable Value" = "" }
        $DialogResult = Get-MultiTextBoxInput -Title "Add Variable" -Message "Show this Sample Message Prompt to the User" -OrderedItems $OrderedItems -AllRequired
        If ($DialogResult.Success)
        {
          [Void]$PILTCVariablesListBox.Items.Add([PILVariable]::New($DialogResult.OrderedItems["Variable Name"], $DialogResult.OrderedItems["Variable Value"]))
        }
        Break
      }
      "Edit"
      {
        $OrderedItems = [Ordered]@{ "Variable Name" = $PILTCVariablesListBox.SelectedItem.Name; "Variable Value" = $PILTCVariablesListBox.SelectedItem.Value }
        $DialogResult = Get-MultiTextBoxInput -Title "Edit Variable" -Message "Show this Sample Message Prompt to the User" -OrderedItems $OrderedItems -AllRequired
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
  (New-MenuItem -Menu $PILTCVariablesContextMenuStrip -Text "Clear All" -Name "Clear" -Tag "Clear" -DisplayStyle "ImageAndText" -ImageKey "Clear16Icon" -PassThru).add_Click({Start-PILTCVariablesContextMenuStripItemClick -Sender $This -EventArg $PSItem})

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
    
    Switch ($Sender.Name)
    {
      "Add"
      {
        $TmpCurMods = @($PILTCModulesListBox.Items | Select-Object -ExpandProperty Name)
        $TmpNewMods = @([MyRuntime]::Modules.Values | Where-Object { $PSItem.Name -notin $TmpCurMods } | Sort-Object -Property Location, Name)
        If ($TmpNewMods.Count -eq 0)
        {
          $DialogResult = Get-UserResponse -Title "No More Modules" -Message "No New Modules are Avaible for to Add to the PIL Thread Configuration." -ButtonMid OK -ButtonDefault OK -Icon ([System.Drawing.SystemIcons]::Information)
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
  (New-MenuItem -Menu $PILTCModulesContextMenuStrip -Text "Clear All" -Name "Clear" -Tag "Clear" -DisplayStyle "ImageAndText" -ImageKey "Clear16Icon" -PassThru).add_Click({ Start-PILTCModulesContextMenuStripItemClick -Sender $This -EventArg $PSItem})
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
    
    $PILTCScriptContextMenuStrip.Items["Clear"].Enabled = ($PILTCScriptTextBox.Text.Length -gt 0)
    $PILTCScriptContextMenuStrip.Show($Sender, $EventArg.Location)
    
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
        }
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
  (New-MenuItem -Menu $PILTCScriptContextMenuStrip -Text "Clear Script" -Name "Clear" -Tag "Clear" -DisplayStyle "ImageAndText" -ImageKey "Delete16Icon" -PassThru).add_Click({Start-PILTCScriptContextMenuStripItemClick -Sender $This -EventArg $PSItem})
  
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

#endregion ******** PIL Custom Dialogs ********

#region ******** Begin **** PIL **** Begin ********

#$Result = [System.Windows.Forms.MessageBox]::Show($PILForm, "Message Text", [MyConfig]::ScriptName, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

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

#region ******** $Play64Icon ********
$Play64Icon = @"
AAABAAEAQEAAAAEAIAAoQgAAFgAAACgAAABAAAAAgAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAA2AAAAagAAAJYAAAC8AAAA2QAAAO4AAAD+AAAA/wAAAP8AAAD9AAAA7QAAANgAAAC5AAAAkwAAAGUAAAAwAAAABgAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVAAAAYgAAALEAAADtAAAA/gAAAP4AAAD/AAAA/gAAAP8DAgD+BwQA/wkFAf4JBQH/BgQA/gMCAP8AAAD+AAAA/wAAAP4AAAD/AAAA/QAA
AOkAAACoAAAAWgAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAADwAAACoAAAA9AAAAP8AAAD/AQAA/xYOAv85JAf/VzcK/25FDf+BUQ//i1cR/5BaEf+TWxL/kloS/45XEf+IUxH/fEwP/2lA
DP9SMQn/NB8G/xQMAv8AAAD/AAAA/wAAAP8AAADvAAAAnwAAADMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARwAAAMQAAAD/AAAA/gIBAP8pGgX+XjwM/4pZEf6dZRT/nWQT/pxjE/+bYhP+mmET/5phEv6ZYBP/mF8S/phe
E/+XXRL+llwT/5ZcEv6VWxL/lFoS/pRZEv+TWRH+f0wP/1QyCv4kFQT/AgEA/gAAAP8AAAD+AAAAuwAAAD4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArAAAAuAAAAP0AAAD/DQgB/00yCv+MWxL/oGgU/59nFP+fZhT/nmYU/51lFP+dZBT/nGMT/5ti
E/+aYRP/mmET/5lgE/+ZXxP/mF4T/5ddE/+WXBP/llwT/5VbEv+UWhL/lFkS/5NZEv+SWBH/kVYR/31KDv9CJwf/CgYB/wAAAP8AAAD8AAAArQAAACMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAB9AAAA9QAAAP4LBwH/VjkL/pllFP+iahX+oWoV/6FpFP6gaBX/n2cU/p9m
FP+eZhP+nWUU/51kE/6cYxP/m2IT/pphE/+aYRL+mWAT/5hfEv6YXhP/l10S/pZcE/+WXBL+lVsS/5RaEv6UWRL/k1kR/pJYEf+RVxH+kVYR/4ZPD/5JKwj/CAUB/gAAAP8AAADxAAAAcAAAAAMAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB4AAADCAAAA/wEBAP8/Kgj/l2QU/6RtFv+kbBb/o2sW/6Jq
Ff+hahX/oWkV/6BoFf+gZxT/n2YU/55mFP+dZBP/m2IT/5lgE/+YXhL/llwS/5ZcEv+WXBL/ll0S/5ddEv+XXRL/llwT/5ZcE/+VWxL/lFoS/5RZEv+TWRL/klgR/5JXEf+RVhH/kFUR/4JMD/81Hgb/AQAA/wAA
AP8AAAC5AAAAGQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD0AAADmAAAA/xMNAv55URD/pm8W/qZv
Fv+lbhb+pG0W/6RsFf6jaxb/omoV/qFqFf+gaBT+nWQU/5heEv6TWRH/kFUR/o9UEf+PVBD+j1QQ/49UEP6PVBD/j1QQ/o9UEP+PVBH+kFUR/5FWEf6TWRH/lFoS/pVbEv+UWhL+lFkS/5NZEf6SWBH/kVcR/pFW
Ef+QVRD+j1QQ/2U7C/4PCQH/AAAA/gAAAOEAAAA2AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFAAAAD0AAAA/y4f
Bv+aaBX/qHEX/6dwF/+mbxf/pm8W/6VuFv+kbRb/pGwW/6FpFf+aYRP/k1kR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+SVxH/k1kR/5Ra
Ev+UWRL/k1kS/5JYEf+SVxH/kVYR/5BVEf+PVBH/gUwP/yUVBP8AAAD/AAAA8QAAAEkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AFkAAAD6AAAA/0cwCf6mcRf/qXMX/qhyF/+ocRb+p3AX/6ZvFv6mbxb/pGwV/pthE/+UWRL+k1gS/5NYEf6TWBL/k1gR/pNYEv+TWBH+k1gS/5NYEf6TWBL/k1gR/pNYEv+TWBH+k1gS/5NYEf6TWBL/k1gR/pNY
Ev+TWBH+k1gS/5NYEf6TWBL/lFkR/pRZEv+TWRH+klgR/5FXEf6RVhH/kFUQ/o9UEf+LUhD+OSEG/wAAAP4AAAD3AAAAUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAE0AAAD3AAAA/1U6DP+qdBf/q3QY/6pzGP+pcxf/qHIX/6hxF/+ncBb/n2cU/5ZcEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5Ra
Ev+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5NZEv+SWBH/klcR/5FWEf+QVRH/j1QR/45TEP9DJwf/AAAA/wAAAPYAAABFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADIAAADwAAAA/1I4C/6sdhj/rHYY/qt1GP+qdBf+qnMY/6lzF/6ncBb/m2IU/pZcE/+WXBL+llwT/5ZcEv6WXBP/llwS/pZcE/+WXBL+llwT/5ZcEv6WXBP/llwS/pZc
E/+WXBL+llwT/5ZcEv6WXBP/llwS/pZcE/+WXBL+llwT/5ZcEv6WXBP/llwS/pZcE/+WXBL+llwT/5ZcEv6VWxL/k1kS/pJYEf+RVxH+kVYR/5BVEP6PVBH/jlMQ/kEmB/8AAAD+AAAA7gAAACwAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIAAADdAAAA/0EtCf+tdxj/rXgZ/613Gf+sdhj/q3UY/6p0GP+ncBf/mmET/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5he
E/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5deEv+TWRL/klgR/5JXEf+RVhH/kFUR/49UEf+NUxD/NB4G/wAA
AP8AAADZAAAADwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAACrAAAA/ycbBf6pdhj/r3kZ/q55Gf+teBj+rXcZ/6x2GP6ncRb/m2IT/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppg
E/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mWAS/5RaEv6SWBH/kVcR/pFW
Ef+QVRD+j1QR/4lQEP4fEgP/AAAA/gAAAKUAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABaAAAA/QsIAf+Zaxb/sHsa/696Gf+veRn/rnkZ/614Gf+pcxf/nGMT/5tiE/+bYhP/m2IT/5ti
E/+bYhP/m2IT/5tiE/+bYhP/m2IT/5phEv+aYRL/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5ti
E/+aYRP/lFoS/5JYEf+SVxH/kVYR/5BVEf+PVBH/e0gO/wkFAf8AAAD9AAAAVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAA5QAAAP5sTA//sXwZ/rF8Gv+wexn+r3oZ/695Gf6tdxj/n2YU/p1k
FP+dZBP+nWQU/51kE/6dZBT/nWQT/p1kFP+dZBP+nGIT/4lNDv6BQwz/gUMM/odLDv+VWxH+nWQU/51kE/6dZBT/nWQT/p1kFP+dZBP+nWQU/51kE/6dZBT/nWQT/p1kFP+dZBP+nWQU/51kE/6dZBT/nWQT/p1k
FP+dZBP+nWQU/51kE/6dZBT/nWQT/pxjE/+UWRL+klgR/5FXEf6RVhH/kFUQ/o9UEf9VMgr+AAAA/wAAAOIAAAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhwAAAP8rHgb/sn4a/7J9Gv+xfBr/sXwa/7B7
Gv+vehn/o2sV/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/4xQEP+wimj/2si4/9nHt/+4lnj/iE4c/4lNDv+aYRP/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59m
FP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/m2IT/5NZEv+SWBH/klcR/5FWEf+QVRH/jlQQ/yITA/8AAAD/AAAAgwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFwAAAPICAQD/jGMU/rN/
G/+zfhr+sn0a/7F8Gf6xfBr/qXMX/qBoFf+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/59nFP65l3f//v7+/v/////+/v7+//////Ls5v6yjGv/hEgR/pBVEP+eZhT+oGgV/6BoFP6gaBX/oGgU/qBo
Ff+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/6BoFP6ZXxP/k1kR/pJYEf+RVxH+kVYR/5BVEP5wQg3/AQEA/gAAAPIAAAAWAAAAAAAAAAAAAAAAAAAAAAAA
AIEAAAD/NycI/7WAGv+0gBv/s38b/7N+Gv+yfRr/sHsZ/6NrFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+fZxf/8+3n/////////////////////////////////+XYzf+hc0f/h0oO/5dd
Ev+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/oWkU/5VbEv+TWRL/klgR/5JXEf+RVhH/kFUQ/ysZBf8AAAD/AAAAgAAA
AAAAAAAAAAAAAAAAAAkAAADlAAAA/otjFf+2gRv+tYEb/7SAGv6zfxv/s34a/qpzF/+jbBX+o2wW/6NsFf6jbBb/o2wV/qNsFv+jbBX+o2wW/6NsFf6jbBb/rXsx/v/////+/v7+//////7+/v7//////v7+/v//
///+/v7+/v39/9XArP6VYCz/jVEP/p1kFP+jbBX+o2wW/6NsFf6jbBb/o2wV/qNsFv+jbBX+o2wW/6NsFf6jbBb/o2wV/qNsFv+jbBX+o2wW/6NsFf6jbBb/o2wV/qNsFv+dZBT+lFkS/5NZEf6SWBH/kVcR/pFW
Ef9uQQz+AAAA/wAAAOQAAAAJAAAAAAAAAAAAAABPAAAA/yIYBf+1ghv/toIb/7aBG/+1gRv/tIAb/7J+Gv+mbxb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/7aJQf//////////////
////////////////////////////////////////+fbz/8SmiP+QVxz/k1kR/6JqFf+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pG0V/5Zc
Ev+UWRL/k1kS/5JYEf+SVxH/j1UQ/xsQA/8AAAD/AAAATwAAAAAAAAAAAAAApAAAAP9jRw/+uIQc/7eDG/62ghv/toEb/rWBG/+ueBn+p3AX/6dwFv6ncBf/p3AW/qdwF/+ncBb+p3AX/6dwFv6ncBf/p3AW/qdw
F/+3ikL+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+8Onh/7SNYv6PVBL/mmET/qZvFv+ncBb+p3AX/6dwFv6ncBf/p3AW/qdwF/+ncBb+p3AX/6dwFv6ncBf/p3AW/qdw
F/+ncBb+p3AX/6dwFv6dZBT/lFoS/pRZEv+TWRH+klgR/5FXEf5OLwn/AAAA/gAAAKYAAAAAAAAABwAAAOoAAAD/n3IY/7iFHP+4hBz/t4Mc/7aCG/+2gRv/q3QX/6hxF/+ocRf/qHEX/6hxF/+ocRf/qHEX/6hx
F/+ocRf/qHEX/6hxF/+ocRf/uIxC////////////////////////////////////////////////////////////////////////////5NbF/6Z4Q/+TWBH/oGgU/6hxF/+ocRf/qHEX/6hxF/+ocRf/qHEX/6hx
F/+ocRf/qHEX/6hxF/+ocRf/qHEX/6hxF/+ocRf/pG0W/5VbEv+UWhL/lFkS/5NZEv+SWBH/fksO/wAAAP8AAADrAAAABwAAADQAAAD+GhME/7mGHP65hhz/uIUc/riEHP+3gxv+tYAb/6pzGP6qcxj/qnMX/qpz
GP+qcxf+qnMY/6pzF/6qcxj/qnMX/qpzGP+qcxf+qnMY/7qNQv7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////9/Pv+1b6k/51pKv6YXxP/pW4W/qpz
F/+qcxf+qnMY/6pzF/6qcxj/qnMX/qpzGP+qcxf+qnMY/6pzF/6qcxj/qnMX/qpzF/+YXhP+lVsS/5RaEv6UWRL/k1kR/pJYEf8WDQL+AAAA/gAAADcAAABrAAAA/0czCv+6hx3/uoYd/7mGHP+4hRz/uIQc/7N+
Gv+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+7j0L/////////////////////////////////////////////////////////////////////////////////////////
///38+//xqZ//5lhGv+eZhT/qXMX/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/nmUU/5ZcE/+VWxL/lFoS/5RZEv+TWRL/OiMH/wAAAP8AAABvAAAAmwAAAP5sTxD/u4gc/rqH
Hf+6hhz+uYYc/7iFHP6yfRr/rXcY/q13Gf+tdxj+rXcZ/613GP6tdxn/rXcY/q13Gf+tdxj+rXcZ/613GP6tdxn/vJBD/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v//
///+/v7+//////7+/v7//////v7+/v/////v59v+t5Bc/5phFP6kbRb/rXcY/q13Gf+tdxj+rXcZ/613GP6tdxn/rXcY/q13Gf+tdxj+rXcZ/6NrFv6WXBP/llwS/pVbEv+UWhL+lFkS/1c0Cv4AAAD/AAAAnwAA
AMEAAAD/i2YV/7yJHf+7iB3/uocd/7qGHf+5hhz/sXwZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/72SQ///////////////////////////////////////////////
//////////////////////////////////////////////////////////////7+/v/j1L//rH07/6RsFf+ueRj/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ocRf/l10T/5ZcE/+WXBP/lVsS/5Va
Ev9wQw3/AAAA/wAAAMYAAADdAAAA/qJ2Gf+9ih3+vIkd/7uIHP66hx3/uoYc/rJ9Gv+wexn+sHsZ/7B7Gf6wexn/sHsZ/rB7Gf+wexn+sHsZ/7B7Gf6wexn/sHsZ/rB7Gf++k0T+//////7+/v7//////v7+/v//
///+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/vv49f+8lV3+rHYY/7B7Gf6wexn/sHsZ/rB7Gf+wexn+sHsZ/7B7Gf6wexn/q3UY/phe
E/+XXRL+llwT/5ZcEv6VWxL/gk8P/gEAAP8AAADjAAAA7wUEAP+ufxv/vYoe/72KHf+8iR3/u4gd/7qHHf+yfhr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/wJVE////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+vfy/7WFMv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8
Gv+xfBr/sXwa/655Gf+ZXxP/mF4T/5ddE/+WXBP/llwT/4pUEf8FAwD/AAAA9wAAAP4HBQH+soMc/76LHf69ih7/vYod/ryJHf+7iBz+tH8b/7N+Gv6zfhr/s34a/rN+Gv+zfhr+s34a/7N+Gv6zfhr/s34a/rN+
Gv+zfhr+s34a/8GWRf7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7RsHT/s34a/rN+
Gv+zfhr+s34a/7N+Gv6zfhr/s34a/rN+Gv+wexn+mWAT/5hfEv6YXhP/l10S/pZcE/+QWBL+CAUB/wAAAP4AAAD/CAYB/7WFHP+/jB7/vose/72KHv+9ih3/vIkd/7WBG/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SA
G/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SAG//Cl0X/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////zqto/7SAG/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SAG/+0gBv/sHsZ/5phE/+ZYBP/mV8T/5heE/+XXRP/kVkS/wkFAf8AAAD/AAAA/woHAf64hx3/v40e/r+MHv++ix3+vYoe/72KHf63gxv/toEb/raB
G/+2gRv+toEb/7aBG/62gRv/toEb/raBG/+2gRv+toEb/7aBG/62gRv/w5lF/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v//
///+/v7+//////7+/v7/////9O3e/rqIJv+2gRv+toEb/7aBG/62gRv/toEb/raBG/+2gRv+toEb/696Gf6aYRP/mmES/plgE/+YXxL+mF4T/5JaEv4IBQH/AAAA/gAAAP8IBQD/toUc/8COHv+/jR7/v4we/76L
Hv+9ih7/uYYc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/8WaRv//////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////8urY/8GUOv+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+tdxj/m2IT/5phE/+aYRP/mWAT/5lfE/+SWhH/CAUA/wAAAP8AAAD/CAQA/rKA
Gv/Bjx7+wI4e/7+NHv6/jB7/vosd/ruIHf+4hRz+uYUc/7iFHP65hRz/uIUc/rmFHP+4hRz+uYUc/7iFHP65hRz/uIUc/rmFHP/FnEb+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v//
///+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v769/D/2b6G/ruJJP+4hRz+uYUc/7iFHP65hRz/uIUc/rmFHP+4hRz+uYUc/7iFHP65hRz/qXMX/pxjE/+bYhP+mmET/5phEv6ZYBP/jlcR/gYD
AP8AAAD+AAAA/wcEAP+qdxf/wY8f/8GPH//Ajh7/v40e/7+MHv+9ih3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/x51G////////////////////////////////////
//////////////////////////////////////////////////////////////79/P/l0an/wZM0/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/6RsFf+dZBT/nGMT/5ti
E/+aYRP/mmET/4lTEP8FAgD/AAAA/gAAAPcEAgD+m2gT/8KQHv7Bjx//wY8e/sCOHv+/jR7+v4wd/7yIHf68iB3/vIgc/ryIHf+8iBz+vIgd/7yIHP68iB3/vIgc/ryIHf+8iBz+vIgd/8ieRv7//////v7+/v//
///+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+/////+7iyP7KoUv/vIgc/ryIHf+8iBz+vIgd/7yIHP68iB3/vIgc/ryIHf+8iBz+vIgd/7yIHP68iB3/vIgc/rmF
HP+fZxT+nWUU/51kE/6cYxP/m2IT/pphE/+CTQ7+AwEA/wAAAO8AAADjAAAA/4dVD//DkR//wpAf/8GPH//Bjx//wI4e/7+NHv++ix3/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72K
Hf/JoEf/////////////////////////////////////////////////////////////////////////////////9/Dj/9Szbf++iyD/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72K
Hf+9ih3/vYod/72KHf+xfBr/n2YU/55mFP+dZRT/nWQU/5xjE/+bYhP/dkIM/wAAAP8AAADcAAAAxQAAAP5qPQn/wY8f/sORH//CkB7+wY8f/8GPHv7Ajh7/v4we/r6LHv++ix3+vose/76LHf6+ix7/vosd/r6L
Hv++ix3+vose/76LHf6+ix7/yqFI/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////8+vX+38WP/8KSKv6+ix7/vosd/r6LHv++ix3+vose/76LHf6+ix7/vosd/r6L
Hv++ix3+vose/76LHf6+ix7/vosd/r6LHv++ix3+p3AW/59nFP6fZhT/nmYT/p1lFP+dZBP+nGMT/2AzCf4AAAD/AAAAwAAAAJ4AAAD/Ui0G/7N+Gf/Dkh//w5Ef/8KQH//Bjx//wY8f/8COHv+/jR7/v40e/8CN
Hv+/jR7/wI0e/7+NHv/AjR7/v40e/8CNHv+/jR7/wI0e/8uiSP////////////////////////////////////////////////////////////79/P/p2LP/x5s6/8CNHv+/jR7/wI0e/7+NHv/AjR7/v40e/8CN
Hv+/jR7/wI0e/7+NHv/AjR7/v40e/8CNHv+/jR7/wI0e/7+NHv/AjR7/uYUc/6FpFf+gaBX/oGcU/59mFP+eZhT/nWUU/5ZcEv9KJwb/AAAA/wAAAJsAAABwAAAA/jYeBP+gZxL+xJMg/8OSH/7DkR//wpAe/sGP
H//Bjx7+wI4f/8GPHv7Bjx//wY8e/sGPH//Bjx7+wY8f/8GPHv7Bjx//wY8e/sGPH//Mokb+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/vLo0P/Qqlb+wY8f/8GPHv7Bjx//wY8e/sGP
H//Bjx7+wY8f/8GPHv7Bjx//wY8e/sGPH//Bjx7+wY8f/8GPHv7Bjx//wY8e/sGPH//Bjx7+wY4e/6p0F/6hahX/oWkU/qBoFf+fZxT+n2YU/55mE/6NUg//MBkE/gAAAP8AAABrAAAANwAAAP4UCwH/jVEM/8GO
Hv/EkyD/w5If/8ORH//CkB//wY8f/8GPH//Bjx//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//x5oy//7+/v//////////////////////////////////////+fTp/9q8d//DkiP/wpAf/8KQ
H//CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/7iEG/+jaxb/omoV/6FqFf+haRX/oGgV/59nFP+dZRP/hEcM/xIJAf8AAAD+AAAANAAA
AAgAAADrAAAA/3hDCf6sdBb/xpMf/sSTIP/Dkh/+w5Ef/8KQHv7Bjx//wY8f/sORH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8SSIf7069b//v7+/v/////+/v7+//////7+/v79+/j/5M2Z/seZ
Lv/DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8GPHv6ocRb/pGwV/qNrFv+iahX+oWoV/6FpFP6gaBX/lFoR/m86
Cv8AAAD+AAAA6QAAAAYAAAAAAAAApwAAAP9MKgX/klYN/8SRHv/GlCD/xJMg/8OSH//DkR//wpAf/8GPH//CkB//xZMf/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/1bBb//z69v////////////79
/P/t3rv/zqRC/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTH/+vehn/pW4W/6RtFv+kbBb/o2sW/6Jq
Ff+hahX/oGgU/4hLDf9GJQb/AAAA/wAAAKMAAAAAAAAAAAAAAFEAAAD/Gg4B/opOCv+rdBb+x5Qg/8aTH/7EkyD/w5If/sORH//CkB7+wY8f/8ORH/7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seU
IP/QpUL+4MSE/+DDgv7SqUr/x5Uh/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP62gRv/pm8W/qZv
Fv+lbhb+pG0W/6RsFf6jaxb/omoV/pVbEP+CRQv+FwwC/wAAAP4AAABOAAAAAAAAAAAAAAAJAAAA5QAAAP9qPAj/kFQM/8GNHv/HlCD/xpQg/8STIP/Dkh//w5Ef/8KQH//Bjx//xJIf/8iWIP/IliH/yJYh/8iW
If/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iW
If+5hRz/qHEX/6dwF/+mbxf/pm8W/6VuFv+kbRb/pGwW/6BoFf+HSgv/ZDUJ/wAAAP8AAADkAAAACQAAAAAAAAAAAAAAAAAAAIEAAAD+KhcD/4xQCv6fZhH/x5Qg/seUIP/Gkx/+xJMg/8OSH/7DkR//wpAe/sGP
H//Fkx/+yZcg/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smX
If/JlyD+yZch/8mXIP65hRz/qXMX/qhyF/+ocRb+p3AX/6ZvFv6mbxb/pW4W/qRtFf+RVg7+hUcL/ycVA/4AAAD/AAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAXAAAA8gEBAP9tPgj/jlIK/7B5F//HlSD/x5Qg/8aU
IP/EkyD/w5If/8ORH//CkB//wY8f/8SSH//KmCH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZ
If/KmSH/ypkh/8qZIf/KmSH/ypkh/8mXIf+3ghv/q3QY/6pzGP+pcxf/qHIX/6hxF/+ncBf/pm8X/6ZvFv+bYRH/h0kL/2c4CP8BAAD/AAAA8wAAABYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIYAAAD/IRIC/o1R
Cv+SVgr+u4cb/8eVIP7HlCD/xpMf/sSTIP/Dkh/+w5Ef/8KQHv7Bjx//wpAf/smXIf/MmiH+zJoi/8yaIf7MmiL/zJoh/syaIv/MmiH+zJoi/8yaIf7MmiL/zJoh/syaIv/MmiH+zJoi/8yaIf7MmiL/zJoh/sya
Iv/MmiH+zJoi/8yaIf7MmiL/zJoh/syaIv/MmiH+zJoi/8WTIP6yfRr/rHYY/qt1GP+qdBf+qnMY/6lzF/6ochf/qHEW/qdwF/+haRT+ik4L/4ZJCv4fEQL/AAAA/gAAAIYAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAPAAAA5AAAAP9VMQb/kFMK/5ZbDP/AjB3/x5Ug/8eUIP/GlCD/xJMg/8OSH//DkR//wpAf/8GPH//Bjx//xZMg/8uaIf/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/zZwi/82c
Iv/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/ypgh/7uIHP+veRn/rXgZ/613Gf+sdhj/q3UY/6t0GP+qcxj/qXMX/6hyF/+lbRX/jlIM/4lMC/9SLQb/AAAA/wAAAOQAAAAPAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFcAAAD9CQUA/3tHCP6RVQr/ml8M/sKOHv/HlSD+x5Qg/8aTH/7EkyD/w5If/sORH//CkB7+wY8f/8GPHv7Bjh//xZQf/syaIf/OnSL+zp0i/86dIv7OnSL/zp0i/s6d
Iv/OnSL+zp0i/86dIv7OnSL/zp0i/s6dIv/OnSL+zp0i/86dIv7OnSL/zp0i/s6dIv/KmCD+vosd/7J9Gv6vehn/r3kZ/q55Gf+teBj+rXcZ/6x2GP6rdRj/qnQX/qpzGP+ncBb+klYN/4tOCv52QQn/CAQA/gAA
AP0AAABZAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAqAAAAP8fEgL/i1EJ/5NXCv+cYQ3/wY0d/8eVIP/HlCD/xpQg/8STIP/Dkh//w5Ef/8KQH//Bjx//wY8f/8COHv/AjR7/w5Ef/8iW
IP/NnCL/z54i/8+eI//PniP/z54j/8+eI//PniP/z54j/8+eI//PniP/z54j/8+eI//PniL/zJsi/8SSH/+6hx3/s34a/7F8Gv+xfBr/sHsa/696Gf+veRn/rnkZ/614Gf+tdxn/rHYY/6t1GP+ncRb/lFgM/41Q
C/+GSwr/HhAC/wAAAP8AAACqAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABEAAADbAAAA/jUfA/+SVgn+lVkK/5tgC/68iBv/x5Ug/seUIP/Gkx/+xJMg/8OSH/7DkR//wpAe/sGP
H//Bjx7+wI4e/7+NHv6/jB7/vosd/sCNHv/DkB/+xZMf/8eWIP7JlyD/ypgh/sqYIf/IlyD+xpQg/8ORH/6/jB7/uYYc/rWBG/+0gBr+s38b/7N+Gv6yfRr/sXwZ/rF8Gv+wexn+r3oZ/695Gf6ueRn/rXgY/q13
Gf+ncBX+lFkL/49TCv6MTwr/Mx0E/gAAAP8AAADcAAAAEgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALwAAAO8AAAD/QygE/5RYCf+XWwn/ml8J/7R+Fv/HlCD/x5Qg/8aU
IP/EkyD/w5If/8ORH//CkB//wY8f/8GPH//Ajh7/v40e/7+MHv++ix7/vYoe/72KHf+8iR3/u4gd/7qHHf+6hh3/uYYc/7iFHP+4hBz/t4Mc/7aCG/+2gRv/tYEb/7SAG/+zfxv/s34a/7J9Gv+xfBr/sXwa/7B7
Gv+vehn/r3kZ/654GP+kbBP/lVkK/5FVCv+PUgn/QiYE/wAAAP8AAADwAAAAMQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABJAAAA9wAAAP5FKQT/lloJ/pld
Cf+bYAn+qXAP/8KPHf7HlCD/xpMf/sSTIP/Dkh/+w5Ef/8KQHv7Bjx//wY8e/sCOHv+/jR7+v4we/76LHf69ih7/vYod/ryJHf+7iBz+uocd/7qGHP65hhz/uIUc/riEHP+3gxv+toIb/7aBG/61gRv/tIAa/rN/
G/+zfhr+sn0a/7F8Gf6xfBr/sHsZ/q13F/+gZg7+lloK/5RYCf6QVAn/RykE/gAAAP8AAAD3AAAATAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AFAAAAD3AAAA/zwkA/+VWwj/m2AJ/51iCf+hZwr/tX8U/8WSHv/GlCD/xJMg/8OSH//DkR//wpAf/8GPH//Bjx//wI4e/7+NHv+/jB7/vose/72KHv+9ih3/vIkd/7uIHf+6hx3/uoYd/7mGHP+4hRz/uIQc/7eD
HP+2ghv/toEb/7WBG/+0gBv/s38b/7N+Gv+yfRr/sXwZ/6hxEv+dYgr/mV0J/5ZaCv+RVgn/PCME/wAAAP8AAAD6AAAAWQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAATQAAAPMAAAD+JxgC/41XCP6dYgn/n2QI/qJnCP+obgr+uIIV/8SRHv7EkyD/w5If/sORH//CkB7+wY8f/8GPHv7Ajh7/v40e/r+MHv++ix3+vYoe/72KHf68iR3/u4gc/rqH
Hf+6hhz+uYYc/7iFHP64hBz/t4Mb/raCG/+2gRv+tYEb/7SAGv6zfhr/rHUT/qNpC/+eYwn+m2AJ/5ldCf6JUgn/KBgC/gAAAP8AAAD0AAAATwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5AAAA4wAAAP8QCgD/cEYG/59kCP+iZwj/pGoI/6ZsCP+rcgn/toAR/8CMGv/Dkh//w5Ef/8KQH//Bjx//wY8f/8COHv+/jR7/v4we/76L
Hv+9ih7/vYod/7yJHf+7iB3/uocd/7qGHf+5hhz/uIUc/7iEHP+3gxz/toIb/7R/F/+vdxD/qG4J/6RpCP+hZgj/nmMJ/5tgCP9vQwb/EQoB/wAAAP8AAADlAAAAPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABsAAAC8AAAA/gEAAP86JAP+kl0H/6RpB/6mbAj/qW8H/qxyB/+udAf+s3sK/7qEEf6/ixf/wY8c/sGP
Hv/Bjx7+wI4e/7+NHv6/jB7/vosd/r2KHv+9ih3+vIkd/7uIHP66hx3/uoYc/rmFGv+3ghX+tX4Q/7F4Cv6scwf/qm8H/qZsCP+kaQf+oWYI/5FaCP47JQP/AQAA/gAAAP8AAADAAAAAHAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwAAAHUAAADzAAAA/woGAP9TNQP/m2QH/6lvCP+scgf/rnQH/7F3
B/+0egf/tn0G/7mABv+8hAj/v4kM/8GLD//CjRH/w44S/8SPE//EjxP/w44S/8GMEf/Aig7/vocL/7yECf+5gAb/tn0G/7N5B/+vdgf/rHMH/6pvB/+nbAj/mWIH/1Q1BP8KBgD/AAAA/wAAAPQAAAB5AAAABAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAALAAAAD9AAAA/gwI
AP9OMwP+k2EG/650Bv6xdwf/s3oG/rZ9Bv+5gAX+vIMG/76GBf7CigX/xY0E/siQBP/LkwP+zZUE/8qSBP7HjwX/w4sE/r+HBf+8hAX+uYAG/7Z9Bv6zeQf/r3YG/qxyBv+TYQb+TzMD/w0IAP4AAAD/AAAA/QAA
ALMAAAAnAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAQAAAAL0AAAD+AAAA/wIBAP8qHAH/ZkQD/5xqBv+2fQb/uYAG/7yDBv+/hgX/wooF/8WNBf/JkAT/y5ME/82VBP/KkgT/x48F/8OLBf+/hwX/vIQG/7mABv+2fQb/nGoG/2dFBP8sHQH/AwEA/wAA
AP8AAAD+AAAAwAAAAEMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANAAAAKAAAADwAAAA/gAAAP8AAAD+GBEA/0IuAv5oSAP/h18D/qFzBP+ygAT+vYcE/8SOA/7GkAP/v4oD/rWBBP+idAT+iGAD/2hJA/5ELwL/GREA/gAA
AP8AAAD+AAAA/wAAAPEAAACjAAAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEQAAAFoAAACqAAAA6QAAAP0AAAD/AAAA/wAAAP8AAAD/BAMA/wkGAP8MCAD/DAgA/wkGAP8EAwD/AAAA/wAA
AP8AAAD/AAAA/wAAAP4AAADrAAAArAAAAF4AAAASAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYAAAAwAAAAZQAAAJMAAAC4AAAA2AAAAO0AAAD9AAAA/gAA
AP8AAAD9AAAA7QAAANgAAAC6AAAAlQAAAGcAAAAzAAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA///+AAB///////AAAA//////gAAAA/////8AAAAA/////AAAAAA////wAAAAAA///+AAAAAAB///wAAAAAAD//+AAAAAAAH//wAAAAAAAP/+AAAAAAAAf/wAAAAAAAA/+AAAAAAAAB/wAAAAAAAAD/AA
AAAAAAAP4AAAAAAAAAfgAAAAAAAAB8AAAAAAAAADwAAAAAAAAAOAAAAAAAAAAYAAAAAAAAABgAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAGAAAAAAAAAAYAA
AAAAAAABwAAAAAAAAAPAAAAAAAAAA+AAAAAAAAAH4AAAAAAAAAfwAAAAAAAAD/AAAAAAAAAP+AAAAAAAAB/8AAAAAAAAP/4AAAAAAAB//wAAAAAAAP//gAAAAAAB///AAAAAAAP//+AAAAAAB///8AAAAAAP///8
AAAAAD////8AAAAA/////8AAAAP/////8AAAD//////+AAB///8=
"@
#endregion ******** $Play64Icon ********
$PILLargeImageList.Images.Add("Play64Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Play64Icon))))

#region ******** $Pause64Icon ********
$Pause64Icon = @"
AAABAAEAQEAAAAEAIAAoQgAAFgAAACgAAABAAAAAgAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAA2AAAAagAAAJYAAAC8AAAA2QAAAO4AAAD+AAAA/wAAAP8AAAD9AAAA7QAAANgAAAC5AAAAkwAAAGUAAAAwAAAABgAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVAAAAYgAAALEAAADtAAAA/gAAAP4AAAD/AAAA/gAAAP8DAgD+BwQA/wkFAf4JBQH/BgQA/gMCAP8AAAD+AAAA/wAAAP4AAAD/AAAA/QAA
AOkAAACoAAAAWgAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAADwAAACoAAAA9AAAAP8AAAD/AQAA/xYOAv85JAf/VzcK/25FDf+BUQ//i1cR/5BaEf+TWxL/kloS/45XEf+IUxH/fEwP/2lA
DP9SMQn/NB8G/xQMAv8AAAD/AAAA/wAAAP8AAADvAAAAnwAAADMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARwAAAMQAAAD/AAAA/gIBAP8pGgX+XjwM/4pZEf6dZRT/nWQT/pxjE/+bYhP+mmET/5phEv6ZYBP/mF8S/phe
E/+XXRL+llwT/5ZcEv6VWxL/lFoS/pRZEv+TWRH+f0wP/1QyCv4kFQT/AgEA/gAAAP8AAAD+AAAAuwAAAD4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArAAAAuAAAAP0AAAD/DQgB/00yCv+MWxL/oGgU/59nFP+fZhT/nmYU/51lFP+dZBT/nGMT/5ti
E/+aYRP/mmET/5lgE/+ZXxP/mF4T/5ddE/+WXBP/llwT/5VbEv+UWhL/lFkS/5NZEv+SWBH/kVYR/31KDv9CJwf/CgYB/wAAAP8AAAD8AAAArQAAACMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAB9AAAA9QAAAP4LBwH/VjkL/pllFP+iahX+oWoV/6FpFP6gaBX/n2cU/p9m
FP+eZhP+nWUU/51kE/6cYxP/m2IT/pphE/+aYRL+mWAT/5hfEv6YXhP/l10S/pZcE/+WXBL+lVsS/5RaEv6UWRL/k1kR/pJYEf+RVxH+kVYR/4ZPD/5JKwj/CAUB/gAAAP8AAADxAAAAcAAAAAMAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB4AAADCAAAA/wEBAP8/Kgj/l2QU/6RtFv+kbBb/o2sW/6Jq
Ff+hahX/oWkV/6BoFf+gZxT/n2YU/55mFP+dZBP/m2IT/5lgE/+YXhL/llwS/5ZcEv+WXBL/ll0S/5ddEv+XXRL/llwT/5ZcE/+VWxL/lFoS/5RZEv+TWRL/klgR/5JXEf+RVhH/kFUR/4JMD/81Hgb/AQAA/wAA
AP8AAAC5AAAAGQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD0AAADmAAAA/xMNAv55URD/pm8W/qZv
Fv+lbhb+pG0W/6RsFf6jaxb/omoV/qFqFf+gaBT+nWQU/5heEv6TWRH/kFUR/o9UEf+PVBD+j1QQ/49UEP6PVBD/j1QQ/o9UEP+PVBH+kFUR/5FWEf6TWRH/lFoS/pVbEv+UWhL+lFkS/5NZEf6SWBH/kVcR/pFW
Ef+QVRD+j1QQ/2U7C/4PCQH/AAAA/gAAAOEAAAA2AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFAAAAD0AAAA/y4f
Bv+aaBX/qHEX/6dwF/+mbxf/pm8W/6VuFv+kbRb/pGwW/6FpFf+aYRP/k1kR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+SVxH/k1kR/5Ra
Ev+UWRL/k1kS/5JYEf+SVxH/kVYR/5BVEf+PVBH/gUwP/yUVBP8AAAD/AAAA8QAAAEkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AFkAAAD6AAAA/0cwCf6mcRf/qXMX/qhyF/+ocRb+p3AX/6ZvFv6mbxb/pGwV/pthE/+UWRL+k1gS/5NYEf6TWBL/k1gR/pNYEv+TWBH+k1gS/5NYEf6TWBL/k1gR/pNYEv+TWBH+k1gS/5NYEf6TWBL/k1gR/pNY
Ev+TWBH+k1gS/5NYEf6TWBL/lFkR/pRZEv+TWRH+klgR/5FXEf6RVhH/kFUQ/o9UEf+LUhD+OSEG/wAAAP4AAAD3AAAAUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAE0AAAD3AAAA/1U6DP+qdBf/q3QY/6pzGP+pcxf/qHIX/6hxF/+ncBb/n2cU/5ZcEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5Ra
Ev+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5NZEv+SWBH/klcR/5FWEf+QVRH/j1QR/45TEP9DJwf/AAAA/wAAAPYAAABFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADIAAADwAAAA/1I4C/6sdhj/rHYY/qt1GP+qdBf+qnMY/6lzF/6ncBb/m2IU/pZcE/+WXBL+llwT/5ZcEv6WXBP/llwS/pZcE/+WXBL+llwT/5ZcEv6WXBP/llwS/pZc
E/+WXBL+llwT/5ZcEv6WXBP/llwS/pZcE/+WXBL+llwT/5ZcEv6WXBP/llwS/pZcE/+WXBL+llwT/5ZcEv6VWxL/k1kS/pJYEf+RVxH+kVYR/5BVEP6PVBH/jlMQ/kEmB/8AAAD+AAAA7gAAACwAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIAAADdAAAA/0EtCf+tdxj/rXgZ/613Gf+sdhj/q3UY/6p0GP+ncBf/mmET/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5he
E/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5deEv+TWRL/klgR/5JXEf+RVhH/kFUR/49UEf+NUxD/NB4G/wAA
AP8AAADZAAAADwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAACrAAAA/ycbBf6pdhj/r3kZ/q55Gf+teBj+rXcZ/6x2GP6ncRb/m2IT/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppg
E/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mWAS/5RaEv6SWBH/kVcR/pFW
Ef+QVRD+j1QR/4lQEP4fEgP/AAAA/gAAAKUAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABaAAAA/QsIAf+Zaxb/sHsa/696Gf+veRn/rnkZ/614Gf+pcxf/nGMT/5tiE/+bYhP/m2IT/5ti
E/+bYhP/m2IT/5tiE/+bYhP/m2IT/5phE/+ZYBL/mmES/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+aYRL/mWAS/5phE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5ti
E/+aYRP/lFoS/5JYEf+SVxH/kVYR/5BVEf+PVBH/e0gO/wkFAf8AAAD9AAAAVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAA5QAAAP5sTA//sXwZ/rF8Gv+wexn+r3oZ/695Gf6tdxj/n2YU/p1k
FP+dZBP+nWQU/51kE/6dZBT/nWQT/p1kFP+dZBP+nGMT/41RD/6CRQz/gEIM/oFDDP+ISw7+mF4S/51kE/6dZBT/nWQT/p1kFP+dZBP+nWQU/5heEv6ISw7/gUMM/oBCDP+CRQz+jVEP/5xjE/6dZBT/nWQT/p1k
FP+dZBP+nWQU/51kE/6dZBT/nWQT/pxjE/+UWRL+klgR/5FXEf6RVhH/kFUQ/o9UEf9VMgr+AAAA/wAAAOIAAAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhwAAAP8rHgb/sn4a/7J9Gv+xfBr/sXwa/7B7
Gv+vehn/o2sV/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/nmYU/4lNEP+rgl//28m6/+rg1//k1sv/wKGG/4lQG/+aYRP/n2YU/59mFP+fZhT/n2YU/5phE/+JUBv/wKGG/+TWy//q4Nf/28m6/6uC
X/+JTRD/nmYU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/m2IT/5NZEv+SWBH/klcR/5FWEf+QVRH/jlQQ/yITA/8AAAD/AAAAgwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFwAAAPICAQD/jGMU/rN/
G/+zfhr+sn0a/7F8Gf6xfBr/qXMX/qBoFf+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/5ddEv7IrZX//v7+/v/////+/v7+//////7+/v7p39b/lV4h/qBoFf+gaBT+oGgV/6BoFP6VXiH/6d/W/v//
///+/v7+//////7+/v7+/v7/yK2V/pddEv+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/6BoFP6ZXxP/k1kR/pJYEf+RVxH+kVYR/5BVEP5wQg3/AQEA/gAAAPIAAAAWAAAAAAAAAAAAAAAAAAAAAAAA
AIEAAAD/NycI/7WAGv+0gBv/s38b/7N+Gv+yfRr/sHsZ/6NrFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+sf0b//v79/////////////////////////////////8uxlP+iahX/omoV/6Jq
Ff+iahX/y7GU//////////////////////////////////7+/f+tf0b/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/oWkU/5VbEv+TWRL/klgR/5JXEf+RVhH/kFUQ/ysZBf8AAAD/AAAAgAAA
AAAAAAAAAAAAAAAAAAkAAADlAAAA/otjFf+2gRv+tYEb/7SAGv6zfxv/s34a/qpzF/+jbBX+o2wW/6NsFf6jbBb/o2wV/qNsFv+jbBX+o2wW/6NsFf6jbBb/zrKI/v/////+/v7+//////7+/v7//////v7+/v//
///v5dj+o2wW/6NsFf6jbBb/o2wV/u/l2P/+/v7+//////7+/v7//////v7+/v/////+/v7+zrKI/6NsFf6jbBb/o2wV/qNsFv+jbBX+o2wW/6NsFf6jbBb/o2wV/qNsFv+dZBT+lFkS/5NZEf6SWBH/kVcR/pFW
Ef9uQQz+AAAA/wAAAOQAAAAJAAAAAAAAAAAAAABPAAAA/yIYBf+1ghv/toIb/7aBG/+1gRv/tIAb/7J+Gv+mbxb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/9e/mP//////////////
////////////////////////8+vf/6VuFv+lbhb/pW4W/6VuFv/z6+D//////////////////////////////////////9e/mP+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pG0V/5Zc
Ev+UWRL/k1kS/5JYEf+SVxH/j1UQ/xsQA/8AAAD/AAAATwAAAAAAAAAAAAAApAAAAP9jRw/+uIQc/7eDG/62ghv/toEb/rWBG/+ueBn+p3AX/6dwFv6ncBf/p3AW/qdwF/+ncBb+p3AX/6dwFv6ncBf/p3AW/qdw
F//YwJj+//////7+/v7//////v7+/v/////+/v7+//////Pr4P6ncBf/p3AW/qdwF/+ncBb+8+vg//7+/v7//////v7+/v/////+/v7+//////7+/v7YwJj/p3AW/qdwF/+ncBb+p3AX/6dwFv6ncBf/p3AW/qdw
F/+ncBb+p3AX/6dwFv6dZBT/lFoS/pRZEv+TWRH+klgR/5FXEf5OLwn/AAAA/gAAAKYAAAAAAAAABwAAAOoAAAD/n3IY/7iFHP+4hBz/t4Mc/7aCG/+2gRv/q3QX/6hxF/+ocRf/qHEX/6hxF/+ocRf/qHEX/6hx
F/+ocRf/qHEX/6hxF/+ocRf/2cCY///////////////////////////////////////z7OD/qHEX/6hxF/+ocRf/qHEX//Ps4P//////////////////////////////////////2cCY/6hxF/+ocRf/qHEX/6hx
F/+ocRf/qHEX/6hxF/+ocRf/qHEX/6hxF/+ocRf/pG0W/5VbEv+UWhL/lFkS/5NZEv+SWBH/fksO/wAAAP8AAADrAAAABwAAADQAAAD+GhME/7mGHP65hhz/uIUc/riEHP+3gxv+tYAb/6pzGP6qcxj/qnMX/qpz
GP+qcxf+qnMY/6pzF/6qcxj/qnMX/qpzGP+qcxf+qnMY/9nBmf7//////v7+/v/////+/v7+//////7+/v7/////8+zg/qpzGP+qcxf+qnMY/6pzF/7z7OD//v7+/v/////+/v7+//////7+/v7//////v7+/tnB
mf+qcxf+qnMY/6pzF/6qcxj/qnMX/qpzGP+qcxf+qnMY/6pzF/6qcxj/qnMX/qpzF/+YXhP+lVsS/5RaEv6UWRL/k1kR/pJYEf8WDQL+AAAA/gAAADcAAABrAAAA/0czCv+6hx3/uoYd/7mGHP+4hRz/uIQc/7N+
Gv+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP/awpn///////////////////////////////////////Ps4P+rdRj/q3UY/6t1GP+rdRj/8+zg////////////////////
///////////////////awpn/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/nmUU/5ZcE/+VWxL/lFoS/5RZEv+TWRL/OiMH/wAAAP8AAABvAAAAmwAAAP5sTxD/u4gc/rqH
Hf+6hhz+uYYc/7iFHP6yfRr/rXcY/q13Gf+tdxj+rXcZ/613GP6tdxn/rXcY/q13Gf+tdxj+rXcZ/613GP6tdxn/28OZ/v/////+/v7+//////7+/v7//////v7+/v/////07OD+rXcZ/613GP6tdxn/rXcY/vTs
4P/+/v7+//////7+/v7//////v7+/v/////+/v7+28OZ/613GP6tdxn/rXcY/q13Gf+tdxj+rXcZ/613GP6tdxn/rXcY/q13Gf+tdxj+rXcZ/6NrFv6WXBP/llwS/pVbEv+UWhL+lFkS/1c0Cv4AAAD/AAAAnwAA
AMEAAAD/i2YV/7yJHf+7iB3/uocd/7qGHf+5hhz/sXwZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/9vDmf//////////////////////////////////////9O3g/655
Gf+ueRn/rnkZ/655Gf/07eD//////////////////////////////////////9vDmf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ocRf/l10T/5ZcE/+WXBP/lVsS/5Va
Ev9wQw3/AAAA/wAAAMYAAADdAAAA/qJ2Gf+9ih3+vIkd/7uIHP66hx3/uoYc/rJ9Gv+wexn+sHsZ/7B7Gf6wexn/sHsZ/rB7Gf+wexn+sHsZ/7B7Gf6wexn/sHsZ/rB7Gf/cxJr+//////7+/v7//////v7+/v//
///+/v7+//////Tt4P6wexn/sHsZ/rB7Gf+wexn+9O3g//7+/v7//////v7+/v/////+/v7+//////7+/v7cxJr/sHsZ/rB7Gf+wexn+sHsZ/7B7Gf6wexn/sHsZ/rB7Gf+wexn+sHsZ/7B7Gf6wexn/q3UY/phe
E/+XXRL+llwT/5ZcEv6VWxL/gk8P/gEAAP8AAADjAAAA7wUEAP+ufxv/vYoe/72KHf+8iR3/u4gd/7qHHf+yfhr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/3cWa////
///////////////////////////////////07eD/sXwa/7F8Gv+xfBr/sXwa//Tt4P//////////////////////////////////////3cWa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8
Gv+xfBr/sXwa/655Gf+ZXxP/mF4T/5ddE/+WXBP/llwT/4pUEf8FAwD/AAAA9wAAAP4HBQH+soMc/76LHf69ih7/vYod/ryJHf+7iBz+tH8b/7N+Gv6zfhr/s34a/rN+Gv+zfhr+s34a/7N+Gv6zfhr/s34a/rN+
Gv+zfhr+s34a/93Gmv7//////v7+/v/////+/v7+//////7+/v7/////9O3g/rN+Gv+zfhr+s34a/7N+Gv707eD//v7+/v/////+/v7+//////7+/v7//////v7+/t3Gmv+zfhr+s34a/7N+Gv6zfhr/s34a/rN+
Gv+zfhr+s34a/7N+Gv6zfhr/s34a/rN+Gv+wexn+mWAT/5hfEv6YXhP/l10S/pZcE/+QWBL+CAUB/wAAAP4AAAD/CAYB/7WFHP+/jB7/vose/72KHv+9ih3/vIkd/7WBG/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SA
G/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SAG//ex5r///////////////////////////////////////Xu4P+0gBv/tIAb/7SAG/+0gBv/9e7g///////////////////////////////////////ex5r/tIAb/7SA
G/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SAG/+0gBv/sHsZ/5phE/+ZYBP/mV8T/5heE/+XXRP/kVkS/wkFAf8AAAD/AAAA/woHAf64hx3/v40e/r+MHv++ix3+vYoe/72KHf63gxv/toEb/raB
G/+2gRv+toEb/7aBG/62gRv/toEb/raBG/+2gRv+toEb/7aBG/62gRv/38ea/v/////+/v7+//////7+/v7//////v7+/v/////17uD+toEb/7aBG/62gRv/toEb/vXu4P/+/v7+//////7+/v7//////v7+/v//
///+/v7+38ea/7aBG/62gRv/toEb/raBG/+2gRv+toEb/7aBG/62gRv/toEb/raBG/+2gRv+toEb/696Gf6aYRP/mmES/plgE/+YXxL+mF4T/5JaEv4IBQH/AAAA/gAAAP8IBQD/toUc/8COHv+/jR7/v4we/76L
Hv+9ih7/uYYc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/9/Im///////////////////////////////////////9e7g/7eDHP+3gxz/t4Mc/7eDHP/17uD/////////
/////////////////////////////9/Im/+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+tdxj/m2IT/5phE/+aYRP/mWAT/5lfE/+SWhH/CAUA/wAAAP8AAAD/CAQA/rKA
Gv/Bjx7+wI4e/7+NHv6/jB7/vosd/ruIHf+4hRz+uYUc/7iFHP65hRz/uIUc/rmFHP+4hRz+uYUc/7iFHP65hRz/uIUc/rmFHP/gyZv+//////7+/v7//////v7+/v/////+/v7+//////Xu4P65hRz/uIUc/rmF
HP+4hRz+9e7g//7+/v7//////v7+/v/////+/v7+//////7+/v7gyZv/uIUc/rmFHP+4hRz+uYUc/7iFHP65hRz/uIUc/rmFHP+4hRz+uYUc/7iFHP65hRz/qXMX/pxjE/+bYhP+mmET/5phEv6ZYBP/jlcR/gYD
AP8AAAD+AAAA/wcEAP+qdxf/wY8f/8GPH//Ajh7/v40e/7+MHv+9ih3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/4Mqb////////////////////////////////////
///17+D/uocd/7qHHf+6hx3/uocd//Xv4P//////////////////////////////////////4Mqb/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/6RsFf+dZBT/nGMT/5ti
E/+aYRP/mmET/4lTEP8FAgD/AAAA/gAAAPcEAgD+m2gT/8KQHv7Bjx//wY8e/sCOHv+/jR7+v4wd/7yIHf68iB3/vIgc/ryIHf+8iBz+vIgd/7yIHP68iB3/vIgc/ryIHf+8iBz+vIgd/+HKm/7//////v7+/v//
///+/v7+//////7+/v7/////9u/g/ryIHf+8iBz+vIgd/7yIHP727+D//v7+/v/////+/v7+//////7+/v7//////v7+/uHKm/+8iBz+vIgd/7yIHP68iB3/vIgc/ryIHf+8iBz+vIgd/7yIHP68iB3/vIgc/rmF
HP+fZxT+nWUU/51kE/6cYxP/m2IT/pphE/+CTQ7+AwEA/wAAAO8AAADjAAAA/4dVD//DkR//wpAf/8GPH//Bjx//wI4e/7+NHv++ix3/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72K
Hf/iy5v///////////////////////////////////////bv4P+9ih3/vYod/72KHf+9ih3/9u/h///////////////////////////////////////iy5v/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72K
Hf+9ih3/vYod/72KHf+xfBr/n2YU/55mFP+dZRT/nWQU/5xjE/+bYhP/dkIM/wAAAP8AAADcAAAAxQAAAP5qPQn/wY8f/sORH//CkB7+wY8f/8GPHv7Ajh7/v4we/r6LHv++ix3+vose/76LHf6+ix7/vosd/r6L
Hv++ix3+vose/76LHf6+ix7/4syb/v/////+/v7+//////7+/v7//////v7+/v/////27+D+vose/76LHf6+ix7/vosd/vbv4f/+/v7+//////7+/v7//////v7+/v/////+/v7+4syb/76LHf6+ix7/vosd/r6L
Hv++ix3+vose/76LHf6+ix7/vosd/r6LHv++ix3+p3AW/59nFP6fZhT/nmYT/p1lFP+dZBP+nGMT/2AzCf4AAAD/AAAAwAAAAJ4AAAD/Ui0G/7N+Gf/Dkh//w5Ef/8KQH//Bjx//wY8f/8COHv+/jR7/v40e/8CN
Hv+/jR7/wI0e/7+NHv/AjR7/v40e/8CNHv+/jR7/wI0e/+LKl///////////////////////////////////////9u/g/8CNHv+/jR7/wI0e/7+NHv/27+D//////////////////////////////////////+HK
lv+/jR7/wI0e/7+NHv/AjR7/v40e/8CNHv+/jR7/wI0e/7+NHv/AjR7/uYUc/6FpFf+gaBX/oGcU/59mFP+eZhT/nWUU/5ZcEv9KJwb/AAAA/wAAAJsAAABwAAAA/jYeBP+gZxL+xJMg/8OSH/7DkR//wpAe/sGP
H//Bjx7+wI4f/8GPHv7Bjx//wY8e/sGPH//Bjx7+wY8f/8GPHv7Bjx//wY8e/sGPH//Xtm3+//////7+/v7//////v7+/v/////+/v7+/////+zcuv7Bjx//wY8e/sGPH//Bjx7+7Ny6//7+/v7//////v7+/v//
///+/v7+//////7+/v7Xtm3/wY8e/sGPH//Bjx7+wY8f/8GPHv7Bjx//wY8e/sGPH//Bjx7+wY4e/6p0F/6hahX/oWkU/qBoFf+fZxT+n2YU/55mE/6NUg//MBkE/gAAAP8AAABrAAAANwAAAP4UCwH/jVEM/8GO
Hv/EkyD/w5If/8ORH//CkB//wY8f/8GPH//Bjx//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//xJQm//Ts2f////////////////////////////38+f/QqVH/wpAf/8KQH//CkB//wpAf/9Cp
Uf/9/Pn////////////////////////////07Nj/xJQm/8KQH//CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/7iEG/+jaxb/omoV/6FqFf+haRX/oGgV/59nFP+dZRP/hEcM/xIJAf8AAAD+AAAANAAA
AAgAAADrAAAA/3hDCf6sdBb/xpMf/sSTIP/Dkh/+w5Ef/8KQHv7Bjx//wY8f/sORH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8ORH/7KnTj/7d+9/v79/P/+/v7+//////bv4P7UsWD/w5If/sOS
H//DkR/+w5If/8ORH/7Dkh//1LFg/vbw4P/+/v7+//////79/P7t373/yp04/sOSH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8GPHv6ocRb/pGwV/qNrFv+iahX+oWoV/6FpFP6gaBX/lFoR/m86
Cv8AAAD+AAAA6QAAAAYAAAAAAAAApwAAAP9MKgX/klYN/8SRHv/GlCD/xJMg/8OSH//DkR//wpAf/8GPH//CkB//xZMf/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/LnjX/0qxS/86k
Qv/GlCL/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/GlSL/zqRC/9KsUv/LnTX/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTH/+vehn/pW4W/6RtFv+kbBb/o2sW/6Jq
Ff+hahX/oGgU/4hLDf9GJQb/AAAA/wAAAKMAAAAAAAAAAAAAAFEAAAD/Gg4B/opOCv+rdBb+x5Qg/8aTH/7EkyD/w5If/sORH//CkB7+wY8f/8ORH/7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seU
IP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP62gRv/pm8W/qZv
Fv+lbhb+pG0W/6RsFf6jaxb/omoV/pVbEP+CRQv+FwwC/wAAAP4AAABOAAAAAAAAAAAAAAAJAAAA5QAAAP9qPAj/kFQM/8GNHv/HlCD/xpQg/8STIP/Dkh//w5Ef/8KQH//Bjx//xJIf/8iWIP/IliH/yJYh/8iW
If/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iW
If+5hRz/qHEX/6dwF/+mbxf/pm8W/6VuFv+kbRb/pGwW/6BoFf+HSgv/ZDUJ/wAAAP8AAADkAAAACQAAAAAAAAAAAAAAAAAAAIEAAAD+KhcD/4xQCv6fZhH/x5Qg/seUIP/Gkx/+xJMg/8OSH/7DkR//wpAe/sGP
H//Fkx/+yZcg/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smX
If/JlyD+yZch/8mXIP65hRz/qXMX/qhyF/+ocRb+p3AX/6ZvFv6mbxb/pW4W/qRtFf+RVg7+hUcL/ycVA/4AAAD/AAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAXAAAA8gEBAP9tPgj/jlIK/7B5F//HlSD/x5Qg/8aU
IP/EkyD/w5If/8ORH//CkB//wY8f/8SSH//KmCH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZ
If/KmSH/ypkh/8qZIf/KmSH/ypkh/8mXIf+3ghv/q3QY/6pzGP+pcxf/qHIX/6hxF/+ncBf/pm8X/6ZvFv+bYRH/h0kL/2c4CP8BAAD/AAAA8wAAABYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIYAAAD/IRIC/o1R
Cv+SVgr+u4cb/8eVIP7HlCD/xpMf/sSTIP/Dkh/+w5Ef/8KQHv7Bjx//wpAf/smXIf/MmiH+zJoi/8yaIf7MmiL/zJoh/syaIv/MmiH+zJoi/8yaIf7MmiL/zJoh/syaIv/MmiH+zJoi/8yaIf7MmiL/zJoh/sya
Iv/MmiH+zJoi/8yaIf7MmiL/zJoh/syaIv/MmiH+zJoi/8WTIP6yfRr/rHYY/qt1GP+qdBf+qnMY/6lzF/6ochf/qHEW/qdwF/+haRT+ik4L/4ZJCv4fEQL/AAAA/gAAAIYAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAPAAAA5AAAAP9VMQb/kFMK/5ZbDP/AjB3/x5Ug/8eUIP/GlCD/xJMg/8OSH//DkR//wpAf/8GPH//Bjx//xZMg/8uaIf/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/zZwi/82c
Iv/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/ypgh/7uIHP+veRn/rXgZ/613Gf+sdhj/q3UY/6t0GP+qcxj/qXMX/6hyF/+lbRX/jlIM/4lMC/9SLQb/AAAA/wAAAOQAAAAPAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFcAAAD9CQUA/3tHCP6RVQr/ml8M/sKOHv/HlSD+x5Qg/8aTH/7EkyD/w5If/sORH//CkB7+wY8f/8GPHv7Bjh//xZQf/syaIf/OnSL+zp0i/86dIv7OnSL/zp0i/s6d
Iv/OnSL+zp0i/86dIv7OnSL/zp0i/s6dIv/OnSL+zp0i/86dIv7OnSL/zp0i/s6dIv/KmCD+vosd/7J9Gv6vehn/r3kZ/q55Gf+teBj+rXcZ/6x2GP6rdRj/qnQX/qpzGP+ncBb+klYN/4tOCv52QQn/CAQA/gAA
AP0AAABZAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAqAAAAP8fEgL/i1EJ/5NXCv+cYQ3/wY0d/8eVIP/HlCD/xpQg/8STIP/Dkh//w5Ef/8KQH//Bjx//wY8f/8COHv/AjR7/w5Ef/8iW
IP/NnCL/z54i/8+eI//PniP/z54j/8+eI//PniP/z54j/8+eI//PniP/z54j/8+eI//PniL/zJsi/8SSH/+6hx3/s34a/7F8Gv+xfBr/sHsa/696Gf+veRn/rnkZ/614Gf+tdxn/rHYY/6t1GP+ncRb/lFgM/41Q
C/+GSwr/HhAC/wAAAP8AAACqAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABEAAADbAAAA/jUfA/+SVgn+lVkK/5tgC/68iBv/x5Ug/seUIP/Gkx/+xJMg/8OSH/7DkR//wpAe/sGP
H//Bjx7+wI4e/7+NHv6/jB7/vosd/sCNHv/DkB/+xZMf/8eWIP7JlyD/ypgh/sqYIf/IlyD+xpQg/8ORH/6/jB7/uYYc/rWBG/+0gBr+s38b/7N+Gv6yfRr/sXwZ/rF8Gv+wexn+r3oZ/695Gf6ueRn/rXgY/q13
Gf+ncBX+lFkL/49TCv6MTwr/Mx0E/gAAAP8AAADcAAAAEgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALwAAAO8AAAD/QygE/5RYCf+XWwn/ml8J/7R+Fv/HlCD/x5Qg/8aU
IP/EkyD/w5If/8ORH//CkB//wY8f/8GPH//Ajh7/v40e/7+MHv++ix7/vYoe/72KHf+8iR3/u4gd/7qHHf+6hh3/uYYc/7iFHP+4hBz/t4Mc/7aCG/+2gRv/tYEb/7SAG/+zfxv/s34a/7J9Gv+xfBr/sXwa/7B7
Gv+vehn/r3kZ/654GP+kbBP/lVkK/5FVCv+PUgn/QiYE/wAAAP8AAADwAAAAMQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABJAAAA9wAAAP5FKQT/lloJ/pld
Cf+bYAn+qXAP/8KPHf7HlCD/xpMf/sSTIP/Dkh/+w5Ef/8KQHv7Bjx//wY8e/sCOHv+/jR7+v4we/76LHf69ih7/vYod/ryJHf+7iBz+uocd/7qGHP65hhz/uIUc/riEHP+3gxv+toIb/7aBG/61gRv/tIAa/rN/
G/+zfhr+sn0a/7F8Gf6xfBr/sHsZ/q13F/+gZg7+lloK/5RYCf6QVAn/RykE/gAAAP8AAAD3AAAATAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AFAAAAD3AAAA/zwkA/+VWwj/m2AJ/51iCf+hZwr/tX8U/8WSHv/GlCD/xJMg/8OSH//DkR//wpAf/8GPH//Bjx//wI4e/7+NHv+/jB7/vose/72KHv+9ih3/vIkd/7uIHf+6hx3/uoYd/7mGHP+4hRz/uIQc/7eD
HP+2ghv/toEb/7WBG/+0gBv/s38b/7N+Gv+yfRr/sXwZ/6hxEv+dYgr/mV0J/5ZaCv+RVgn/PCME/wAAAP8AAAD6AAAAWQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAATQAAAPMAAAD+JxgC/41XCP6dYgn/n2QI/qJnCP+obgr+uIIV/8SRHv7EkyD/w5If/sORH//CkB7+wY8f/8GPHv7Ajh7/v40e/r+MHv++ix3+vYoe/72KHf68iR3/u4gc/rqH
Hf+6hhz+uYYc/7iFHP64hBz/t4Mb/raCG/+2gRv+tYEb/7SAGv6zfhr/rHUT/qNpC/+eYwn+m2AJ/5ldCf6JUgn/KBgC/gAAAP8AAAD0AAAATwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5AAAA4wAAAP8QCgD/cEYG/59kCP+iZwj/pGoI/6ZsCP+rcgn/toAR/8CMGv/Dkh//w5Ef/8KQH//Bjx//wY8f/8COHv+/jR7/v4we/76L
Hv+9ih7/vYod/7yJHf+7iB3/uocd/7qGHf+5hhz/uIUc/7iEHP+3gxz/toIb/7R/F/+vdxD/qG4J/6RpCP+hZgj/nmMJ/5tgCP9vQwb/EQoB/wAAAP8AAADlAAAAPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABsAAAC8AAAA/gEAAP86JAP+kl0H/6RpB/6mbAj/qW8H/qxyB/+udAf+s3sK/7qEEf6/ixf/wY8c/sGP
Hv/Bjx7+wI4e/7+NHv6/jB7/vosd/r2KHv+9ih3+vIkd/7uIHP66hx3/uoYc/rmFGv+3ghX+tX4Q/7F4Cv6scwf/qm8H/qZsCP+kaQf+oWYI/5FaCP47JQP/AQAA/gAAAP8AAADAAAAAHAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwAAAHUAAADzAAAA/woGAP9TNQP/m2QH/6lvCP+scgf/rnQH/7F3
B/+0egf/tn0G/7mABv+8hAj/v4kM/8GLD//CjRH/w44S/8SPE//EjxP/w44S/8GMEf/Aig7/vocL/7yECf+5gAb/tn0G/7N5B/+vdgf/rHMH/6pvB/+nbAj/mWIH/1Q1BP8KBgD/AAAA/wAAAPQAAAB5AAAABAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAALAAAAD9AAAA/gwI
AP9OMwP+k2EG/650Bv6xdwf/s3oG/rZ9Bv+5gAX+vIMG/76GBf7CigX/xY0E/siQBP/LkwP+zZUE/8qSBP7HjwX/w4sE/r+HBf+8hAX+uYAG/7Z9Bv6zeQf/r3YG/qxyBv+TYQb+TzMD/w0IAP4AAAD/AAAA/QAA
ALMAAAAnAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAQAAAAL0AAAD+AAAA/wIBAP8qHAH/ZkQD/5xqBv+2fQb/uYAG/7yDBv+/hgX/wooF/8WNBf/JkAT/y5ME/82VBP/KkgT/x48F/8OLBf+/hwX/vIQG/7mABv+2fQb/nGoG/2dFBP8sHQH/AwEA/wAA
AP8AAAD+AAAAwAAAAEMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANAAAAKAAAADwAAAA/gAAAP8AAAD+GBEA/0IuAv5oSAP/h18D/qFzBP+ygAT+vYcE/8SOA/7GkAP/v4oD/rWBBP+idAT+iGAD/2hJA/5ELwL/GREA/gAA
AP8AAAD+AAAA/wAAAPEAAACjAAAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEQAAAFoAAACqAAAA6QAAAP0AAAD/AAAA/wAAAP8AAAD/BAMA/wkGAP8MCAD/DAgA/wkGAP8EAwD/AAAA/wAA
AP8AAAD/AAAA/wAAAP4AAADrAAAArAAAAF4AAAASAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYAAAAwAAAAZQAAAJMAAAC4AAAA2AAAAO0AAAD9AAAA/gAA
AP8AAAD9AAAA7QAAANgAAAC6AAAAlQAAAGcAAAAzAAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA///+AAB///////AAAA//////gAAAA/////8AAAAA/////AAAAAA////wAAAAAA///+AAAAAAB///wAAAAAAD//+AAAAAAAH//wAAAAAAAP/+AAAAAAAAf/wAAAAAAAA/+AAAAAAAAB/wAAAAAAAAD/AA
AAAAAAAP4AAAAAAAAAfgAAAAAAAAB8AAAAAAAAADwAAAAAAAAAOAAAAAAAAAAYAAAAAAAAABgAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAGAAAAAAAAAAYAA
AAAAAAABwAAAAAAAAAPAAAAAAAAAA+AAAAAAAAAH4AAAAAAAAAfwAAAAAAAAD/AAAAAAAAAP+AAAAAAAAB/8AAAAAAAAP/4AAAAAAAB//wAAAAAAAP//gAAAAAAB///AAAAAAAP//+AAAAAAB///8AAAAAAP///8
AAAAAD////8AAAAA/////8AAAAP/////8AAAD//////+AAB///8=
"@
#endregion ******** $Pause64Icon ********
$PILLargeImageList.Images.Add("Pause64Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Pause64Icon))))

#region ******** $Stop64Icon ********
$Stop64Icon = @"
AAABAAEAQEAAAAEAIAAoQgAAFgAAACgAAABAAAAAgAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAA2AAAAagAAAJYAAAC8AAAA2QAAAO4AAAD+AAAA/wAAAP8AAAD9AAAA7QAAANgAAAC5AAAAkwAAAGUAAAAwAAAABgAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVAAAAYgAAALEAAADtAAAA/gAAAP4AAAD/AAAA/gAAAP8DAgD+BwQA/wkFAf4JBQH/BgQA/gMCAP8AAAD+AAAA/wAAAP4AAAD/AAAA/QAA
AOkAAACoAAAAWgAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAADwAAACoAAAA9AAAAP8AAAD/AQAA/xYOAv85JAf/VzcK/25FDf+BUQ//i1cR/5BaEf+TWxL/kloS/45XEf+IUxH/fEwP/2lA
DP9SMQn/NB8G/xQMAv8AAAD/AAAA/wAAAP8AAADvAAAAnwAAADMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARwAAAMQAAAD/AAAA/gIBAP8pGgX+XjwM/4pZEf6dZRT/nWQT/pxjE/+bYhP+mmET/5phEv6ZYBP/mF8S/phe
E/+XXRL+llwT/5ZcEv6VWxL/lFoS/pRZEv+TWRH+f0wP/1QyCv4kFQT/AgEA/gAAAP8AAAD+AAAAuwAAAD4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArAAAAuAAAAP0AAAD/DQgB/00yCv+MWxL/oGgU/59nFP+fZhT/nmYU/51lFP+dZBT/nGMT/5ti
E/+aYRP/mmET/5lgE/+ZXxP/mF4T/5ddE/+WXBP/llwT/5VbEv+UWhL/lFkS/5NZEv+SWBH/kVYR/31KDv9CJwf/CgYB/wAAAP8AAAD8AAAArQAAACMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAB9AAAA9QAAAP4LBwH/VjkL/pllFP+iahX+oWoV/6FpFP6gaBX/n2cU/p9m
FP+eZhP+nWUU/51kE/6cYxP/m2IT/pphE/+aYRL+mWAT/5hfEv6YXhP/l10S/pZcE/+WXBL+lVsS/5RaEv6UWRL/k1kR/pJYEf+RVxH+kVYR/4ZPD/5JKwj/CAUB/gAAAP8AAADxAAAAcAAAAAMAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB4AAADCAAAA/wEBAP8/Kgj/l2QU/6RtFv+kbBb/o2sW/6Jq
Ff+hahX/oWkV/6BoFf+gZxT/n2YU/55mFP+dZBP/m2IT/5lgE/+YXhL/llwS/5ZcEv+WXBL/ll0S/5ddEv+XXRL/llwT/5ZcE/+VWxL/lFoS/5RZEv+TWRL/klgR/5JXEf+RVhH/kFUR/4JMD/81Hgb/AQAA/wAA
AP8AAAC5AAAAGQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD0AAADmAAAA/xMNAv55URD/pm8W/qZv
Fv+lbhb+pG0W/6RsFf6jaxb/omoV/qFqFf+gaBT+nWQU/5heEv6TWRH/kFUR/o9UEf+PVBD+j1QQ/49UEP6PVBD/j1QQ/o9UEP+PVBH+kFUR/5FWEf6TWRH/lFoS/pVbEv+UWhL+lFkS/5NZEf6SWBH/kVcR/pFW
Ef+QVRD+j1QQ/2U7C/4PCQH/AAAA/gAAAOEAAAA2AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFAAAAD0AAAA/y4f
Bv+aaBX/qHEX/6dwF/+mbxf/pm8W/6VuFv+kbRb/pGwW/6FpFf+aYRP/k1kR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+RVhH/kVYR/5FWEf+SVxH/k1kR/5Ra
Ev+UWRL/k1kS/5JYEf+SVxH/kVYR/5BVEf+PVBH/gUwP/yUVBP8AAAD/AAAA8QAAAEkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AFkAAAD6AAAA/0cwCf6mcRf/qXMX/qhyF/+ocRb+p3AX/6ZvFv6mbxb/pGwV/pthE/+UWRL+k1gS/5NYEf6TWBL/k1gR/pNYEv+TWBH+k1gS/5NYEf6TWBL/k1gR/pNYEv+TWBH+k1gS/5NYEf6TWBL/k1gR/pNY
Ev+TWBH+k1gS/5NYEf6TWBL/lFkR/pRZEv+TWRH+klgR/5FXEf6RVhH/kFUQ/o9UEf+LUhD+OSEG/wAAAP4AAAD3AAAAUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAE0AAAD3AAAA/1U6DP+qdBf/q3QY/6pzGP+pcxf/qHIX/6hxF/+ncBb/n2cU/5ZcEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5Ra
Ev+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5RaEv+UWhL/lFoS/5NZEv+SWBH/klcR/5FWEf+QVRH/j1QR/45TEP9DJwf/AAAA/wAAAPYAAABFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADIAAADwAAAA/1I4C/6sdhj/rHYY/qt1GP+qdBf+qnMY/6lzF/6ncBb/m2IU/pZcE/+WXBL+llwT/5ZcEv6WXBP/llwS/pZcE/+WXBL+llwT/5ZcEv6WXBP/llwS/pZc
E/+WXBL+llwT/5ZcEv6WXBP/llwS/pZcE/+WXBL+llwT/5ZcEv6WXBP/llwS/pZcE/+WXBL+llwT/5ZcEv6VWxL/k1kS/pJYEf+RVxH+kVYR/5BVEP6PVBH/jlMQ/kEmB/8AAAD+AAAA7gAAACwAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIAAADdAAAA/0EtCf+tdxj/rXgZ/613Gf+sdhj/q3UY/6p0GP+ncBf/mmET/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5he
E/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5heE/+YXhP/mF4T/5deEv+TWRL/klgR/5JXEf+RVhH/kFUR/49UEf+NUxD/NB4G/wAA
AP8AAADZAAAADwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAACrAAAA/ycbBf6pdhj/r3kZ/q55Gf+teBj+rXcZ/6x2GP6ncRb/m2IT/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppg
E/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mmAT/5pgEv6aYBP/mmAS/ppgE/+aYBL+mWAS/5RaEv6SWBH/kVcR/pFW
Ef+QVRD+j1QR/4lQEP4fEgP/AAAA/gAAAKUAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABaAAAA/QsIAf+Zaxb/sHsa/696Gf+veRn/rnkZ/614Gf+pcxf/nGMT/5tiE/+bYhP/m2IT/5ti
E/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5tiE/+bYhP/m2IT/5ti
E/+aYRP/lFoS/5JYEf+SVxH/kVYR/5BVEf+PVBH/e0gO/wkFAf8AAAD9AAAAVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAA5QAAAP5sTA//sXwZ/rF8Gv+wexn+r3oZ/695Gf6tdxj/n2YU/p1k
FP+dZBP+nWQU/51kE/6dZBT/nWQT/p1kFP+dZBP+nWQU/51kE/6dZBT/nWQT/p1kFP+dZBP+nWQU/51kE/6dZBT/nWQT/p1kFP+dZBP+nWQU/51kE/6dZBT/nWQT/p1kFP+dZBP+nWQU/51kE/6dZBT/nWQT/p1k
FP+dZBP+nWQU/51kE/6dZBT/nWQT/pxjE/+UWRL+klgR/5FXEf6RVhH/kFUQ/o9UEf9VMgr+AAAA/wAAAOIAAAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhwAAAP8rHgb/sn4a/7J9Gv+xfBr/sXwa/7B7
Gv+vehn/o2sV/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59m
FP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/n2YU/59mFP+fZhT/m2IT/5NZEv+SWBH/klcR/5FWEf+QVRH/jlQQ/yITA/8AAAD/AAAAgwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFwAAAPICAQD/jGMU/rN/
G/+zfhr+sn0a/7F8Gf6xfBr/qXMX/qBoFf+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/6BoFP6gaBX/oGgU/qBo
Ff+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/6BoFP6gaBX/oGgU/qBoFf+gaBT+oGgV/6BoFP6ZXxP/k1kR/pJYEf+RVxH+kVYR/5BVEP5wQg3/AQEA/gAAAPIAAAAWAAAAAAAAAAAAAAAAAAAAAAAA
AIEAAAD/NycI/7WAGv+0gBv/s38b/7N+Gv+yfRr/sHsZ/6NrFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6BnFP+TWBH/i08P/4lNDv+JTQ7/iU0O/4lNDv+JTQ7/iU0O/4lN
Dv+JTQ7/iU0O/4lNDv+JTQ7/iU0O/4tPD/+TWBH/oGcU/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/omoV/6JqFf+iahX/oWkU/5VbEv+TWRL/klgR/5JXEf+RVhH/kFUQ/ysZBf8AAAD/AAAAgAAA
AAAAAAAAAAAAAAAAAAkAAADlAAAA/otjFf+2gRv+tYEb/7SAGv6zfxv/s34a/qpzF/+jbBX+o2wW/6NsFf6jbBb/o2wV/qNsFv+jbBX+o2wW/6NsFf6jbBb/o2wV/p5mFP+MURH+pnlK/8aqjf7OtJv/zbSb/s60
m//NtJv+zrSb/820m/7OtJv/zbSb/s60m//NtJv+zrSb/820m/7Hqo3/pnlK/oxREf+eZhT+o2wW/6NsFf6jbBb/o2wV/qNsFv+jbBX+o2wW/6NsFf6jbBb/o2wV/qNsFv+dZBT+lFkS/5NZEf6SWBH/kVcR/pFW
Ef9uQQz+AAAA/wAAAOQAAAAJAAAAAAAAAAAAAABPAAAA/yIYBf+1ghv/toIb/7aBG/+1gRv/tIAb/7J+Gv+mbxb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/6RtFf+YYSH/3Mq3//7+
/f////////////////////////////////////////////////////////////////////////////7+/f/cyrf/mGEh/6RtFf+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pW4W/6VuFv+lbhb/pG0V/5Zc
Ev+UWRL/k1kS/5JYEf+SVxH/j1UQ/xsQA/8AAAD/AAAATwAAAAAAAAAAAAAApAAAAP9jRw/+uIQc/7eDG/62ghv/toEb/rWBG/+ueBn+p3AX/6dwFv6ncBf/p3AW/qdwF/+ncBb+p3AX/6dwFv6ncBf/p3AW/qdw
F/+fZxb+38+7//7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+/////9/Pu/6fZxb/p3AW/qdwF/+ncBb+p3AX/6dwFv6ncBf/p3AW/qdw
F/+ncBb+p3AX/6dwFv6dZBT/lFoS/pRZEv+TWRH+klgR/5FXEf5OLwn/AAAA/gAAAKYAAAAAAAAABwAAAOoAAAD/n3IY/7iFHP+4hBz/t4Mc/7aCG/+2gRv/q3QX/6hxF/+ocRf/qHEX/6hxF/+ocRf/qHEX/6hx
F/+ocRf/qHEX/6hxF/+ocRf/u5NZ////////////////////////////////////////////////////////////////////////////////////////////////////////////u5NZ/6hxF/+ocRf/qHEX/6hx
F/+ocRf/qHEX/6hxF/+ocRf/qHEX/6hxF/+ocRf/pG0W/5VbEv+UWhL/lFkS/5NZEv+SWBH/fksO/wAAAP8AAADrAAAABwAAADQAAAD+GhME/7mGHP65hhz/uIUc/riEHP+3gxv+tYAb/6pzGP6qcxj/qnMX/qpz
GP+qcxf+qnMY/6pzF/6qcxj/qnMX/qpzGP+qcxf+qnMY/9vEoP7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/tvE
oP+qcxf+qnMY/6pzF/6qcxj/qnMX/qpzGP+qcxf+qnMY/6pzF/6qcxj/qnMX/qpzF/+YXhP+lVsS/5RaEv6UWRL/k1kR/pJYEf8WDQL+AAAA/gAAADcAAABrAAAA/0czCv+6hx3/uoYd/7mGHP+4hRz/uIQc/7N+
Gv+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP/j0bL/////////////////////////////////////////////////////////////////////////////////////////
///////////////////j0bL/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/q3UY/6t1GP+rdRj/nmUU/5ZcE/+VWxL/lFoS/5RZEv+TWRL/OiMH/wAAAP8AAABvAAAAmwAAAP5sTxD/u4gc/rqH
Hf+6hhz+uYYc/7iFHP6yfRr/rXcY/q13Gf+tdxj+rXcZ/613GP6tdxn/rXcY/q13Gf+tdxj+rXcZ/613GP6tdxn/49Gy/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v//
///+/v7+//////7+/v7//////v7+/v/////+/v7+49Gy/613GP6tdxn/rXcY/q13Gf+tdxj+rXcZ/613GP6tdxn/rXcY/q13Gf+tdxj+rXcZ/6NrFv6WXBP/llwS/pVbEv+UWhL+lFkS/1c0Cv4AAAD/AAAAnwAA
AMEAAAD/i2YV/7yJHf+7iB3/uocd/7qGHf+5hhz/sXwZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/+TSsv//////////////////////////////////////////////
/////////////////////////////////////////////////////////////+TSsv+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ueRn/rnkZ/655Gf+ocRf/l10T/5ZcE/+WXBP/lVsS/5Va
Ev9wQw3/AAAA/wAAAMYAAADdAAAA/qJ2Gf+9ih3+vIkd/7uIHP66hx3/uoYc/rJ9Gv+wexn+sHsZ/7B7Gf6wexn/sHsZ/rB7Gf+wexn+sHsZ/7B7Gf6wexn/sHsZ/rB7Gf/k07L+//////7+/v7//////v7+/v//
///+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7k07L/sHsZ/rB7Gf+wexn+sHsZ/7B7Gf6wexn/sHsZ/rB7Gf+wexn+sHsZ/7B7Gf6wexn/q3UY/phe
E/+XXRL+llwT/5ZcEv6VWxL/gk8P/gEAAP8AAADjAAAA7wUEAP+ufxv/vYoe/72KHf+8iR3/u4gd/7qHHf+yfhr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/5dOy////
////////////////////////////////////////////////////////////////////////////////////////////////////////5dOy/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8Gv+xfBr/sXwa/7F8
Gv+xfBr/sXwa/655Gf+ZXxP/mF4T/5ddE/+WXBP/llwT/4pUEf8FAwD/AAAA9wAAAP4HBQH+soMc/76LHf69ih7/vYod/ryJHf+7iBz+tH8b/7N+Gv6zfhr/s34a/rN+Gv+zfhr+s34a/7N+Gv6zfhr/s34a/rN+
Gv+zfhr+s34a/+XUsv7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/uXUsv+zfhr+s34a/7N+Gv6zfhr/s34a/rN+
Gv+zfhr+s34a/7N+Gv6zfhr/s34a/rN+Gv+wexn+mWAT/5hfEv6YXhP/l10S/pZcE/+QWBL+CAUB/wAAAP4AAAD/CAYB/7WFHP+/jB7/vose/72KHv+9ih3/vIkd/7WBG/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SA
G/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SAG//m1LP////////////////////////////////////////////////////////////////////////////////////////////////////////////m1LP/tIAb/7SA
G/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SAG/+0gBv/tIAb/7SAG/+0gBv/sHsZ/5phE/+ZYBP/mV8T/5heE/+XXRP/kVkS/wkFAf8AAAD/AAAA/woHAf64hx3/v40e/r+MHv++ix3+vYoe/72KHf63gxv/toEb/raB
G/+2gRv+toEb/7aBG/62gRv/toEb/raBG/+2gRv+toEb/7aBG/62gRv/5tWz/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v//
///+/v7+5tWz/7aBG/62gRv/toEb/raBG/+2gRv+toEb/7aBG/62gRv/toEb/raBG/+2gRv+toEb/696Gf6aYRP/mmES/plgE/+YXxL+mF4T/5JaEv4IBQH/AAAA/gAAAP8IBQD/toUc/8COHv+/jR7/v4we/76L
Hv+9ih7/uYYc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/+fVs///////////////////////////////////////////////////////////////////////////////
/////////////////////////////+fVs/+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+3gxz/t4Mc/7eDHP+tdxj/m2IT/5phE/+aYRP/mWAT/5lfE/+SWhH/CAUA/wAAAP8AAAD/CAQA/rKA
Gv/Bjx7+wI4e/7+NHv6/jB7/vosd/ruIHf+4hRz+uYUc/7iFHP65hRz/uIUc/rmFHP+4hRz+uYUc/7iFHP65hRz/uIUc/rmFHP/n1rP+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v//
///+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7n1rP/uIUc/rmFHP+4hRz+uYUc/7iFHP65hRz/uIUc/rmFHP+4hRz+uYUc/7iFHP65hRz/qXMX/pxjE/+bYhP+mmET/5phEv6ZYBP/jlcR/gYD
AP8AAAD+AAAA/wcEAP+qdxf/wY8f/8GPH//Ajh7/v40e/7+MHv+9ih3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/6Nez////////////////////////////////////
////////////////////////////////////////////////////////////////////////6Nez/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/7qHHf+6hx3/uocd/6RsFf+dZBT/nGMT/5ti
E/+aYRP/mmET/4lTEP8FAgD/AAAA/gAAAPcEAgD+m2gT/8KQHv7Bjx//wY8e/sCOHv+/jR7+v4wd/7yIHf68iB3/vIgc/ryIHf+8iBz+vIgd/7yIHP68iB3/vIgc/ryIHf+8iBz+vIgd/+jXs/7//////v7+/v//
///+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/ujXs/+8iBz+vIgd/7yIHP68iB3/vIgc/ryIHf+8iBz+vIgd/7yIHP68iB3/vIgc/rmF
HP+fZxT+nWUU/51kE/6cYxP/m2IT/pphE/+CTQ7+AwEA/wAAAO8AAADjAAAA/4dVD//DkR//wpAf/8GPH//Bjx//wI4e/7+NHv++ix3/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72K
Hf/l0qn////////////////////////////////////////////////////////////////////////////////////////////////////////////m0qn/vYod/72KHf+9ih3/vYod/72KHf+9ih3/vYod/72K
Hf+9ih3/vYod/72KHf+xfBr/n2YU/55mFP+dZRT/nWQU/5xjE/+bYhP/dkIM/wAAAP8AAADcAAAAxQAAAP5qPQn/wY8f/sORH//CkB7+wY8f/8GPHv7Ajh7/v4we/r6LHv++ix3+vose/76LHf6+ix7/vosd/r6L
Hv++ix3+vose/76LHf6+ix7/1bVu/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+1bVu/76LHf6+ix7/vosd/r6L
Hv++ix3+vose/76LHf6+ix7/vosd/r6LHv++ix3+p3AW/59nFP6fZhT/nmYT/p1lFP+dZBP+nGMT/2AzCf4AAAD/AAAAwAAAAJ4AAAD/Ui0G/7N+Gf/Dkh//w5Ef/8KQH//Bjx//wY8f/8COHv+/jR7/v40e/8CN
Hv+/jR7/wI0e/7+NHv/AjR7/v40e/8CNHv+/jR7/wI0e/8GQJP/z6tX/////////////////////////////////////////////////////////////////////////////////////////////////8+rV/8GQ
JP+/jR7/wI0e/7+NHv/AjR7/v40e/8CNHv+/jR7/wI0e/7+NHv/AjR7/uYUc/6FpFf+gaBX/oGcU/59mFP+eZhT/nWUU/5ZcEv9KJwb/AAAA/wAAAJsAAABwAAAA/jYeBP+gZxL+xJMg/8OSH/7DkR//wpAe/sGP
H//Bjx7+wI4f/8GPHv7Bjx//wY8e/sGPH//Bjx7+wY8f/8GPHv7Bjx//wY8e/sGPH//Bjx7+yZ48//Tr1/7//////v7+/v/////+/v7+//////7+/v7//////v7+/v/////+/v7+//////7+/v7//////v7+/v//
///+/v7+9OvX/8mdPP7Bjx//wY8e/sGPH//Bjx7+wY8f/8GPHv7Bjx//wY8e/sGPH//Bjx7+wY4e/6p0F/6hahX/oWkU/qBoFf+fZxT+n2YU/55mE/6NUg//MBkE/gAAAP8AAABrAAAANwAAAP4UCwH/jVEM/8GO
Hv/EkyD/w5If/8ORH//CkB//wY8f/8GPH//Bjx//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/8KQH//FlSj/2715/+zcuP/v48b/7+PG/+/jxv/v48b/7+PG/+/jxv/v48b/7+PG/+/j
xv/v48b/7+PG/+/jxv/s3Lj/2715/8WVKP/CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/8KQH//CkB//wpAf/7iEG/+jaxb/omoV/6FqFf+haRX/oGgV/59nFP+dZRP/hEcM/xIJAf8AAAD+AAAANAAA
AAgAAADrAAAA/3hDCf6sdBb/xpMf/sSTIP/Dkh/+w5Ef/8KQHv7Bjx//wY8f/sORH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOS
H//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8ORH/7Dkh//w5Ef/sOSH//DkR/+w5If/8GPHv6ocRb/pGwV/qNrFv+iahX+oWoV/6FpFP6gaBX/lFoR/m86
Cv8AAAD+AAAA6QAAAAYAAAAAAAAApwAAAP9MKgX/klYN/8SRHv/GlCD/xJMg/8OSH//DkR//wpAf/8GPH//CkB//xZMf/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WT
IP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTIP/FkyD/xZMg/8WTH/+vehn/pW4W/6RtFv+kbBb/o2sW/6Jq
Ff+hahX/oGgU/4hLDf9GJQb/AAAA/wAAAKMAAAAAAAAAAAAAAFEAAAD/Gg4B/opOCv+rdBb+x5Qg/8aTH/7EkyD/w5If/sORH//CkB7+wY8f/8ORH/7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seU
IP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP7HlCD/x5Qg/seUIP/HlCD+x5Qg/8eUIP62gRv/pm8W/qZv
Fv+lbhb+pG0W/6RsFf6jaxb/omoV/pVbEP+CRQv+FwwC/wAAAP4AAABOAAAAAAAAAAAAAAAJAAAA5QAAAP9qPAj/kFQM/8GNHv/HlCD/xpQg/8STIP/Dkh//w5Ef/8KQH//Bjx//xJIf/8iWIP/IliH/yJYh/8iW
If/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iWIf/IliH/yJYh/8iW
If+5hRz/qHEX/6dwF/+mbxf/pm8W/6VuFv+kbRb/pGwW/6BoFf+HSgv/ZDUJ/wAAAP8AAADkAAAACQAAAAAAAAAAAAAAAAAAAIEAAAD+KhcD/4xQCv6fZhH/x5Qg/seUIP/Gkx/+xJMg/8OSH/7DkR//wpAe/sGP
H//Fkx/+yZcg/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smXIf/JlyD+yZch/8mXIP7JlyH/yZcg/smX
If/JlyD+yZch/8mXIP65hRz/qXMX/qhyF/+ocRb+p3AX/6ZvFv6mbxb/pW4W/qRtFf+RVg7+hUcL/ycVA/4AAAD/AAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAXAAAA8gEBAP9tPgj/jlIK/7B5F//HlSD/x5Qg/8aU
IP/EkyD/w5If/8ORH//CkB//wY8f/8SSH//KmCH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZIf/KmSH/ypkh/8qZ
If/KmSH/ypkh/8qZIf/KmSH/ypkh/8mXIf+3ghv/q3QY/6pzGP+pcxf/qHIX/6hxF/+ncBf/pm8X/6ZvFv+bYRH/h0kL/2c4CP8BAAD/AAAA8wAAABYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIYAAAD/IRIC/o1R
Cv+SVgr+u4cb/8eVIP7HlCD/xpMf/sSTIP/Dkh/+w5Ef/8KQHv7Bjx//wpAf/smXIf/MmiH+zJoi/8yaIf7MmiL/zJoh/syaIv/MmiH+zJoi/8yaIf7MmiL/zJoh/syaIv/MmiH+zJoi/8yaIf7MmiL/zJoh/sya
Iv/MmiH+zJoi/8yaIf7MmiL/zJoh/syaIv/MmiH+zJoi/8WTIP6yfRr/rHYY/qt1GP+qdBf+qnMY/6lzF/6ochf/qHEW/qdwF/+haRT+ik4L/4ZJCv4fEQL/AAAA/gAAAIYAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAPAAAA5AAAAP9VMQb/kFMK/5ZbDP/AjB3/x5Ug/8eUIP/GlCD/xJMg/8OSH//DkR//wpAf/8GPH//Bjx//xZMg/8uaIf/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/zZwi/82c
Iv/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/zZwi/82cIv/NnCL/ypgh/7uIHP+veRn/rXgZ/613Gf+sdhj/q3UY/6t0GP+qcxj/qXMX/6hyF/+lbRX/jlIM/4lMC/9SLQb/AAAA/wAAAOQAAAAPAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFcAAAD9CQUA/3tHCP6RVQr/ml8M/sKOHv/HlSD+x5Qg/8aTH/7EkyD/w5If/sORH//CkB7+wY8f/8GPHv7Bjh//xZQf/syaIf/OnSL+zp0i/86dIv7OnSL/zp0i/s6d
Iv/OnSL+zp0i/86dIv7OnSL/zp0i/s6dIv/OnSL+zp0i/86dIv7OnSL/zp0i/s6dIv/KmCD+vosd/7J9Gv6vehn/r3kZ/q55Gf+teBj+rXcZ/6x2GP6rdRj/qnQX/qpzGP+ncBb+klYN/4tOCv52QQn/CAQA/gAA
AP0AAABZAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAqAAAAP8fEgL/i1EJ/5NXCv+cYQ3/wY0d/8eVIP/HlCD/xpQg/8STIP/Dkh//w5Ef/8KQH//Bjx//wY8f/8COHv/AjR7/w5Ef/8iW
IP/NnCL/z54i/8+eI//PniP/z54j/8+eI//PniP/z54j/8+eI//PniP/z54j/8+eI//PniL/zJsi/8SSH/+6hx3/s34a/7F8Gv+xfBr/sHsa/696Gf+veRn/rnkZ/614Gf+tdxn/rHYY/6t1GP+ncRb/lFgM/41Q
C/+GSwr/HhAC/wAAAP8AAACqAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABEAAADbAAAA/jUfA/+SVgn+lVkK/5tgC/68iBv/x5Ug/seUIP/Gkx/+xJMg/8OSH/7DkR//wpAe/sGP
H//Bjx7+wI4e/7+NHv6/jB7/vosd/sCNHv/DkB/+xZMf/8eWIP7JlyD/ypgh/sqYIf/IlyD+xpQg/8ORH/6/jB7/uYYc/rWBG/+0gBr+s38b/7N+Gv6yfRr/sXwZ/rF8Gv+wexn+r3oZ/695Gf6ueRn/rXgY/q13
Gf+ncBX+lFkL/49TCv6MTwr/Mx0E/gAAAP8AAADcAAAAEgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALwAAAO8AAAD/QygE/5RYCf+XWwn/ml8J/7R+Fv/HlCD/x5Qg/8aU
IP/EkyD/w5If/8ORH//CkB//wY8f/8GPH//Ajh7/v40e/7+MHv++ix7/vYoe/72KHf+8iR3/u4gd/7qHHf+6hh3/uYYc/7iFHP+4hBz/t4Mc/7aCG/+2gRv/tYEb/7SAG/+zfxv/s34a/7J9Gv+xfBr/sXwa/7B7
Gv+vehn/r3kZ/654GP+kbBP/lVkK/5FVCv+PUgn/QiYE/wAAAP8AAADwAAAAMQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABJAAAA9wAAAP5FKQT/lloJ/pld
Cf+bYAn+qXAP/8KPHf7HlCD/xpMf/sSTIP/Dkh/+w5Ef/8KQHv7Bjx//wY8e/sCOHv+/jR7+v4we/76LHf69ih7/vYod/ryJHf+7iBz+uocd/7qGHP65hhz/uIUc/riEHP+3gxv+toIb/7aBG/61gRv/tIAa/rN/
G/+zfhr+sn0a/7F8Gf6xfBr/sHsZ/q13F/+gZg7+lloK/5RYCf6QVAn/RykE/gAAAP8AAAD3AAAATAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AFAAAAD3AAAA/zwkA/+VWwj/m2AJ/51iCf+hZwr/tX8U/8WSHv/GlCD/xJMg/8OSH//DkR//wpAf/8GPH//Bjx//wI4e/7+NHv+/jB7/vose/72KHv+9ih3/vIkd/7uIHf+6hx3/uoYd/7mGHP+4hRz/uIQc/7eD
HP+2ghv/toEb/7WBG/+0gBv/s38b/7N+Gv+yfRr/sXwZ/6hxEv+dYgr/mV0J/5ZaCv+RVgn/PCME/wAAAP8AAAD6AAAAWQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAATQAAAPMAAAD+JxgC/41XCP6dYgn/n2QI/qJnCP+obgr+uIIV/8SRHv7EkyD/w5If/sORH//CkB7+wY8f/8GPHv7Ajh7/v40e/r+MHv++ix3+vYoe/72KHf68iR3/u4gc/rqH
Hf+6hhz+uYYc/7iFHP64hBz/t4Mb/raCG/+2gRv+tYEb/7SAGv6zfhr/rHUT/qNpC/+eYwn+m2AJ/5ldCf6JUgn/KBgC/gAAAP8AAAD0AAAATwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5AAAA4wAAAP8QCgD/cEYG/59kCP+iZwj/pGoI/6ZsCP+rcgn/toAR/8CMGv/Dkh//w5Ef/8KQH//Bjx//wY8f/8COHv+/jR7/v4we/76L
Hv+9ih7/vYod/7yJHf+7iB3/uocd/7qGHf+5hhz/uIUc/7iEHP+3gxz/toIb/7R/F/+vdxD/qG4J/6RpCP+hZgj/nmMJ/5tgCP9vQwb/EQoB/wAAAP8AAADlAAAAPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABsAAAC8AAAA/gEAAP86JAP+kl0H/6RpB/6mbAj/qW8H/qxyB/+udAf+s3sK/7qEEf6/ixf/wY8c/sGP
Hv/Bjx7+wI4e/7+NHv6/jB7/vosd/r2KHv+9ih3+vIkd/7uIHP66hx3/uoYc/rmFGv+3ghX+tX4Q/7F4Cv6scwf/qm8H/qZsCP+kaQf+oWYI/5FaCP47JQP/AQAA/gAAAP8AAADAAAAAHAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwAAAHUAAADzAAAA/woGAP9TNQP/m2QH/6lvCP+scgf/rnQH/7F3
B/+0egf/tn0G/7mABv+8hAj/v4kM/8GLD//CjRH/w44S/8SPE//EjxP/w44S/8GMEf/Aig7/vocL/7yECf+5gAb/tn0G/7N5B/+vdgf/rHMH/6pvB/+nbAj/mWIH/1Q1BP8KBgD/AAAA/wAAAPQAAAB5AAAABAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAALAAAAD9AAAA/gwI
AP9OMwP+k2EG/650Bv6xdwf/s3oG/rZ9Bv+5gAX+vIMG/76GBf7CigX/xY0E/siQBP/LkwP+zZUE/8qSBP7HjwX/w4sE/r+HBf+8hAX+uYAG/7Z9Bv6zeQf/r3YG/qxyBv+TYQb+TzMD/w0IAP4AAAD/AAAA/QAA
ALMAAAAnAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAQAAAAL0AAAD+AAAA/wIBAP8qHAH/ZkQD/5xqBv+2fQb/uYAG/7yDBv+/hgX/wooF/8WNBf/JkAT/y5ME/82VBP/KkgT/x48F/8OLBf+/hwX/vIQG/7mABv+2fQb/nGoG/2dFBP8sHQH/AwEA/wAA
AP8AAAD+AAAAwAAAAEMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANAAAAKAAAADwAAAA/gAAAP8AAAD+GBEA/0IuAv5oSAP/h18D/qFzBP+ygAT+vYcE/8SOA/7GkAP/v4oD/rWBBP+idAT+iGAD/2hJA/5ELwL/GREA/gAA
AP8AAAD+AAAA/wAAAPEAAACjAAAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEQAAAFoAAACqAAAA6QAAAP0AAAD/AAAA/wAAAP8AAAD/BAMA/wkGAP8MCAD/DAgA/wkGAP8EAwD/AAAA/wAA
AP8AAAD/AAAA/wAAAP4AAADrAAAArAAAAF4AAAASAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYAAAAwAAAAZQAAAJMAAAC4AAAA2AAAAO0AAAD9AAAA/gAA
AP8AAAD9AAAA7QAAANgAAAC6AAAAlQAAAGcAAAAzAAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA///+AAB///////AAAA//////gAAAA/////8AAAAA/////AAAAAA////wAAAAAA///+AAAAAAB///wAAAAAAD//+AAAAAAAH//wAAAAAAAP/+AAAAAAAAf/wAAAAAAAA/+AAAAAAAAB/wAAAAAAAAD/AA
AAAAAAAP4AAAAAAAAAfgAAAAAAAAB8AAAAAAAAADwAAAAAAAAAOAAAAAAAAAAYAAAAAAAAABgAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAGAAAAAAAAAAYAA
AAAAAAABwAAAAAAAAAPAAAAAAAAAA+AAAAAAAAAH4AAAAAAAAAfwAAAAAAAAD/AAAAAAAAAP+AAAAAAAAB/8AAAAAAAAP/4AAAAAAAB//wAAAAAAAP//gAAAAAAB///AAAAAAAP//+AAAAAAB///8AAAAAAP///8
AAAAAD////8AAAAA/////8AAAAP/////8AAAD//////+AAB///8=
"@
#endregion ******** $Stop64Icon ********
$PILLargeImageList.Images.Add("Stop64Icon", [System.Drawing.Icon]::New([System.IO.MemoryStream]::New([System.Convert]::FromBase64String($Stop64Icon))))

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

  #Write-KPIEvent -Source "Utility" -EntryType "Information" -EventID 2 -Category 0 -Message "Exiting $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"

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
          $Script:VerbosePreference = "SilentlyContinue"
          $Script:DebugPreference = "SilentlyContinue"
          [System.Console]::Title = "RUNNING: $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
          [Void][Console.Window]::Hide()
          $PILForm.Tag = $False
        }
        Else
        {
          # Show Console Window
          $Script:VerbosePreference = "Continue"
          $Script:DebugPreference = "Continue"
          [Void][Console.Window]::Show()
          [System.Console]::Title = "DEBUG: $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
          $PILForm.Tag = $True
        }
        $PILForm.Activate()
        $PILForm.Select()
        Break
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

  $HashTable = @{"ShowHeader" = $True; "ConfigFile" = $ConfigFile; "ExportFile" = $ExportFile}
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

#region ******** Controls for PIL Form ********

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
      
      $PILItemListContextMenuStrip.Items["Process"].Text = $PILItemListContextMenuStrip.Items["Process"].Tag -f $TmpMenuText
      $PILItemListContextMenuStrip.Items["Export"].Text = $PILItemListContextMenuStrip.Items["Export"].Tag -f $TmpMenuText
      $PILItemListContextMenuStrip.Items["Clear"].Text = $PILItemListContextMenuStrip.Items["Clear"].Tag -f $TmpMenuText
      
      $PILItemListContextMenuStrip.Show($Sender, $EventArg.Location)
    }
  }

  Write-Verbose -Message "Exit MouseDown Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItemListListViewMouseDown ********
$PILItemListListView.add_MouseDown({Start-PILItemListListViewMouseDown -Sender $This -EventArg $PSItem})

For ($I = 0; $I -lt $MaxColumns; $I++)
{
  New-ColumnHeader -ListView $PILItemListListView -Text ([MyRuntime]::ThreadConfig.ColumnNames[$I]) -Name ([MyRuntime]::ThreadConfig.ColumnNames[$I]) -Tag ([MyRuntime]::ThreadConfig.ColumnNames[$I])
}
New-ColumnHeader -ListView $PILItemListListView -Text " " -Name "Blank" -Tag " " -Width ($PILForm.Width * 4)

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
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Process $($TmpListText) Items"
      $PILBtmStatusStrip.Refresh()
      
      # *****************************************
      # **** Testing - Exit to Nested Prompt ****
      # *****************************************
      Write-Host -Object "Line Num: $((Get-PSCallStack).ScriptLineNumber)"
      $Host.EnterNestedPrompt()
      # *****************************************
      # **** Testing - Exit to Nested Prompt ****
      # *****************************************
      
      Break
    }
    "Export"
    {
      #region Export Slected / Checked Items
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Export $($TmpListText) Items"
      $PILBtmStatusStrip.Refresh()
      
      # Save Export File
      $PILSaveFileDialog.FileName = ""
      $PILSaveFileDialog.Filter = "CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
      $PILSaveFileDialog.FilterIndex = 1
      $PILSaveFileDialog.Title = "Export PIL CSV Report"
      $PILSaveFileDialog.Tag = $Null
      $Response = $PILSaveFileDialog.ShowDialog()
      If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
      {
        $TmpCount = ([MyRuntime]::MaxColumns - 1)
        $StringBuilder = [System.Text.StringBuilder]::New()
        [Void]$StringBuilder.AppendLine(($PILItemListListView.Columns[0..$($TmpCount)] | Select-Object -ExpandProperty Text) -Join ",")
        $TmpLisTViewItems | ForEach-Object -Process { [Void]$StringBuilder.AppendLine("`"{0}`"" -f (($PSItem.SubItems[0 .. $($TmpCount)] | Select-Object -ExpandProperty Text) -join "`",`"")) }
        ConvertFrom-Csv -InputObject (($StringBuilder.ToString())) -Delimiter "," | Export-Csv -Path $PILSaveFileDialog.FileName -NoTypeInformation -Encoding ASCII
        $StringBuilder.Clear()
        
        # Save Current Directory
        $PILSaveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILSaveFileDialog.FileName)
        $PILBtmStatusStrip.Items["Status"].Text = "Success Exporting $($TmpListText) Items"
      }
      Else
      {
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Exporting $($TmpListText) Items"
      }
      Break
      #endregion Export Slected / Checked Items
    }
    "Clear"
    {
      #region Clear Selected / Checked Item List
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Clear $($TmpListText) Items"
      $PILBtmStatusStrip.Refresh()
      
      $DialogResult = Get-UserResponse -Title "Clear Item List?" -Message "Do you want to Clear the $($TmpListText) Items?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
      If ($DialogResult.Success)
      {
        # Clear Item List
        $TmpLisTViewItems | ForEach-Object { $PILItemListListView.Items.Remove($PSItem) }
        
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Successfully Cleared $($TmpListText) Items"
      }
      Else
      {
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Canceled Clearing $($TmpListText) Items"
      }
      Break
      #endregion Clear Selected / Checked Item List
    }
  }
  
  Write-Verbose -Message "Exit ItemClicked Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItemListContextMenuStripItemClick ********

(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Process" -Name "Process" -Tag "Process {0} Items" -DisplayStyle "ImageAndText" -ImageKey "Process16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Export" -Name "Export" -Tag "Export {0} Items" -DisplayStyle "ImageAndText" -ImageKey "Export16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $PILItemListContextMenuStrip -Text "Clear" -Name "Clear" -Tag "Clear {0} Items" -DisplayStyle "ImageAndText" -ImageKey "Clear16Icon" -PassThru).add_Click({Start-PILItemListContextMenuStripItemClick -Sender $This -EventArg $PSItem})

# ************************************************
# PILItelList ToolStrip
# ************************************************
#region $PILItelListToolStrip = [System.Windows.Forms.ToolStrip]::New()
$PILItelListToolStrip = [System.Windows.Forms.ToolStrip]::New()
$PILMainPanel.Controls.Add($PILItelListToolStrip)
#$PILForm.ToolStrip = $PILItelListToolStrip
$PILItelListToolStrip.Anchor = [System.Windows.Forms.AnchorStyles]("Top")
$PILItelListToolStrip.AutoSize = $True
$PILItelListToolStrip.BackColor = [MyConfig]::Colors.Fore
$PILItelListToolStrip.Dock = [System.Windows.Forms.DockStyle]::None
$PILItelListToolStrip.Font = [MyConfig]::Font.Regular
$PILItelListToolStrip.ForeColor = [MyConfig]::Colors.Fore
$PILItelListToolStrip.GripStyle = [System.Windows.Forms.ToolStripGripStyle]::Hidden
$PILItelListToolStrip.ImageList = $PILLargeImageList
$PILItelListToolStrip.ImageScalingSize = [System.Drawing.Size]::New(64, 64)
$PILItelListToolStrip.Name = "PILItelListToolStrip"
$PILItelListToolStrip.ShowItemToolTips = $True
$PILItelListToolStrip.Stretch = $True
#$PILItelListToolStrip.TabIndex = 0
#$PILItelListToolStrip.TabStop = $False
#$PILItelListToolStrip.Tag = [System.Object]::New()
#endregion $PILItelListToolStrip = [System.Windows.Forms.ToolStrip]::New()

$PILItelListToolStrip.SendToBack()

#region ******** Function Start-PILItelListToolStripItemClick ********
function Start-PILItelListToolStripItemClick
{
  <#
    .SYNOPSIS
      Click Event for the PILItelList ToolStripItem Control
    .DESCRIPTION
      Click Event for the PILItelList ToolStripItem Control
    .PARAMETER Sender
       The ItelList Control that fired the Click Event
    .PARAMETER EventArg
       The Event Arguments for the ItelList Click Event
    .EXAMPLE
       Start-PILItelListToolStripItemClick -Sender $Sender -EventArg $EventArg
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
  
  If (($Sender.CheckState -eq "Unchecked") -and ($Sender.Name -in ("Process", "Pause")))
  {
    $Sender.Checked = $True
  }
  Else
  {
    Switch ($Sender.Name)
    {
      "Process"
      {
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Clicked $($Sender.Name)"
        $PILBtmStatusStrip.Refresh()
        
        # Set Processing ToolStrip Menu Items
        $PILItelListToolStrip.Items["Pause"].Checked = $False
        $PILItelListToolStrip.Items["Stop"].Checked = $False
        
        # *****************************************
        # **** Testing - Exit to Nested Prompt ****
        # *****************************************
        Write-Host -Object "Line Num: $((Get-PSCallStack).ScriptLineNumber)"
        $Host.EnterNestedPrompt()
        # *****************************************
        # **** Testing - Exit to Nested Prompt ****
        # *****************************************
        
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Processing Item List"
        Break
      }
      "Pause"
      {
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Clicked $($Sender.Name)"
        $PILBtmStatusStrip.Refresh()
        
        # Set Pauseing ToolStrip Menu Items
        $PILItelListToolStrip.Items["Process"].Checked = $False
        $PILItelListToolStrip.Items["Stop"].Checked = $False
        
        # *****************************************
        # **** Testing - Exit to Nested Prompt ****
        # *****************************************
        Write-Host -Object "Line Num: $((Get-PSCallStack).ScriptLineNumber)"
        $Host.EnterNestedPrompt()
        # *****************************************
        # **** Testing - Exit to Nested Prompt ****
        # *****************************************
        
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Pause Processing Item List"
        Break
      }
      "Stop"
      {
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Clicked $($Sender.Name)"
        $PILBtmStatusStrip.Refresh()
        
        # Set Stopping ToolStrip Menu Items
        $PILItelListToolStrip.Items["Process"].Checked = $False
        $PILItelListToolStrip.Items["Pause"].Checked = $False
        $PILItelListToolStrip.Items["Stop"].Checked = $False
        $PILItelListToolStrip.SendToBack()
        
        # Re-Enable Main Menu Items
        $PILTopMenuStrip.Items["AddItems"].Enabled = $True
        $PILTopMenuStrip.Items["Configure"].Enabled = $True
        $PILTopMenuStrip.Items["ProcessItems"].Enabled = $True
        $PILTopMenuStrip.Items["ListData"].Enabled = $True
        
        # Re-Enable Right Click Menu
        $PILItemListContextMenuStrip.Enabled = $True
        
        # Enable ListView Sort
        $PILItemListListView.ListViewItemSorter.Enable = $True
        
        # *****************************************
        # **** Testing - Exit to Nested Prompt ****
        # *****************************************
        Write-Host -Object "Line Num: $((Get-PSCallStack).ScriptLineNumber)"
        $Host.EnterNestedPrompt()
        # *****************************************
        # **** Testing - Exit to Nested Prompt ****
        # *****************************************
        
        # Set Status Message
        $PILBtmStatusStrip.Items["Status"].Text = "Stop Processing Item List"
        Break
      }
    }
  }

  Write-Verbose -Message "Exit Click Event for $($MyInvocation.MyCommand)"
}
#endregion ******** Function Start-PILItelListToolStripItemClick ********

(New-MenuItem -Menu $PILItelListToolStrip -Text "Process" -Name "Process" -Tag "Process" -ToolTip "Process Item List" -DisplayStyle Image -TextImageRelation Overlay -ImageKey "Play64Icon" -ClickOnCheck -PassThru).add_Click({Start-PILItelListToolStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $PILItelListToolStrip -Text "Pause" -Name "Pause" -Tag "Pause" -ToolTip "Pause Processing" -DisplayStyle Image -TextImageRelation Overlay -ImageKey "Pause64Icon" -ClickOnCheck -PassThru).add_Click({Start-PILItelListToolStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $PILItelListToolStrip -Text "Stop" -Name "Stop" -Tag "Stop" -ToolTip "Stop Processing" -DisplayStyle Image -TextImageRelation Overlay -ImageKey "Stop64Icon" -ClickOnCheck -PassThru).add_Click({Start-PILItelListToolStripItemClick -Sender $This -EventArg $PSItem})

$PILItelListToolStrip.Location = [System.Drawing.Point]::New((($PILMainPanel.ClientSize.Width - $PILItelListToolStrip.Width) / 2), ([MyConfig]::FormSpacer * 16))

#endregion ******** $PILMainPanel Controls ********

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

  Switch ($Sender.Name)
  {
    "AddList"
    {
      #region Add New Items List
      
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Add New Items for Processing"
      $PILBtmStatusStrip.Refresh()
      
      $DialogResult = Get-TextBoxInput -Title "Get Item List" -Message "Enter the list of items to add for processing" -Multi -NoDuplicates
      If ($DialogResult.Success)
      {
        $NewCount = 0
        $TmpSubItems = @("") * [MyRuntime]::MaxColumns
        ForEach ($TmpItem In $DialogResult.Items)
        {
          If (-not $PILItemListListView.Items.ContainsKey($TmpItem))
          {
            $TmpListItem = [System.Windows.Forms.ListViewItem]::New($TmpItem, "StatusInfo16Icon")
            $TmpListItem.Name = $TmpItem
            $TmpListItem.Font = [MyConfig]::Font.Regular
            $TmpListItem.SubItems.AddRange($TmpSubItems)
            [Void]$PILItemListListView.Items.Add($TmpListItem)
            $NewCount++
          }
        }
        
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
            $TmpSubItems = @("") * [MyRuntime]::MaxColumns
            ForEach ($TmpItem In $TmpItems)
            {
              If (-not $PILItemListListView.Items.ContainsKey($TmpItem))
              {
                $TmpListItem = [System.Windows.Forms.ListViewItem]::New($TmpItem, "StatusInfo16Icon")
                $TmpListItem.Name = $TmpItem
                $TmpListItem.Font = [MyConfig]::Font.Regular
                $TmpListItem.SubItems.AddRange($TmpSubItems)
                [Void]$PILItemListListView.Items.Add($TmpListItem)
                $NewCount++
              }
            }
            
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
        $HashTable = @{"ShowHeader" = $True; "ExportFile" = $PILOpenFileDialog.FileName }
        $ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Load-PILDataExport -RichTextBox $RichTextBox -HashTable $HashTable }
        $DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title "Initializing $([MyConfig]::ScriptName)" -ButtonMid "OK" -HashTable $HashTable
        If ($DialogResult.Success)
        {
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
    }
    "TotalColumns"
    {
      #region Set Total Columns
      $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::MaxColumns - 2].ImageKey = $Null
      [MyRuntime]::UpdateTotalColumn($Sender.Tag)
      $TmpColumns = [MyRuntime]::ThreadConfig.GetColumnNames()
      $PILItemListListView.BeginUpdate()
      $PILItemListListView.Columns.Clear()
      For ($I = 0; $I -lt ([MyRuntime]::MaxColumns); $I++)
      {
        New-ColumnHeader -ListView $PILItemListListView -Text ([MyRuntime]::ThreadConfig.ColumnNames[$I]) -Name ([MyRuntime]::ThreadConfig.ColumnNames[$I]) -Tag ([MyRuntime]::ThreadConfig.ColumnNames[$I]) -Width -2
      }
      $PILItemListListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
      New-ColumnHeader -ListView $PILItemListListView -Text " " -Name "Blank" -Tag " " -Width ($PILForm.Width * 4)
      $PILItemListListView.EndUpdate()
      $PILTopMenuStrip.Items["Configure"].DropDownItems["TotalColumns"].DropDownItems[[MyRuntime]::MaxColumns - 2].ImageKey = "Selected16Icon"
      #endregion Set Total Columns
    }
    "ColumnNames"
    {
      #region Set Column Names
      $PILBtmStatusStrip.Items["Status"].Text = "Update Column Names"
      $PILBtmStatusStrip.Refresh()
      $DialogResult = Get-MultiTextBoxInput -Title "Update Column Names" -Message "Enter the New Column Names for the $([MyConfig]::ScriptName) Utility" -OrderedItems ([MyRuntime]::ThreadConfig.GetColumnNames()) -AllRequired
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
    "ThreadScript"
    {
      #region Update Thread Config
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Update PIL Threads Configuration"
      $PILBtmStatusStrip.Refresh()
      $DialogResult = Update-ThreadConfiguration
      If ($DialogResult.Success)
      {
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
      $PILOpenFileDialog.Filter = "PIL Config File|*.PILConfig|All Files (*.*)|*.*"
      $PILOpenFileDialog.FilterIndex = 1
      $PILOpenFileDialog.Multiselect = $False
      $PILOpenFileDialog.Title = "Load PIL Configuration File"
      $PILOpenFileDialog.Tag = $Null
      $Response = $PILOpenFileDialog.ShowDialog()
      If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
      {
        $HashTable = @{"ShowHeader" = $True; "ConfigFile" = $PILOpenFileDialog.FileName }
        $ScriptBlock = { [CmdletBinding()] param ([System.Windows.Forms.RichTextBox]$RichTextBox, [HashTable]$HashTable) Load-PILConfigFIle -RichTextBox $RichTextBox -HashTable $HashTable }
        $DialogResult = Show-RichTextStatus -ScriptBlock $ScriptBlock -Title ($PILBtmStatusStrip.Items["Status"].Text) -ButtonMid "OK" -HashTable $HashTable
        If ($DialogResult.Success)
        {
          $PILBtmStatusStrip.Items["Status"].Text = "Success Loading PIL Configuration File"
        }
        Else
        {
          $PILBtmStatusStrip.Items["Status"].Text = "Errors Loading PIL Configuration File"
        }
        
        # Save Current Directory
        $PILOpenFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILOpenFileDialog.FileName)
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
      $PILSaveFileDialog.FileName = ""
      $PILSaveFileDialog.Filter = "PIL Config File|*.PILConfig|All Files (*.*)|*.*"
      $PILSaveFileDialog.FilterIndex = 1
      $PILSaveFileDialog.Title = "Save PIL Configuration File"
      $PILSaveFileDialog.Tag = $Null
      $Response = $PILSaveFileDialog.ShowDialog()
      If ($Response -eq [System.Windows.Forms.DialogResult]::OK)
      {
        # Save Config
        [MyRuntime]::ThreadConfig | Export-Clixml -Path $PILSaveFileDialog.FileName -Encoding ASCII
        
        # Save Current Directory
        $PILSaveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILSaveFileDialog.FileName)
        $PILBtmStatusStrip.Items["Status"].Text = "Success Saving PIL Configuration File"
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
      # Set Status Message
      $PILBtmStatusStrip.Items["Status"].Text = "Processing Item List"
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
      
      # Set Processing ToolStrip
      $PILItelListToolStrip.Items["Process"].Checked = $True
      $PILItelListToolStrip.Items["Pause"].Checked = $False
      $PILItelListToolStrip.Items["Stop"].Checked = $False
      $PILItelListToolStrip.BringToFront()
      Break
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
          $TmpCount = ([MyRuntime]::MaxColumns - 1)
          $StringBuilder = [System.Text.StringBuilder]::New()
          [Void]$StringBuilder.AppendLine(($PILItemListListView.Columns[0..$($TmpCount)] | Select-Object -ExpandProperty Text) -Join ",")
          $PILItemListListView.Items | ForEach-Object -Process { [Void]$StringBuilder.AppendLine("`"{0}`"" -f (($PSItem.SubItems[0..$($TmpCount)] | Select-Object -ExpandProperty Text) -join "`",`"")) }
          ConvertFrom-Csv -InputObject (($StringBuilder.ToString())) -Delimiter "," | Export-Csv -Path $PILSaveFileDialog.FileName -NoTypeInformation -Encoding ASCII
          $StringBuilder.Clear()
          
          # Save Current Directory
          $PILSaveFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($PILSaveFileDialog.FileName)
          $PILBtmStatusStrip.Items["Status"].Text = "Success Exporting CSV Report"
        }
        Else
        {
          $PILBtmStatusStrip.Items["Status"].Text = "Canceled Exporting CSV Report"
        }
      }
      Break
      #endregion Export CSV Report
    }
    "ClearList"
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
        $DialogResult = Get-UserResponse -Title "Clear Item List?" -Message "Do you want to Clear the Item List?" -ButtonLeft Yes -ButtonRight No -ButtonDefault Yes -Icon ([System.Drawing.SystemIcons]::Question)
        If ($DialogResult.Success)
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
    "Help"
    {
      #region Show Help
      $PILBtmStatusStrip.Items["Status"].Text = "Show Help"
      $PILBtmStatusStrip.Refresh()
      $DialogResult = Show-ScriptInfo -Topics $ScriptInfoTopics -Title "$([MyConfig]::ScriptName) $([MyConfig]::ScriptVersion)" -InfoTitle "PIL Help Topics"
      If ($DialogResult.Success)
      {
        # Success
        $PILBtmStatusStrip.Items["Status"].Text = "Success Help Shown"
      }
      Else
      {
        # Failed
        $PILBtmStatusStrip.Items["Status"].Text = "Failed Help Shown"
      }
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
        $FCGForm.Close()
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
(New-MenuItem -Menu $DropDownMenu -Text "Load Exported Data" -Name "LoadExport" -Tag "LoadExport" -DisplayStyle "ImageAndText" -ImageKey "LoadData16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})

$DropDownMenu = New-MenuItem -Menu $PILTopMenuStrip -Text "Configure $([char]0x00BB)" -Name "Configure" -Tag "Configure" -DisplayStyle "ImageAndText" -ImageKey "Config16Icon" -TextImageRelation "ImageBeforeText" -PassThru
$SubDropDownMenu = New-MenuItem -Menu $DropDownMenu -Text "Number of Columns" -Name "TotalColumns" -Tag "TotalColumns" -DisplayStyle "ImageAndText" -ImageKey "Calc16Icon" -TextImageRelation "ImageBeforeText" -PassThru
For ($I = 2; $I -le [MyRuntime]::MaxColumns; $I++)
{
  (New-MenuItem -Menu $SubDropDownMenu -Text ("{0:00} Total Columns" -f $I) -ToolTip "Set the Number of Item List Columns" -Name "TotalColumns" -Tag $I -DisplayStyle "ImageAndText" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
}
$SubDropDownMenu.DropDownItems[$SubDropDownMenu.DropDownItems.Count - 1].ImageKey = "Selected16Icon"
(New-MenuItem -Menu $DropDownMenu -Text "Set Column Names" -Name "ColumnNames" -Tag "ColumnNames" -DisplayStyle "ImageAndText" -ImageKey "Column16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})

(New-MenuItem -Menu $DropDownMenu -Text "Config Thread Script" -Name "ThreadScript" -Tag "ThreadScript" -DisplayStyle "ImageAndText" -ImageKey "Threads16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $DropDownMenu
(New-MenuItem -Menu $DropDownMenu -Text "Load Configuration" -Name "LoadConfig" -Tag "LoadConfig" -DisplayStyle "ImageAndText" -ImageKey "LoadConfig16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
(New-MenuItem -Menu $DropDownMenu -Text "Save Configuration" -Name "SaveConfig" -Tag "SaveConfig" -DisplayStyle "ImageAndText" -ImageKey "SaveConfig16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $PILTopMenuStrip

(New-MenuItem -Menu $PILTopMenuStrip -Text "Process Items" -Name "ProcessItems" -Tag "ProcessItems" -DisplayStyle "ImageAndText" -ImageKey "Process16Icon" -TextImageRelation "ImageBeforeText" -ClickOnCheck -Check -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})

$DropDownMenu = New-MenuItem -Menu $PILTopMenuStrip -Text "List Data $([char]0x00BB)" -Name "ListData" -Tag "ListData" -DisplayStyle "ImageAndText" -ImageKey "ListData16Icon" -TextImageRelation "ImageBeforeText" -PassThru
(New-MenuItem -Menu $DropDownMenu -Text "Export CSV Report" -Name "ExportCSV" -Tag "ExportCSV" -DisplayStyle "ImageAndText" -ImageKey "Export16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
New-MenuSeparator -Menu $DropDownMenu
(New-MenuItem -Menu $DropDownMenu -Text "Clear Item List Data" -Name "ClearList" -Tag "ClearList" -DisplayStyle "ImageAndText" -ImageKey "Clear16Icon" -TextImageRelation "ImageBeforeText" -PassThru).add_Click({Start-PILTopMenuStripItemClick -Sender $This -EventArg $PSItem})
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

#endregion ******** End **** PIL **** End ********

#region ******** Start Form  ********
# *********************
# Add Form Code here...
# *********************
[System.Console]::Title = "RUNNING: $([MyConfig]::ScriptName) - $([MyConfig]::ScriptVersion)"
if ([MyConfig]::Production)
{
  [Void][Console.Window]::Hide()
}

Try
{
  [System.Windows.Forms.Application]::Run($PILForm)
}
Catch
{
  if (-not [MyConfig]::Production)
  {
    # **** Testing - Exit to Nested Prompt ****
    Write-Host -Object "Line Num: $((Get-PSCallStack).ScriptLineNumber)"
    #$Host.EnterNestedPrompt()
    # **** Testing - Exit to Nested Prompt ****
  }
}

$PILOpenFileDialog.Dispose()
$PILSaveFileDialog.Dispose()
$PILFormComponents.Dispose()
$PILForm.Dispose()
# *********************
# Add Form Code here...
# *********************

#endregion ******** Start Form  ********

if ([MyConfig]::Production)
{
  [System.Environment]::Exit(0)
}
