
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
      if (($Computer -match "^(([a-zA-Z]|[a-zA-Z][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z]|[A-Za-z][A-Za-z0-9\-]*[A-Za-z0-9])$") -or ($Computer -match "(?:25[0-5]|2[0-4][0-9]|1\d{2}|[1-9]?\d)(?:\.(?:25[0-5]|2[0-4][0-9]|1\d{2}|[1-9]?\d)){3}"))
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
