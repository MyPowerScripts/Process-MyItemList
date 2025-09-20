
#region **** Function Get-MyWorkstationInfo ****
Function Get-MyWorkstationInfo ()
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
  Param (
    [parameter(Mandatory = $False, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
    [String[]]$ComputerName = [System.Environment]::MachineName,
    [PSCredential]$Credential,
    [Switch]$Serial,
    [Switch]$Mobile
  )
  Begin
  {
    Write-Verbose -Message "Enter Function Get-MyWorkstationInfo"
    
    # Default Common Get-WmiObject Options
    If ($PSBoundParameters.ContainsKey("Credential"))
    {
      $Params = @{
        "ComputerName" = $Null
        "Credential"   = $Credential
      }
    }
    Else
    {
      $Params = @{
        "ComputerName" = $Null
      }
    }
  }
  Process
  {
    Write-Verbose -Message "Enter Function Get-MyWorkstationInfo - Process"
    
    ForEach ($Computer In $ComputerName)
    {
      # Start Setting Return Values as they are Found
      $VerifyObject = [MyWorkstationInfo]::New($Computer)
      
      # Validate ComputerName
      If (($Computer -match "^(([a-zA-Z]|[a-zA-Z][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z]|[A-Za-z][A-Za-z0-9\-]*[A-Za-z0-9])$") -or ($Computer -match "(?:25[0-5]|2[0-4][0-9]|1\d{2}|[1-9]?\d)(?:\.(?:25[0-5]|2[0-4][0-9]|1\d{2}|[1-9]?\d)){3}"))
      {
        Try
        {
          # Get IP Address from DNS, you want to do all remote checks using IP rather than ComputerName.  If you connect to a computer using the wrong name Get-WmiObject will fail and using the IP Address will not
          $IPAddresses = @([System.Net.Dns]::GetHostAddresses($Computer) | Where-Object -FilterScript {
              $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork
            } | Select-Object -ExpandProperty IPAddressToString)
          :FoundMyWork ForEach ($IPAddress In $IPAddresses)
          {
            If ([System.Net.NetworkInformation.Ping]::New().Send($IPAddress).Status -eq [System.Net.NetworkInformation.IPStatus]::Success)
            {
              # Set Default Parms
              $Params.ComputerName = $IPAddress
              
              # Get ComputerSystem
              [Void]($MyCompData = Get-WmiObject @Params -Class Win32_ComputerSystem)
              $VerifyObject.AddComputerSystem($Computer, $IPAddress, ($MyCompData.Name), ($MyCompData.PartOfDomain), ($MyCompData.Domain), ($MyCompData.Manufacturer), ($MyCompData.Model), ($MyCompData.UserName), ($MyCompData.TotalPhysicalMemory))
              $MyCompData.Dispose()
              
              # Verify Remote Computer is the Connect Computer, No need to get any more information
              If ($VerifyObject.Found)
              {
                # Start Secondary Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
                [Void]($MyOSData = Get-WmiObject @Params -Class Win32_OperatingSystem)
                $VerifyObject.AddOperatingSystem(($MyOSData.ProductType), ($MyOSData.Caption), ($MyOSData.CSDVersion), ($MyOSData.BuildNumber), ($MyOSData.Version), ($MyOSData.OSArchitecture), ([System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.LocalDateTime)), ([System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.InstallDate)), ([System.Management.ManagementDateTimeConverter]::ToDateTime($MyOSData.LastBootUpTime)))
                $MyOSData.Dispose()
                
                # Optional SerialNumber Job
                If ($Serial.IsPresent)
                {
                  # Start Optional Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
                  [Void]($MyBIOSData = Get-WmiObject @Params -Class Win32_Bios)
                  $VerifyObject.AddSerialNumber($MyBIOSData.SerialNumber)
                  $MyBIOSData.Dispose()
                }
                
                # Optional Mobile / ChassisType Job
                If ($Mobile.IsPresent)
                {
                  # Start Optional Job, Pass IP Address and Credentials to Job Script to make Connection to Remote Computer
                  [Void]($MyChassisData = Get-WmiObject @Params -Class Win32_SystemEnclosure)
                  $VerifyObject.AddIsMobile($MyChassisData.ChassisTypes)
                  $MyChassisData.Dispose()
                }
              }
              Else
              {
                $VerifyObject.UpdateStatus("Wrong Workstation Name")
              }
              # Beak out of Loop, Verify was a Success no need to try other IP Address if any
              Break FoundMyWork
            }
          }
        }
        Catch
        {
          # Workstation Not in DNS
          $VerifyObject.UpdateStatus("Workstation Not in DNS")
        }
      }
      Else
      {
        $VerifyObject.UpdateStatus("Invalid Computer Name")
      }
      
      # Set End Time and Return Results
      $VerifyObject.SetEndTime()
    }
    Write-Verbose -Message "Exit Function Get-MyWorkstationInfo - Process"
  }
  End
  {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Verbose -Message "Exit Function Get-MyWorkstationInfo"
  }
}
#endregion **** Function Get-MyWorkstationInfo ****

