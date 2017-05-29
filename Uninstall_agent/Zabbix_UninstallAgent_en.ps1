[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [array]$computerName,
    [string]$ZabbixPath = "C:\Users\tpoccard\Documents\Projet\Zabbix\Uninstall_agent",
    [string]$ZabbixVersion = "3.2"
)


$ZabbixService = "Zabbix Agent"
$ZabbixSource = "$ZabbixPath"
$loginRemoteMachine = "siriona\tpoccard"
$uninstOk = ""
$uninstBad = ""
$notfound = ""
$path_zabbix_exe = "C:\Zabbix\zabbix_agentd.exe"
$path_zabbix_conf = "C:\Zabbix\zabbix_agentd.win.conf"
#$computername = 'siriona-dc1'
$domaincredentials = "siriona\tpoccard"
$destinationFolder = "C:\Zabbix"
$sName = "Zabbix Agent"
$service = ""

#$cred = Get-Credential -username $loginRemoteMachine -Message "Enter the password"

function addComputerToString($computers, $newComputer) 
{
    if(-not ($computers -match $newComputer)) 
    {
        return "$computers $newComputer"
    }
    else 
    {
        return $computers
    }
}


Write-Host "================================================================================"
Write-Host " Welcome!" 
Write-Host " This script requires administrative privileges on all implicated servers!"  
Write-Host "================================================================================"

# ===========================================================================================
# Checking Administrative privilege
# ===========================================================================================

if (!(Test-Path "C:\Users\tpoccard\Documents\Projet\Zabbix\Install_agent\Test-UserCredentials.ps1"))
{
    Write-Host " "
    Write-Host "#########################################################"
    Write-Host "Error! - Script checking AD authentication not found !" -foregroundcolor "red"
    Write-Host "#########################################################"
    Write-Host " "
    exit
}

else 
{
    . C:\Users\tpoccard\Documents\Projet\Zabbix\Install_agent\Test-UserCredentials.ps1
    $result = TestUserCredentials
    ClearUserInfo
  

    if ($result -ne $Null)
    {
        <#
        # ===========================================================================================
        # Verifying the Zabbix folder
        # ===========================================================================================
        Write-Host "================================================================================"
        Write-Host " Zabbix folder:" $ZabbixSource 
    
        if (!(Test-Path $ZabbixSource))
        {
        Write-Host " "
        Write-Host "##############################################"
        Write-Host "# Error! - Zabbix folder does not exist!     #"
        Write-Host "##############################################"
        Write-Host " "
        exit
        }
 
        Write-Host "================================================================================"
        #>

        # ===========================================================================================
        # Execute actions for each server in list
        # ===========================================================================================

        foreach ($computer in $computername) 
        {     

            


                $ZabbixServiceState = $false
                If ((Test-Connection $computer -quiet -count 1))
                {
                    
                    #Write-Host "================================================================================"
                    Write-Host " $computer is available" -foregroundcolor "green"
                    Write-Host "================================================================================"
                     
                    $session = New-PSSession -ComputerName $computer -Credential $result

                    
                        $FolderState =  Invoke-Command -ArgumentList $destinationFolder -Session $session  -ScriptBlock {
                            param($destinationFolder) 

                            if(Test-Path $destinationFolder)
                            {
                                return $true                              
                            }
                            else { return $false }
                            
                        }

                
                    if ((Get-Service -ComputerName $computer -Name $ZabbixService -ErrorAction SilentlyContinue) -and ($FolderState -eq $true))
                    {
                        # Zabbix Agent is present!
                        # Bool. Get service state!

                      



                        $sState =  Invoke-Command -ArgumentList $ZabbixSource,$sName -Session $session  -ScriptBlock {
                            param($ZabbixSource,$sName) 
                            $sName = "Zabbix Agent"
                            $service = Get-Service -name $sName -ErrorAction SilentlyContinue 
                            return $service                              
                        }

                        if ($sState.Status -eq "Stopped") 
                        {
                            Write-Host "$sName is already stopped on $computer"

                            try
                            {
                                Write-Host "Uninstallation process on $computer ..."
                                $exec = "c:\Zabbix\zabbix_agentd.exe -c c:\Zabbix\zabbix_agentd.win.conf -d"
                                $remoteWMI = Invoke-WMIMethod -Class Win32_Process -Name Create -Computername $computer -ArgumentList $exec
                                Start-Sleep -Second 2 

                                if ($remoteWMI.ReturnValue -ne 0)
                                {
                                    # Oops...
                                    Write-Host "================================================================================"
                                    Write-Host " Problem while uninstalling previous zabbix agent! Cancelling..." -foregroundcolor "red"
                                    Write-Host " Error:" $remoteWMI.ReturnValue -foregroundcolor "red"
                                    Write-Host " 0 Successful Completion" -foregroundcolor "red"
                                    Write-Host " 3 Insufficient Privilege" -foregroundcolor "red"
                                    Write-Host " 8 Unknown Failure" -foregroundcolor "red"
                                    Write-Host " 9 Path Not Found" -foregroundcolor "red"
                                    Write-Host " 21 Invalid Parameter" -foregroundcolor "red"
                                    Write-Host "================================================================================"
                                    $uninstBad = addComputerToString $uninstBad $computer
                                    continue                         
                                }

                                Invoke-Command -ArgumentList $ZabbixSource -Session $session -ScriptBlock {
                                    param($ZabbixSource)

                                    Set-Location "C:\"
                                    Start-Sleep -Second 1
                                    Remove-Item  "C:\Zabbix" -recurse 
                                    Start-Sleep -Second 1
                                }
                                $uninstOk = addComputerToString $uninstOk $computer
                            }

                            catch
                            {
                                Write-Host "================================================================================"
                                Write-Host " Problem while uninstalling previous zabbix agent! Cancelling..." -foregroundcolor "red"
                                Write-Host $_ -foregroundcolor "red"
                                Write-Host "================================================================================"
                                $uninstBad = addComputerToString $uninstBad $computer
                                continue    
                            }
                        }


                        if ($sState.status -eq "Running") 
                        {
                            try
                            {
                                Write-Host " Stopping Zabbix Agent..."
                                Write-Host "================================================================================"
                                Set-Service -ComputerName $computer -Status Stopped -Name $ZabbixService -ErrorAction Stop
                                Start-Sleep -Seconds 1   

                                Write-Host "Uninstallation process on $computer ..."
                                $exec = "c:\Zabbix\zabbix_agentd.exe -c c:\Zabbix\zabbix_agentd.win.conf -d"

                                $remoteWMI = Invoke-WMIMethod -Class Win32_Process -Name Create -Computername $computer -ArgumentList $exec
                                Start-Sleep -Second 3
                                

                                if ($remoteWMI.ReturnValue -ne 0)
                                {
                                    # Oops...
                                    Write-Host "================================================================================"
                                    Write-Host " Problem while uninstalling previous zabbix agent! Cancelling..." -foregroundcolor "red"
                                    Write-Host " Error:" $remoteWMI.ReturnValue -foregroundcolor "red"
                                    Write-Host " 0 Successful Completion" -foregroundcolor "red"
                                    Write-Host " 3 Insufficient Privilege" -foregroundcolor "red"
                                    Write-Host " 8 Unknown Failure" -foregroundcolor "red"
                                    Write-Host " 9 Path Not Found" -foregroundcolor "red"
                                    Write-Host " 21 Invalid Parameter" -foregroundcolor "red"
                                    Write-Host "================================================================================"
                                    $uninstBad = addComputerToString $uninstBad $computer
                                    continue                         
                                }

                                Invoke-Command -ArgumentList $ZabbixSource -Session $session -ScriptBlock {
                                    param($ZabbixSource)
                                    
                                    Set-Location "C:\"
                                    Start-Sleep -Second 1
                                    Remove-Item  "C:\Zabbix" -recurse 
                                    Start-Sleep -Second 1
                                }

                                $uninstOk = addComputerToString $uninstOk $computer
                            }

                            catch 
                            {
                                Write-Host "================================================================================"
                                Write-Host " Problem while uninstalling previous zabbix agent! Cancelling..."
                                Write-Host $_
                                Write-Host "================================================================================"
                                $uninstBad = addComputerToString $uninstBad $computer
                                continue
                            }

                        }
                        
                        else 
                        {
                            Write-Host "$sName could not be uninstalled on $computer"
                            $uninstBad = addComputerToString $uninstBad $computer
                            continue
                        }  
                    }

                    else
                    {
                        #Write-Host " "
	                     #   Write-Host "################################################"
	                        Write-Host " Error! - Zabbix Agent is not installed!" -foregroundcolor "red"
	                      #  Write-Host "################################################"
	                       # Write-Host " "
                            $uninstBad = addComputerToString $uninstBad $computer
	                        continue
                    }
                }

                else
                {
                    Write-Host " $computer is unavailable"
                    $uninstBad = addComputerToString $uninstBad $computer
                    continue
                }
      
        } 
        }
                

            
    if ($session -ne $null) {
            Remove-PSSession -Session $session  # remove active session
    }


    Write-Host "================================================================================"
    Write-Host " Successful uninstallations: " $uninstOk -foregroundcolor "green"
    Write-Host " Unsuccessful uninstallations: " $uninstBad -foregroundcolor "red"
    Write-Host "================================================================================"
    Write-Host " "
    Read-Host " Press any key to finish!"
    exit 0
}
