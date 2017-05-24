
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True,Position=1)]
	[array]$computerName,
	[string]$ZabbixPath = "http://192.168.3.100/scripts",
	[string]$ZabbixVersion = "3.2"
)
http://192.168.3.100/Scripts/Install_agent/
# ===========================================================================================
# Initialization
# ===========================================================================================
$ZabbixService = "Zabbix Agent"
$ZabbixSource = "$ZabbixPath"
$ZabbixDestination = "C:\Zabbix"
$loginRemoteMachine = "siriona\tpoccard"

$instOk = ""
$instBad = ""

$path_zabbix_exe = "$ZabbixPath\zabbix_agentd.exe"
$path_zabbix_conf = "$ZabbixPath\zabbix_agentd.win.conf"

# ===========================================================================================
# Test User Credential
# ===========================================================================================


# ===========================================================================================
# Functions declarations
# ===========================================================================================

function addComputerToString($computers, $newComputer) {
    if(-not ($computers -match $newComputer)) {
        return "$computers $newComputer"
    }else {
        return $computers
    }
}



# ===========================================================================================
# Welcome message
# ===========================================================================================
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
	Write-Host "##############################################################"
	Write-Host "# Error! - Script Checking AD Authentication does not exist! #"
	Write-Host "##############################################################"
	Write-Host " "
	exit
}


if (!(Test-Path $ZabbixSource))
{
	Write-Host " "
	Write-Host "################################################"
	Write-Host "# Error! - Installation folder does not exist! #"
	Write-Host "################################################"
	Write-Host " "
	exit
}

#Write-Host "================================================================================"

else 
{
    . C:\Users\tpoccard\Documents\Projet\Zabbix\Install_agent\Test-UserCredentials.ps1
    $result = TestUserCredentials
   
    #ClearUserInfo

    if ($result -ne $Null)
    {

        # ===========================================================================================
        # Verifying the source installation folder
        # ===========================================================================================
        Write-Host "================================================================================"
        Write-Host " Zabbix source installation folder:" $ZabbixSource 

        


        # ===========================================================================================
        # Execute actions for each server in list
        # ===========================================================================================
        foreach($computer in $computerName)
        {

	        #$zabbixDestination = "\\$computer\Zabbix"
	        $ZabbixServiceState = $false
	        If ((Test-Connection $computer -quiet -count 1))
	        {
		        # Server pings!
		        Write-Host "================================================================================"
		        Write-Host $computer "is available"
		        Write-Host "================================================================================"
		        # Verify if service is present...
		        # Negative response may be caused by the absence of service (good!) or lack of priviledges (bad!)
		        if (Get-Service -ComputerName $computer -Name $ZabbixService -ErrorAction SilentlyContinue)
		        {
			        # Zabbix Agent is present!
			        # Bool. Get service state!
			        $ZabbixServiceState = ((Get-Service -ComputerName $computer -Name $ZabbixService).status -ne "Stopped")	
			
			        # Verify service state! Stop if necessary!
			        if ($ZabbixServiceState)
			        {
				        # Zabbix agent present and started!
				        try
				        {
					        Write-Host "================================================================================"
					        Write-Host " Stopping Zabbix Agent..."
					        Write-Host "================================================================================"
					        Set-Service -ComputerName $computer -Status Stopped -Name $ZabbixService -ErrorAction Stop
					        Start-Sleep -Seconds 1
				        }
				        catch
				        {
					        Write-Host "================================================================================"
					        Write-Host " $computer - Zabbix agent could not be installed!"
					        Write-Host " Service operation error!" $_			
					        Write-Host "================================================================================"

                    
                            $instBad = addComputerToString $instBad $computer
                            continue
																			        # Next server!
					
				        }
				        Write-Host "================================================================================"
				        Write-Host " Zabbix service stopped."
				        Write-Host " Install new agent..."
			        }
			        else
			        {
				        Write-Host "================================================================================"
				        Write-Host " Service $ZabbixService on $computer is stopped."	# Ok!
				        Write-Host " Install new agent..." 
			        }	
		        }
		        else
		        {
			        Write-Host "================================================================================"
			        Write-Host " Service $ZabbixService not found."				# Ok!
			        Write-Host " Install Zabbix agent..."
		        }
		        # Copy over zabbix folder, make sure no FW could block this
		        try
		        {
			        # If folder exists, it will be overwritten. 
			        # At this point, service should've been stopped...
		

			        $session = New-PSSession -ComputerName $computer -Credential $result

   			  
			        Invoke-Command -ArgumentList $ZabbixDestination -Session $session  -ScriptBlock {
			            param($ZabbixDestination)
		                if (!(Test-Path -path "$ZabbixDestination" -PathType Container)) {
		                    New-Item "$ZabbixDestination" -Type Directory
		                }
		            }


			        Copy-Item -Path $path_zabbix_exe -Destination $ZabbixDestination -ToSession $session
			        Copy-Item -Path $path_zabbix_conf -Destination $ZabbixDestination -ToSession $session
			        Write-Host " Copying folder..."
		        }
		        catch
		        {
			        Write-Host "================================================================================"
			        Write-Host " $computer - Zabbix agent could not be installed!"
			        Write-Host " Make sure you have administrative rights on target!"
			        Write-Host " Copy error!" $_
			        Write-Host "================================================================================"
			        $instBad = addComputerToString $instBad $computer
                    continue
			
			
																		        # Next Server
		        }
		        Write-Host " Folder successfully copied!"									# No catch triggered, so far so good...
		
		        #################################################################################################################################
		        # At this point, the service state has been determined and the folder has been copied to, rights should not be an issue...
		        # Let's install...
		        #################################################################################################################################
		
		        ### Get architecture x86 or x64...
		        try
		        {
			        $os= Get-WMIObject -Class win32_operatingsystem -ComputerName $computer -ErrorAction Stop
		        }
		        catch
		        {
			        Write-Host "================================================================================"
			        Write-Host " $computer - Zabbix agent could not be installed!"
			        Write-Host " -- Make sure you have administrative rights!"
			        Write-Host " -- Make sure no firewall is blocking communications!"
			        Write-Host " Error!" $_
			        Write-Host "================================================================================"
			        $instBad = addComputerToString $instBad $computer
                    continue
			
			
		        }
		        if($os.OSArchitecture -ne $null)
		        {
			        # Architecture can be determined by $os.OSArchitecture...
			        if ($os.OSArchitecture -eq "64-bit")
			        {
				        Write-Host " 64bit system detected!"
				        $osArch = "win64"
			        }
			        elseif($os.OSArchitecture -eq "32-bit")
			        {
				        Write-Host " 32bit system detected!"
				        $osArch = "win32"
			        }
			        else
			        {
				        Write-Host "================================================================================"
				        Write-Host " Unknown architecture! Operation Canceled..."
				        Write-Host $osArch
				        Write-Host "================================================================================"
				        $instBad = addComputerToString $instBad $computer
                        continue
				
				
			        }
		        }
		        else
		        {
			        Write-Host " Windows Pre-2008"
			        # Here have to analyze $os.Caption to determine architecture...
			        if($os.Caption  -match "x64")
			        {
				        Write-Host " 64bit system detected!"
				        $osArch = "win64"
			        }
			        else
			        {
				        Write-Host " 32bit system detected!"
				        $osArch = "win32"
			        }
		        }
		        ### Architecture detection ended.
		        ### Begin installation...
		        try
		        {
			        # Create uninstall string
			        Write-Host " Create uninstall string..."
			        $exec = "$ZabbixDestination\zabbix_agentd.exe -c $ZabbixDestination\zabbix_agentd.win.conf -d"
			        # Execute uninstall string
			        Write-Host " Execute uninstall string..."
			        $remoteWMI = Invoke-WMIMethod -Class Win32_Process -Name Create -Computername $computer -ArgumentList $exec
			        Start-Sleep -Second 2 
			        if ($remoteWMI.ReturnValue -ne 0)
			        {
				        # Oops...
				        Write-Host "================================================================================"
				        Write-Host " Problem while uninstalling previous zabbix agent! Cancelling..."
				        Write-Host " Error:" $remoteWMI.ReturnValue
				        Write-Host " 0 Successful Completion"
				        Write-Host " 3 Insufficient Privilege"
				        Write-Host " 8 Unknown Failure"
				        Write-Host " 9 Path Not Found"
				        Write-Host " 21 Invalid Parameter"
				        Write-Host "================================================================================"
				        $instBad = addComputerToString $instBad $computer
                        continue
				
				
			        }
		        }
		        catch
		        {
			        Write-Host "================================================================================"
			        Write-Host " Problem while uninstalling previous zabbix agent! Cancelling..."
			        Write-Host $_
			        Write-Host "================================================================================"
			        $instBad = addComputerToString $instBad $computer
			        continue
			

			
		        }
		        try
		        {
			        # Create install string
			        Write-Host " Create install string..."
			        $exec = "$ZabbixDestination\zabbix_agentd.exe -c $ZabbixDestination\zabbix_agentd.win.conf -i"
			        # Execute install string
			        Write-Host " Execute install string..."
			        $remoteWMI = Invoke-WMIMethod -Class Win32_Process -Name Create -Computername $computer -ArgumentList $exec
			        Start-Sleep -Second 3
			        if ($remoteWMI.ReturnValue -ne 0)
			        {
				        # Oops...
				        Write-Host "================================================================================"
				        Write-Host " Problem while installing new agent! Cancelling..."
				        Write-Host " Error: " $remoteWMI.ReturnedValue
				        Write-Host " 0 Successful Completion"
				        Write-Host " 3 Insufficient Privilege"
				        Write-Host " 8 Unknown Failure"
				        Write-Host " 9 Path Not Found"
				        Write-Host " 21 Invalid Parameter"
				        Write-Host "================================================================================"
				        $instBad = addComputerToString $instBad $computer
				        continue
				
				
			        }
		        }
		        catch
		        {
			        Write-Host "================================================================================"
			        Write-Host " Problem while installing new agent! Cancelling..."
			        Write-Host $_
			        Write-Host "================================================================================"
			        $instBad = addComputerToString $instBad $computer
			        continue
			
		        }
		        try
		        {
			        # Create run string
			        Write-Host " Create run string..."
			        $exec = "$ZabbixDestination\zabbix_agentd.exe -c $ZabbixDestination\zabbix_agentd.win.conf -s"
			        # Execute run string
			        Write-Host " Execute run string..."
			        $remoteWMI = Invoke-WMIMethod -Class Win32_Process -Name Create -Computername $computer -ArgumentList $exec
			        Start-Sleep -Second 3
			        if ($remoteWMI.ReturnValue -ne 0)
			        {
				        # Problème...
				        Write-Host "================================================================================"
				        Write-Host " Problem while starting the agent! Cancelling..."
				        Write-Host " Error: " $remoteWMI.ReturnedValue
				        Write-Host " 0 Successful Completion"
				        Write-Host " 3 Insufficient Privilege"
				        Write-Host " 8 Unknown Failure"
				        Write-Host " 9 Path Not Found"
				        Write-Host " 21 Invalid Parameter"
				        Write-Host "================================================================================"
				        $instBad = addComputerToString $instBad $computer
				        continue
				
				
			        }
		        }
		        catch
		        {
			        Write-Host "================================================================================"
			        Write-Host " Problem while starting the agent! Cancelling..."
			        Write-Host $_
			        Write-Host "================================================================================"
			        $instBad = addComputerToString $instBad $computer
			        continue
			
			
		        }
		        ### Installation end...
		        ### Start verification
		        try
		        {
			        $InstallStatus = Get-Service -ComputerName $computer -Name $ZabbixService
                    
		        }
		        catch
		        {
			        Write-Host "================================================================================"
			        Write-Host " Problem while verifying service!"
			        Write-Host " Error! " $_
			        Write-Host "================================================================================"
			        $instBad = addComputerToString $instBad $computer
			        continue
			
			
		        }
		        if ($InstallStatus.Status -eq "Running")
		        {
			        Write-Host "================================================================================"
			        Write-Host " Service installed and started!"
			        Write-Host "================================================================================"



					Invoke-Command -ArgumentList $ZabbixDestination -Session $session  -ScriptBlock {
								            param($ZabbixDestination)

						Set-Service -Name "Zabbix Agent" -StartupType Automatic
				}
					
			        
		        }
		        else
		        {
			        Write-Host "================================================================================"
			        Write-Host " Service installed but not started!"
			        Write-Host " Service state: " $InstallStatus.Status
			        Write-Host "================================================================================"
		        }
	
                $instOk = addComputerToString $instOk $computer
                
	        }

	        Else
	        {
		        Write-Host $computer "is not available! Skipping!"
		        $instBad = addComputerToString $instBad $computer
		        continue
																        # Next server
	        }
        }
        Write-Host "================================================================================"
        Write-Host " SCRIPT FINISHED!"
        Write-Host " Successful installations: " $instOk -foregroundcolor "green"
        Write-Host " Unsuccessful installations: " $instBad -foregroundcolor "red"
        Write-Host "================================================================================"
        Write-Host " "
        Read-Host " Press any key to finish!"
        exit 0
            }

}







