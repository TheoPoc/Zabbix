# ===========================================================================================
# 
# NAME: Zabbix_InstallAgent.ps1
# 
# AUTHOR: Pierre-Emmanuel Turcotte, 
# DATE  : 2012-11-05
# 
# COMMENT:	Remote installation of Zabbix agent from central location.
#
#			The zabbix source folder is a folder containing subfolder corresponding to the 
#			different agent versions. These subfolders contain the win32 and win64 subfolders
#			along with the zabbix_agentd.win.conf configuration file.
#
#			You must have administrative priviledges on the remote computers you will specify
#			in the computerName parameter. You can either use the following format:
#
#			zabbix_installagent.ps1 -ComputerName server1,server2,server3
#
#			or do not specify any servers, you will be prompted for servers, one at a time.
#
#			Additionnally, you have the following parameters you can optionally define:
#
#			-ZabbixVersion 		#Version of the agent to be installed(points to subfolder)
#								 Defaults to "2.0.3"
#			-ZabbixPath			#Folder where the agent installation resides on the central
#								 server. This is the folder that will be copied over to the
#								 remote locations.
#								 Defaults to "O:\Zabbix"
#
#			About the zabbix path, on my installation, this folder holds 3 subfolders:
#
#			O:\Zabbix
#					 \1.8.10
#					 \2.0
#					 \2.0.3
#
#			$ZabbixPath and $ZabbixVersion will be assembled in a new variable named
#			$ZabbixSource, look for it for a better undestanding.
#
# DISCLAIMER
#
# The sample script provided here are not supported by Pierre-Emmanuel Turcotte or his 
# employer. All scripts are provided AS IS without warranty of any kind. Pierre-Emmanuel
# Turcotte and his employer further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Pierre-Emmanuel Turcotte or his 
# employer, its authors, or anyone else involved in the creation, production, or delivery
# of the scripts be liable for any damages whatsoever (including, without limitation,
# damages for loss of business profits, business interruption, loss of business information,
# or other pecuniary loss) arising out of the use of or inability to use the sample scripts
# or documentation, even if Pierre-Emmanuel Turcotte or his employer has been advised of the
# possibility of such damages.
# 
# ===========================================================================================

# ===========================================================================================
# Parameters
# ===========================================================================================
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True,Position=1)]
	[array]$computerName,
	[string]$ZabbixPath = "C:\Zabbix",
	[string]$ZabbixVersion = "3.2"
)

# ===========================================================================================
# Initialization
# ===========================================================================================
$ZabbixService = "Zabbix Agent"
$ZabbixSource = "$ZabbixPath"
$loginRemoteMachine = "siriona\tpoccard"

$instOk = ""
$instBad = ""

$path_zabbix_exe = "C:\Zabbix\zabbix_agentd.exe"
$path_zabbix_conf = "C:\Zabbix\zabbix_agentd.win.conf"

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
	Write-Host "##############################################"
	Write-Host "# Error! - Installation folder does not exist!	   #"
	Write-Host "##############################################"
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

        # ===========================================================================================
        # Verifying the source installation folder
        # ===========================================================================================
        Write-Host "================================================================================"
        Write-Host " Zabbix source installation folder:" $ZabbixSource 

        if (!(Test-Path $ZabbixSource))
        {
	        Write-Host " "
	        Write-Host "##############################################"
	        Write-Host "# Error! - Installation folder does not exist!	   #"
	        Write-Host "##############################################"
	        Write-Host " "
	        exit
        }
        Write-Host "================================================================================"


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
		

			        $session = New-PSSession -ComputerName $computer -Credential $cred

   			  
			        Invoke-Command -ArgumentList $ZabbixSource -Session $session  -ScriptBlock {
			            param($ZabbixSource)

		                if (!(Test-Path -path $ZabbixSource)) {
		                    New-Item $ZabbixSource -Type Directory
		                }
		            }


			        Copy-Item -Path $path_zabbix_exe -Destination $ZabbixSource -ToSession $session
			        Copy-Item -Path $path_zabbix_conf -Destination $ZabbixSource -ToSession $session
			        Write-Host " Copying folder..."
		        }
		        catch
		        {
			        Write-Host "================================================================================"
			        Write-Host " $computer - Zabbix agent could not be installed!"g
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
			        $exec = "c:\Zabbix\zabbix_agentd.exe -c c:\Zabbix\zabbix_agentd.win.conf -d"
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
			        $exec = "c:\Zabbix\zabbix_agentd.exe -c c:\Zabbix\zabbix_agentd.win.conf -i"
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
			        $exec = "c:\Zabbix\zabbix_agentd.exe -c c:\Zabbix\zabbix_agentd.win.conf -s"
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
		        }
		        else
		        {
			        Write-Host "================================================================================"
			        Write-Host " Service installed but not started!"
			        Write-Host " Service state: " $InstallStatus.Status
			        Write-Host "================================================================================"
		        }
	
                $instOk = addComputerToString $instOk $computer
                Set-Service -Name "Zabbix Agent" -StartupType Automatic
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







