# ===========================================================================================
# 
# NAME: Zabbix_InstallAgent_en.ps1
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
	[string]$ZabbixPath = "http://192.168.3.100/scripts",
	[string]$ZabbixVersion = "3.2"
)

# ===========================================================================================
# Initialization
# ===========================================================================================
$Path_credential_script = "C:\Users\tpoccard\Documents\Projet\Zabbix\Install_agent\Test-UserCredentials.ps1"
$ZabbixServiceName = "Zabbix Agent"
$ZabbixDirectoryInstall = "C:\Zabbix"
$loginRemoteMachine = "siriona\tpoccard"
$instOk = ""
$instBad = ""
$path_zabbix_exe = "$ZabbixPath\zabbix_agentd.exe"
$path_zabbix_conf = "$ZabbixPath\zabbix_agentd.win.conf"
$launch_install = $false

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

function action_agent($param_action_service){
$exec = "$ZabbixDirectoryInstall\zabbix_agentd.exe -c $ZabbixDirectoryInstall\zabbix_agentd.win.conf $param_action_service"

 

$remoteWMI = Invoke-WMIMethod -Class Win32_Process -Name Create -Computername $computer -ArgumentList $exec
Start-Sleep -Second 1 
if ($remoteWMI.ReturnValue -ne 0)
{
    # Oops...
    Write-Host "################################################################################"
    Write-Host " Problem while uninstalling previous zabbix agent! Cancelling..."
    Write-Host " Error:" $remoteWMI.ReturnValue
    Write-Host " 0 Successful Completion"
    Write-Host " 3 Insufficient Privilege"
    Write-Host " 8 Unknown Failure"
    Write-Host " 9 Path Not Found"
    Write-Host " 21 Invalid Parameter"
    Write-Host "################################################################################"
    $instBad = addComputerToString $instBad $computer
    continue
}	
}


# ===========================================================================================
# Welcome message
# ===========================================================================================
Write-Host "################################################################################"
Write-Host "#                                  Welcome!									   #" 
Write-Host "#  This script requires administrative privileges on all implicated servers!   #"  
Write-Host "################################################################################"

# ===========================================================================================
# Checking Administrative privilege
# ===========================================================================================



if (!(Test-Path $Path_credential_script))
{
	Write-Host " "
	Write-Host "##############################################################"
	Write-Host "# Error! - Script Checking AD Authentication does not exist! #"
	Write-Host "##############################################################"
	Write-Host " "
	exit
}

    . C:\Users\tpoccard\Documents\Projet\Zabbix\Install_agent\Test-UserCredentials.ps1
    $result = TestUserCredentials
  
    if ($result -ne $Null)
    {
        # ===========================================================================================
        # Verifying the source installation folder
        # ===========================================================================================
        Write-Host "##############################################################"
        Write-Host " Zabbix source installation folder:" $ZabbixPath
      

        # ===========================================================================================
        # Execute actions for each server in list
        # ===========================================================================================
        foreach($computer in $computerName)
        {

	        $ZabbixServiceNameState = $false
	        $launch_install = $false  					# Installation unauthorized

	        If ((Test-Connection $computer -quiet -count 1))
	        {

		        # Server pings!
		        Write-Host "################################################################################"
		        Write-Host " $computer is available" -foregroundcolor "green"
		        Write-Host "################################################################################"


                if (Get-Service -ComputerName $computer -Name $ZabbixServiceName -ErrorAction SilentlyContinue)
		        {
			        # Zabbix Agent is present!
			        # Bool. Get service state!
			        $ZabbixServiceNameState = ((Get-Service -ComputerName $computer -Name $ZabbixServiceName).status -ne "Stopped")	
			
			        # Verify service state! Stop if necessary!
			        if ($ZabbixServiceNameState)
			        {
				        # Zabbix agent present and started!
				        try
				        {
					       
					        Write-Host " Stopping Zabbix Agent..."
					        Set-Service -ComputerName $computer -Status Stopped -Name $ZabbixServiceName -ErrorAction Stop
					        Start-Sleep -Seconds 1
				        }
				        catch
				        {
					        Write-Host "##############################################################"
					        Write-Host " $computer - Zabbix agent could not be installed!"
					        Write-Host " Service operation error!" $_			
					        Write-Host "##############################################################"
                    
                            $instBad = addComputerToString $instBad $computer

                            continue  # Next server!			
				        }

				        Write-Host " Zabbix service stopped."				      
			        }			       	
		        }

		        else
		        {
			        Write-Host "################################################################################"
			        Write-Host " Service $ZabbixServiceName not found."				# Ok!
			        
			    }



                # Copy over zabbix folder, make sure no FW could block this
		        try
		        {
			        # If folder exists, it will be overwritten. 
			        # At this point, service should've been stopped...
		
			        $session = New-PSSession -ComputerName $computer -Credential $result



			        $launch_install = Invoke-Command -ArgumentList $ZabbixDirectoryInstall,$ZabbixPath -Session $session  -ScriptBlock {
			            param($ZabbixDirectoryInstall,$ZabbixPath)
	            		if($PSVersionTable.PSVersion.Major -ge 4)
	            		{	

		                		try { $result1 = Invoke-WebRequest  -TimeoutSec 5 -Uri "$ZabbixPath/zabbix_agentd.win.conf" -UseBasicParsing }	# Check binaire presence on web server

			                    catch
			                    {
                                    $launch_install = $False                            # Installation unauthorized
			                        Write-Host -ForegroundColor red $_.Exception.Message
                                    return $launch_install
			                    }

			                    try
			                    { $result2 = Invoke-WebRequest  -TimeoutSec 5 -Uri "$ZabbixPath/zabbix_agentd.exe" -UseBasicParsing } # Check binaire presence on web server

			                    catch 
			                    {
			                        $launch_install = $False                       # Installation unauthorized
			                        Write-Host -ForegroundColor red $_.Exception.Message
                                    return $launch_install
			                    }

								if (!(Test-Path -path "$ZabbixDirectoryInstall" -PathType Container)) { New-Item "$ZabbixDirectoryInstall" -Type Directory }  #Create new repertory

			                    if(($result1.StatusCode -eq 200) -and ($result2.StatusCode -eq 200)) 		# Transfert binaire between web server and remote computer
			                    {
			                        Write-Host " Connected on $ZabbixPath !"
			                        Invoke-WebRequest -Uri "$ZabbixPath/zabbix_agentd.win.conf" -UseBasicParsing  -OutFile "$ZabbixDirectoryInstall\zabbix_agentd.win.conf"
			                        Invoke-WebRequest -Uri "$ZabbixPath/zabbix_agentd.exe" -UseBasicParsing -OutFile "$ZabbixDirectoryInstall\zabbix_agentd.exe"
			                        Write-Host " Folder successfully copied!"									
			                        $launch_install = $True													# Installation authorized
                                    return $launch_install
			                    }
			            }

			            if(-not $launch_install) {
                        throw New-Object -TypeName System.Exception -ArgumentList "Zabbix files download or folder fail from $computer"
                    	}	

			        }			

		        }

		        catch
		        {
			        Write-Host "################################################################################"
			        Write-Host " $computer - Zabbix agent could not be installed!"
			        Write-Host " Make sure you have administrative rights on target!"
			        Write-Host " Copy error!" $_
			        Write-Host "################################################################################"
			        $instBad = addComputerToString $instBad $computer
                    continue
					# Next Server
		        }
 
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
			        Write-Host "################################################################################"
			        Write-Host " $computer - Zabbix agent could not be installed!"
			        Write-Host " -- Make sure you have administrative rights!"
			        Write-Host " -- Make sure no firewall is blocking communications!"
			        Write-Host " Error!" $_
			        Write-Host "################################################################################"
			        $instBad = addComputerToString $instBad $computer
                    continue
			
			
		        }
		        if($os.OSArchitecture -ne $null)
		        {

                $OS_name = $os.Name.Split("{|}")[0]

			        # Architecture can be determined by $os.OSArchitecture...
			        if ($os.OSArchitecture -eq "64-bit")
			        {
                        Write-Host "################################################################################"
                        Write-Host " OS : $($os_Name)  64 Bits"
				        $osArch = "win64"
			        }
			        elseif($os.OSArchitecture -eq "32-bit")
			        {
                        Write-Host "################################################################################"
				        Write-Host " OS : $($os_Name)  32 Bits"
				        $osArch = "win32"
			        }
			        else
			        {
				        Write-Host "################################################################################"
				        Write-Host " Unknown architecture! Operation Canceled..."
				        Write-Host $osArch
				        Write-Host "################################################################################"
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
                        Write-Host "################################################################################"
				        Write-Host " OS : $($os_Name)  64 Bits"
				        $osArch = "win64"
			        }
			        else
			        {
				        Write-Host "################################################################################"
				        Write-Host " OS : $($os_Name)  32 Bits"
				        $osArch = "win32"
			        }
		        }
		        ### Architecture detection ended.

		        ### Begin installation...

		        if ($launch_install -eq $True) 
		        {
			        try
			        {

                        Write-Host "################################################################################"
				        Write-Host " Uninstall agent..."
						action_agent("-d")
			        }
			        catch
			        {
				        Write-Host "################################################################################"
				        Write-Host " Problem while uninstalling previous zabbix agent! Cancelling..."
				        Write-Host $_
				        Write-Host "################################################################################"
				        $instBad = addComputerToString $instBad $computer
				        continue
			        }

			        try
			        {
	                    Write-Host " Install new agent..."

   
                        Invoke-Command -Session $session -ScriptBlock {
                                

                            $myFQDN= (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain
                            $LowerFQDN = $myFQDN.ToLower()

                            gci -LiteralPath "$ZabbixDirectoryInstall\zabbix_agentd.win.conf" -rec -Filter *.conf | % {
                            $regex = '^# Hostname='
                            $line = Select-String -literalpath $_.fullname -pattern $regex | select -ExpandProperty LineNumber


                                $fileContent = Get-Content $ZabbixDirectoryInstall\zabbix_agentd.win.conf
                                $fileContent[$line] += "Hostname=$LowerFQDN"
                                $fileContent | Set-Content $ZabbixDirectoryInstall\zabbix_agentd.win.conf
                            }
                        }   
        
                        action_agent("-i")
			        }

			        catch
			        {
				        Write-Host "################################################################################"
				        Write-Host " Problem while installing new agent! Cancelling..."
				        Write-Host $_
				        Write-Host "################################################################################"
				        $instBad = addComputerToString $instBad $computer
				        continue
			        }

			        try
			        {
				        Write-Host " Starting new agent..."
                        action_agent("-s")
                        Start-Sleep -s 2
                        Write-Host " Restarting new agent..."      
                        action_agent("--stop")
                        Start-Sleep -s 2
                        action_agent("--start")
			        }

			        catch
			        {
				        Write-Host "################################################################################"
				        Write-Host " Problem while starting the agent! Cancelling..."
				        Write-Host $_
				        Write-Host "################################################################################"
				        $instBad = addComputerToString $instBad $computer
				        continue
				
				
			        }
			        ### Installation end...
			        ### Start verification
			        try { $InstallStatus = Get-Service -ComputerName $computer -Name $ZabbixServiceName }

			        catch
			        {
				        Write-Host "################################################################################"
				        Write-Host " Problem while verifying service!"
				        Write-Host " Error! " $_
				        Write-Host "################################################################################"
				        $instBad = addComputerToString $instBad $computer
				        continue
				
				
			        }
			        if ($InstallStatus.Status -eq "Running") { Write-Host " Service installed and started!" }

			        else
			        {
				        Write-Host "################################################################################"
				        Write-Host " Service installed but not started!"
				        Write-Host " Service state: " $InstallStatus.Status		
			        }
		
	                $instOk = addComputerToString $instOk $computer
			           
			    }

	            Else
	            {
		            Write-Host $computer "is not available! Skipping!" -ForegroundColor Red
		            $instBad = addComputerToString $instBad $computer
		            continue
																            # Next server
	            }
                       
	        }

	        Else
	        {
                Write-Host "########################################################################"
		        Write-Host $computer "is not available! Skipping!" -foregroundcolor "red"
		        $instBad = addComputerToString $instBad $computer
		        continue														        # Next server
	        }

        }

        
		if ($session -ne $null) {
		    Remove-PSSession -Session $session  # remove active session
		}
        

        Write-Host "########################################################################"
        Write-Host " Successful installations: " $instOk -foregroundcolor "green"
        Write-Host " Unsuccessful installations: " $instBad -foregroundcolor "red"
        Write-Host "########################################################################"
        Write-Host " "
        Read-Host " Press any key to finish !"
        exit 0
            }

#}







