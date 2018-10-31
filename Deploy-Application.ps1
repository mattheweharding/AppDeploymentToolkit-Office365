<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
	# LICENSE #
	PowerShell App Deployment Toolkit - Provides a set of functions to perform common application deployment tasks on Windows. 
	Copyright (C) 2017 - Sean Lillis, Dan Cunningham, Muhammad Mashwani, Aman Motazedian.
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. 
	You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK 
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}
	
	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'Microsoft'
	[string]$appName = 'Office 365 (2016 ProPlus)'
	[string]$appVersion = '16.0.9126.2275'
	[string]$appArch = 'AMD64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '09/07/2018'
	[string]$appScriptAuthor = 'Matthew Harding'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = 'Microsoft Office 365'
	[string]$installTitle = 'Microsoft Office 365'
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.7.0'
	[string]$deployAppScriptDate = '02/13/2018'
	[hashtable]$deployAppScriptParameters = $psBoundParameters
	
	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent
	
	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}
	
	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================
		
	If ($deploymentType -ine 'Uninstall') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'

		## Show Welcome Message, close Office if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
        Stop-Service -Name ClickToRunSvc -Force
		Show-InstallationWelcome -CloseApps 'officeclicktorun,ose,osppsvc,sppsvc,msoia,excel,groove,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,winproj,visio,iexplore' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt
		
		## Show Progress Message (with the default message)
		Show-InstallationProgress
		
		## <Perform Pre-Installation tasks here>
        Show-InstallationProgress -StatusMessage ‘Uninstalling previous versions of Microsoft Office. This may take up to 30 minutes. Please wait...’ -TopMost $True
        
        #Uninstalling Microsoft Office using the GitHut Removal Scripts at https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls
        Invoke-Expression "$dirSupportFiles\Remove-PreviousOfficeInstalls\Remove-PreviousOfficeInstalls.ps1"

        #Uninstalling Microsoft Office using the GitHut Removal Scripts at https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-OfficeClickToRun
        Invoke-Expression "$dirSupportFiles\Remove-OfficeClickToRun\Remove-OfficeClickToRun.ps1"

        #Delete Office 2016 Shortcuts on the Public Desktop
        Remove-Item -Path 'C:\Users\Public\Desktop\Access 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\Outlook 2016.lnk' -Force -ErrorAction SilentlyContinue
		Remove-Item -Path 'C:\Users\Public\Desktop\Excel 2016.lnk' -Force -ErrorAction SilentlyContinue
		Remove-Item -Path 'C:\Users\Public\Desktop\OneNote 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\PowerPoint 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\Publisher 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\Word 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\OneDrive for Business.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\Skype for Business.lnk' -Force -ErrorAction SilentlyContinue
		
		
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## Handle Zero-Config MSI Installations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) { $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ } }
		}
		
		## <Perform Installation tasks here>
		Show-InstallationProgress -StatusMessage ‘Installing Microsoft Office 365 (2016 ProPlus). This may take up to 30 minutes. Please wait...’ -TopMost $True
        
        #Starting the default Office 2016 ProPlus Installer
        Execute-Process -Path “$dirFiles\setup.exe” -Parameters “/CONFIGURE install.xml”

		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>

		Show-InstallationProgress -StatusMessage ‘Creating desktop shortcuts...’ -TopMost $True

		#Create desktop shortcut for Access 2016
		Start-Process -FilePath "$dirSupportFiles\Create-DesktopShortcuts\ShortcutCreator.exe" -ArgumentList '-location="Desktop" -name="Access 2016" -target="C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE" -args="" -startin=""'

		#Create desktop shortcut for Excel 2016
		Start-Process -FilePath "$dirSupportFiles\Create-DesktopShortcuts\ShortcutCreator.exe" -ArgumentList '-location="Desktop" -name="Excel 2016" -target="C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" -args="" -startin=""'
 
		#Create desktop shortcut for OneNote 2016
		Start-Process -FilePath "$dirSupportFiles\Create-DesktopShortcuts\ShortcutCreator.exe" -ArgumentList '-location="Desktop" -name="OneNote 2016" -target="C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE" -args="" -startin=""'
 
		#Create desktop shortcut for Outlook 2016
		Start-Process -FilePath "$dirSupportFiles\Create-DesktopShortcuts\ShortcutCreator.exe" -ArgumentList '-location="Desktop" -name="Outlook 2016" -target="C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" -args="" -startin=""'

		#Create desktop shortcut for OneDrive for Business
		Start-Process -FilePath "$dirSupportFiles\Create-DesktopShortcuts\ShortcutCreator.exe" -ArgumentList '-location="Desktop" -name="OneDrive for Business" -target="C:\Program Files\Microsoft Office\root\Office16\GROOVE.EXE" -args="" -startin=""'

		#Create desktop shortcut for PowerPoint 2016
		Start-Process -FilePath "$dirSupportFiles\Create-DesktopShortcuts\ShortcutCreator.exe" -ArgumentList '-location="Desktop" -name="PowerPoint 2016" -target="C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" -args="" -startin=""'
 
		#Create desktop shortcut for Publisher 2016
		Start-Process -FilePath "$dirSupportFiles\Create-DesktopShortcuts\ShortcutCreator.exe" -ArgumentList '-location="Desktop" -name="Publisher 2016" -target="C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE" -args="" -startin=""'

		#Create desktop shortcut for Skype for Business
		Start-Process -FilePath "$dirSupportFiles\Create-DesktopShortcuts\ShortcutCreator.exe" -ArgumentList '-location="Desktop" -name="Skype for Business" -target="C:\Program Files\Microsoft Office\root\Office16\lync.exe" -args="" -startin=""'

		#Create desktop shortcut for Word 2016
		Start-Process -FilePath "$dirSupportFiles\Create-DesktopShortcuts\ShortcutCreator.exe" -ArgumentList '-location="Desktop" -name="Word 2016" -target="C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" -args="" -startin=""'

		## Display a message at the end of the install
		If (-not $useDefaultMsi) { Show-InstallationPrompt -Message 'Microsoft Office 365 has been installed.' -ButtonRightText 'OK' -Icon Information -NoWait }
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'
		
		## Show Welcome Message, close Office with a 60 second countdown before automatically closing
		Show-InstallationWelcome -CloseApps 'officeclicktorun,ose,osppsvc,sppsvc,msoia,excel,groove,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,winproj,visio,iexplore' -CloseAppsCountdown 60
		
		## Show Progress Message (with the default message)
		Show-InstallationProgress
		
		## <Perform Pre-Uninstallation tasks here>
		
		
		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'
		
		## Handle Zero-Config MSI Uninstallations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat
		}
		
		# <Perform Uninstallation tasks here>

        Show-InstallationProgress -StatusMessage ‘Uninstalling previous versions of Microsoft Office. This may take up to 30 minutes. Please wait...’ -TopMost $True

        #Uninstalling Microsoft Office using the GitHut Removal Scripts at https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls
        Invoke-Expression "$dirSupportFiles\Remove-PreviousOfficeInstalls\Remove-PreviousOfficeInstalls.ps1"

        #Uninstalling Microsoft Office using the GitHut Removal Scripts at https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/tree/master/Office-ProPlus-Deployment/Remove-OfficeClickToRun
        Invoke-Expression "$dirSupportFiles\Remove-OfficeClickToRun\Remove-OfficeClickToRun.ps1"

        #Delete Office 2016 Shortcuts on the Public Desktop
        Remove-Item -Path 'C:\Users\Public\Desktop\Access 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\Outlook 2016.lnk' -Force -ErrorAction SilentlyContinue
		Remove-Item -Path 'C:\Users\Public\Desktop\Excel 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\PowerPoint 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\Publisher 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\Word 2016.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\OneDrive for Business.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\Skype for Business.lnk' -Force -ErrorAction SilentlyContinue
        Remove-Item -Path 'C:\Users\Public\Desktop\OneNote 2016.lnk' -Force -ErrorAction SilentlyContinue

		
		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'
		
		## <Perform Post-Uninstallation tasks here>
		
		
	}
	
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================
	
	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}