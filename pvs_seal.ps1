#PVS Sealing Script

#Global Variables
$Global:OSName = $win32OS.caption
$Global:OSVersion = $win32OS.version
$Global:OSBitness = $win32OS.OSArchitecture
IF ($OSBitness -eq "32-bit") {
        $Global:ProgramFilesx86 = "${env:ProgramFiles}"
        $Global:CommonProgramFilesx86 = "${env:CommonProgramFiles}"
    } ELSE {
        $Global:ProgramFilesx86 = "${env:ProgramFiles(x86)}"
        $Global:CommonProgramFilesx86 = "${env:CommonProgramFiles(x86)}
    }

#### Functions 
Function Write-Log {
	Param(
	    [Parameter(Mandatory=$True)][Alias('M')][String]$Msg,
	    [Parameter(Mandatory=$False)][Alias('C')][String]$Color = "Yellow"
	)
	
	$date = get-date -Format G
	Write-Host "$date - $Msg" -ForegroundColor $Color
	
}

function Show-ProgressBar{
    PARAM(
		[parameter()][string]$CheckProcess,
        [parameter()][int]$CheckProcessId,
		[parameter(Mandatory=$True)][string]$ActivityText,
		[parameter()][int]$MaximumExecutionMinutes = 30,
        [parameter()][switch]$TerminateRunawayProcess
	)
	$a=0
	if ($MaximumExecutionMinutes) {
		$MaximumExecutionTime = (Get-Date).AddMinutes($MaximumExecutionMinutes)
	} ELSE {
        $MaximumExecutionMinutes = 60
        $MaximumExecutionTime = (Get-Date).AddMinutes($MaximumExecutionMinutes)
    }
	Start-Sleep 5
    for ($a=0; $a -lt 100; $a++) {
		IF ($a -eq "99") {$a=0}
		If ($CheckProcessId)
		{
			$ProcessActive = Get-Process -Id $CheckProcessId -ErrorAction SilentlyContinue
		} else {
			$ProcessActive = Get-Process $CheckProcess -ErrorAction SilentlyContinue
		}
		#$ProcessActive = Get-Process $CheckProcess -ErrorAction SilentlyContinue  #26.07.2017 MS: comment-out:

        if ((Get-Date) -ge $MaximumExecutionTime) {
			if ($TerminateRunawayProcess) {
                Stop-Process $ProcessActive -Force -ErrorAction SilentlyContinue
                Clear-Variable -Name "ProcessActive"
            }
            else {
                Clear-Variable -Name "ProcessActive" #this nulls out the variable allowing the "finish" bar
            }
		}

	   	if($ProcessActive -eq $null) {
           	$a=100
           	Write-Progress -Activity "Finish...wait for next operation in 5 seconds" -PercentComplete $a -Status "Finish."
           	IF ($State -eq "Preparation") {Start-Sleep 5}
            Write-Progress "Done" "Done" -completed
            break
       	} else {
            Start-Sleep 1
            $display= "{0:N2}" -f $a #reduce display to 2 digits on the right side of the numeric
            Write-Progress -Activity "$ActivityText" -PercentComplete $a -Status "Please wait..."
       	}
    }
}

function Start-ProcWithProgBar {
 
	PARAM(
		[parameter(Mandatory=$True)][string]$ProcPath,
		[parameter(Mandatory=$True)][string]$Args,
		[parameter(Mandatory=$True)][string]$ActText
	)

	$ChkProc = [io.fileinfo] "$ProcPath" | % basename  # get name from executable without path and extension
	Write-Log -Msg "Starting Process $ProcPath with ArgumentList $Args"
	Start-Process -FilePath "$ProcPath" -ArgumentList "$Args" -NoNewWindow | Out-Null
	Show-ProgressBar -CheckProcess $ChkProc -ActivityText $ActText
	   
}

function Enforce-PVSBasicSettings {
	# PVS Target settings
	reg add "HKLM\SYSTEM\CurrentControlSet\services\BNNS\Parameters" /v EnableOffload /t REG_DWORD /d 00000000 /f
	reg add "HKLM\SYSTEM\CurrentControlSet\services\Tcpip\Parameters" /v DisableTaskOffload /t REG_DWORD /d 00000001 /f
}

Function Rearm-Office {
	# Find version of office and rearm it....
	# Check the installation path of Office 2010
	$Office2010InstallRoot = $null
	If ([Environment]::Is64BitOperatingSystem) {
		$Office2010InstallRoot = (Get-ItemProperty -Path Registry::HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot -Name Path -ErrorAction SilentlyContinue).Path
	}
	If ($Office2010InstallRoot -isnot [system.object]) {$Office2010InstallRoot = (Get-ItemProperty -Path Registry::HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot -Name Path -ErrorAction SilentlyContinue).Path }

	# Check the installation path of Office 2013
	$Office2013InstallRoot = $null
	If ([Environment]::Is64BitOperatingSystem) {
		$Office2013InstallRoot = (Get-ItemProperty -Path Registry::HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot -Name Path -ErrorAction SilentlyContinue).Path
	}
	If ($Office2013InstallRoot -isnot [system.object]) {$Office2013InstallRoot = (Get-ItemProperty -Path Registry::HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot -Name Path -ErrorAction SilentlyContinue).Path }

	# Check the installation path of Office 2016
	$Office2016InstallRoot = $null
	If ([Environment]::Is64BitOperatingSystem) {
		$Office2016InstallRoot = (Get-ItemProperty -Path Registry::HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot -Name Path -ErrorAction SilentlyContinue).Path
	}
	If ($Office2016InstallRoot -isnot [system.object]) {$Office2016InstallRoot = (Get-ItemProperty -Path Registry::HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot -Name Path -ErrorAction SilentlyContinue).Path }

	


	# Activate the office version if installed
	$result = $null
	IF ($Office2010InstallRoot -is [System.Object]) {
		Write-Log -msg "Office 2010 is installed" -ShowConsole -Color Cyan
		Start-ProcWithProgBar -ProcPath "$env:windir\system32\cscript.exe" -Args "//NoLogo ""$($Office2010InstallRoot)OSPP.VBS"" /act" -ActText "Start triggering activation"
		Start-ProcWithProgBar -ProcPath "$env:windir\system32\cscript.exe" -Args "//NoLogo ""$($Office2010InstallRoot)OSPP.VBS"" /dstatus" -ActText "Get Office Licensing state"
	}
 ELSE {
		Write-Log -msg "Office 2010 is NOT installed"
	}

	IF ($Office2013InstallRoot -is [System.Object]) {
		Write-Log -msg "Office 2013 is installed" -ShowConsole -Color Cyan
		Start-ProcWithProgBar -ProcPath "$env:windir\system32\cscript.exe" -Args "//NoLogo ""$($Office2013InstallRoot)OSPP.VBS"" /act" -ActText "Start triggering activation"
		Start-ProcWithProgBar -ProcPath "$env:windir\system32\cscript.exe" -Args "//NoLogo ""$($Office2013InstallRoot)OSPP.VBS"" /dstatus" -ActText "Get Office Licensing state"
	}
 ELSE {
		Write-Log -msg "Office 2013 is NOT installed"
	}


	IF ($Office2016InstallRoot -is [System.Object]) {
		Write-Log -msg "Office 2016 is installed" -ShowConsole -Color Cyan
		Start-ProcWithProgBar -ProcPath "$env:windir\system32\cscript.exe" -Args "//NoLogo ""$($Office2016InstallRoot)OSPP.VBS"" /act" -ActText "Start triggering activation"
		Start-ProcWithProgBar -ProcPath "$env:windir\system32\cscript.exe" -Args "//NoLogo ""$($Office2016InstallRoot)OSPP.VBS"" /dstatus" -ActText "Get Office Licensing state"
	}
 ELSE {
		Write-Log -msg "Office 2016 is NOT installed"
	}




}

function Rearm-OS {
	Write-Log -Msg "Operating System will be rearmed now" -Color DarkCyan
	Start-ProcWithProgBar -ProcPath "$env:windir\system32\cscript.exe" -Args "//NoLogo $env:windir\system32\slmgr.vbs /rearm" -ActText "OS - Reset OS License state"
	Start-ProcWithProgBar -ProcPath "$env:windir\system32\cscript.exe" -Args "//NoLogo $env:windir\system32\slmgr.vbs /dlv" -ActText "OS - Get detailed license informations"
}

Function Disable-ScheduleTasks {
	Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Application Experience\AitAgent"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Application Experience\ProgramDataUpdater"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Application Experience\StartupAppTask"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Autochk\Proxy"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Bluetooth\UninstallDeviceTask"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Customer Experience Improvement Program\BthSQM"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Customer Experience Improvement Program\Consolidator"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Customer Experience Improvement Program\KernelCeipTask"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Customer Experience Improvement Program\Uploader"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Customer Experience Improvement Program\UsbCeip"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Diagnosis\Scheduled"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticResolver"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Maintenance\WinSAT"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\MobilePC\HotStart"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Power Efficiency Diagnostic\AnalyzeSystem"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\RAC\RacTask"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Ras\MobilityManager"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Shell\FamilySafetyMonitor"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Shell\FamilySafetyRefresh"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\SideShow\AutoWake"" /disable" -Wait -WindowStyle Hidden
	Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\SideShow\GadgetManager"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\SideShow\SessionAgent"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\SideShow\SystemDataProviders"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\UPnP\UPnPHostConfig"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\WDI\ResolutionHost"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Windows Filtering Platform\BfeOnServiceStartTypeChange"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\Windows Media Sharing\UpdateLibrary"" /disable" -Wait -WindowStyle Hidden
    Start-Process "schtasks.exe" -ArgumentList "/change /tn ""microsoft\windows\WindowsBackup\ConfigNotification"" /disable" -Wait -WindowStyle Hidden
}

function Clean-SystemDrive {
	#Create SageRun Set 11 in the cleanmgr Registry Hive. Used by cleanmgr.exe to clean specific Things like old Logs and MemoryDumps...
	New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*" -Name "StateFlags0011" -Value "2" -PropertyType "DWORD" -Force | Out-Null
	#Delete specific SageRun Set 11 Flags for the "Windows Update Cleanup" Task because WU Cleanup requires a restart to complete the Cleanup.
	Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Update Cleanup" -Name "StateFlags0011" -ErrorAction SilentlyContinue
	
	#Launch Cleanup 
	Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:11" -Wait -WindowStyle Minimized
	
	#Clean up other folders
	
	# C:\Windows\SoftwareDistribution (Cleanup)
	#Clear Old Crash Dumps if they exist: del 
	if (Test-Path "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\CrashDumps") {
		#Deletes Crash dumps if folder exists...
		Remove-Item C:\Windows\ServiceProfiles\NetworkService\AppData\Local\CrashDumps\*.*
	}
	
	#Clear Old Font Caches, if they exist: 
	Remove-Item C:\Windows\ServiceProfiles\LocalService\AppData\Local\FontCache-*.dat 
	
	Remove-Item -Path C:\windows\SoftwareDistribution\* -force
}

Function Defrag-SystemDrive {
	Start-Process -FilePath "defrag.exe" -ArgumentList "/d C:" -Wait -WindowStyle Minimized
}

Function Execute-NETFrameworkCleanup {
	Get-ChildItem -Path "C:\Windows\Microsoft.NET" -Recurse | Where {$_.Name -eq "ngen.exe"} | Foreach-Object {& "$($_.FullName)" "executequeueditems"}
}

Function Cleanup-Symantec {
	Start-Process -FilePath "C:\Program Files (x86)\Symantec\Virtual Image Exception\vietool.exe" -ArgumentList "c: --generate" -Wait
}

Function Disable-WindowsUpdateService { 
	# Stops and disables Windows Update
	Stop-Service wuauserv
	Set-Service wuauserv -StartType Disabled
	# Show Results 
	get-service wuauserv | Select Name, DisplayName, Status, StartType
}

Function Optimize-TargetOS {
	Start-ProcWithProgBar -ProcPath	"$ProgramFiles64\Citrix\PvsVm\TargetOSOptimizer\TargetOSOptimizer.exe" -Args "/silent" -ActText "Launch Citrix Provisioning Services Target OS Optimizer...."
}

########### MAIN SCRIPT ######################

Enforce-PVSBasicSettings
Optimize-TargetOS
Disable-Scheduletasks
Rearm-Office
Rearm-OS
Cleanup-Symantec
Execute-NETFrameworkCleanup
Disable-WindowsUpdateService
Clean-SystemDrive
Defrag-SystemDrive