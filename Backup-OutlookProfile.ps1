#.SYNOPSIS
#The following script will export or import Outlook profiles.
#
#.EXAMPLE
#Example Run: .\Backup-OutlookProfile.ps1 -BackupMode Export -BackupPath G: -Compress Yes -Temp Yes
#Example Run: .\Backup-OutlookProfile.ps1 -BackupMode Import -BackupPath G: -Temp Yes
#
#
#.NOTES 
#        File Name  : Backup-OutlookProfile.ps1
#        Authors    : Neil Hennessy (Action IT Services)
#		 Contact	: info@actionitservices.com.au	
#        Requires   : PowerShell Version 5.1
#
#        Version   : 0.1 (01/07/18) - Initial Script
#        Version   : 0.2 (01/07/18) - 
#        Version   : 0.3 (03/10/18) - Improved input error checking and changed log line for script complete.
#        Version   : 0.4 (03/10/18) - Added 7-Zip compression features
#        Version   : 0.5 (04/10/18) - Fixed profile folder name retrieval, where it does not match the username. e.g. domain.user.
#        Version   : 0.6 (08/10/18) - Fixed removing temp folder when compressing the backup. Added optional switch to use %TEMP% to compress/decompress 7-zip backups.
#        Version   : 0.7 (08/10/18) - Added List mode to display list of backups and sorted the backups in Import and List modes.
#        Version   : 0.8 (08/10/18) - Fixed regex on Import selection to allow higher than 10.
#        Version   : 0.9 (08/10/18) - PowerShell Version Check


Param ( 
	[Parameter(mandatory=$true)][string]$BackupMode,
	[Parameter(mandatory=$true)][string]$BackupPath,
	[Parameter(mandatory=$false)][string]$Compress,
	[Parameter(mandatory=$false)][string]$Temp
)

$myVersion = "0.9"
$theDate = (Get-Date).ToString()
$logDate = $theDate.replace("/","-")
$logDate = $logDate.replace(":","-")
$logDate = $logDate.replace(" ","-")


Function PSVersionCheck{
$psVersion = $PSVersionTable.PSVersion.ToString()
	if (($psVersion.SubString(0,1)) -lt "5") {
		Write-host "Backup-OutlookProfile: Abort script - Invalid PowerShell Version ($psVersion)" -ForegroundColor Red
		Invoke-Expression "cmd.exe /C start https://www.microsoft.com/en-us/download/details.aspx?id=54616"
		Exit
	}
}

$systemUser = "$env:UserName"
$systemName = "$env:COMPUTERNAME"
$profilePath = "$env:userprofile"
$profileTEMP = "$env:TEMP"
$systemDomain = (Get-WmiObject Win32_ComputerSystem).Domain
$defaultBackupName = "OutlookProfileBackups"

Write-host "Backup-OutlookProfile: Started $myVersion $theDate"

PSVersionCheck

Write-host "Backup-OutlookProfile: ComputerName = $systemName"
Write-host "Backup-OutlookProfile: User = $systemUser"
Write-host "Backup-OutlookProfile: TEMP = $profileTemp"
Write-host "Backup-OutlookProfile: Domain = $systemDomain"
Write-host "Backup-OutlookProfile: Profile Path = $profilePath"

$ExportJobRoot = "$BackupPath\$defaultBackupName"	
$ExportJobPath = "$BackupPath\$defaultBackupName\$systemUser-$logDate\"

Write-host "Backup-OutlookProfile: Backup Path = $ExportJobRoot"

$outInstPath = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" | Select-Object -ExpandProperty "Path"
$outVerNumber = ($outInstPath.SubString($outInstPath.length-3).TrimEnd("\"))
$outProfileRegPath = "HKCU:\Software\Microsoft\Office\" + $outVerNumber + ".0\Outlook\Profiles"

$7zEXEPath = "bin\7z1805-extra"
$7zJobFile = "$BackupPath\$defaultBackupName\$systemUser-$logDate.7z"
$7zJobTempPath = "$ExportJobRoot\temp"

If ($Compress -eq "") {
	$Compress = "No"
}

If ($Temp -eq "") {
	$Temp = "No"
}

If ($Temp -eq "Yes") {
	$ExportJobPath = "$profileTEMP\$defaultBackupName\$systemUser-$logDate\"
}

switch ($BackupMode) 
{
	"Export" {
		Write-host "Backup-OutlookProfile: Export Outlook Profiles"
		If (Test-path $BackupPath) {
			Write-host "Backup-OutlookProfile: Creating export folders ...please wait" -ForegroundColor Yellow			
			Write-host "Backup-OutlookProfile: Backup Path = $ExportJobPath"
			New-Item $ExportJobPath -ItemType "directory" -Force | Out-Null 
			New-Item "$ExportJobPath\Local" -ItemType "directory" -Force | Out-Null
			New-Item "$ExportJobPath\Roaming" -ItemType "directory" -Force | Out-Null
			New-Item "$ExportJobPath\Signatures" -ItemType "directory" -Force | Out-Null
			
			If((Test-path $ExportJobPath) -eq $false) {
				Write-host "Backup-OutlookProfile: Error, problem creating export folder!"  -ForegroundColor Red
				break
			}			
			$outProfileCMD = "HKCU\Software\Microsoft\Office\" + $outVerNumber + ".0\Outlook\Profiles"
			$outAppDataLocal = "$profilePath\AppData\Local\Microsoft\Outlook"
			$outAppDataRoaming = "$profilePath\AppData\Roaming\Microsoft\Outlook"
			$userSignatures = "$profilePath\AppData\Roaming\Microsoft\Signatures"
			
			If (Test-path $outAppDataLocal) {
				Write-host "Backup-OutlookProfile: Exporting AppData Local folder ...please wait" -ForegroundColor Yellow
				Copy-Item -Path "$outAppDataLocal" -Destination "$ExportJobPath\Local\" -Recurse
			} else {
				Write-host "Backup-OutlookProfile: Error, $outAppDataLocal does NOT exist" -ForegroundColor Red	
			}
			If (Test-path $outAppDataRoaming) {
				Write-host "Backup-OutlookProfile: Exporting AppData Roaming folder ...please wait" -ForegroundColor Yellow
				Copy-Item -Path "$outAppDataRoaming" -Destination "$ExportJobPath\Roaming\" -Recurse
			} else {
				Write-host "Backup-OutlookProfile: Error, $outAppDataRoaming does NOT exist" -ForegroundColor Red	
			}
			If (Test-path $userSignatures) {
				Write-host "Backup-OutlookProfile: Exporting $systemUser Signature folder ...please wait" -ForegroundColor Yellow
				Copy-Item -Path "$userSignatures" -Destination "$ExportJobPath\Signatures\" -Recurse
			} else {
				Write-host "Backup-OutlookProfile: Error, $userSignatures does NOT exist" -ForegroundColor Red	
			}
			If (Test-path $outProfileRegPath) {
				Write-host "Backup-OutlookProfile: Exporting HKCU Outlook registry keys ...please wait" -ForegroundColor Yellow
				$ExportFile = "$ExportJobPath\HKCU-Outlook-Profile.reg"		
				$command = "cmd.exe /C reg export ""$outProfileCMD"" ""$ExportFile"""				
				Invoke-Expression -Command:$command
			} else {
				Write-host "Backup-OutlookProfile: Error, Could NOT find Outlook registry profile for user ($systemUser)"  -ForegroundColor Red
				break	
			}

			If ((Test-path $ExportFile) -eq $false) {
				Write-host "Backup-OutlookProfile: Error, Outlook registry profile backup missing!"  -ForegroundColor Red
			}	

			# Compress			
			If ($Compress -eq "Yes") {
				Write-host "Backup-OutlookProfile: Compressing the backup ...please wait" -ForegroundColor Yellow
				$command = "$7zEXEPath\7za.exe a ""$7zJobFile"" ""$ExportJobPath"""				
				Invoke-Expression -Command:$command	
				Write-host "Backup-OutlookProfile: Cleaning up backup files...please wait" -ForegroundColor Yellow
				Remove-Item -Path $ExportJobPath -Force -Recurse | Out-Null	
			}

		} else {
			Write-host "Backup-OutlookProfile: Error, Invalid backup path! ($BackupPath)"  -ForegroundColor Red
		}
	}
	
	"Import" {
		Write-host "Backup-OutlookProfile: Import Outlook Profile"
		Write-host "Backup-OutlookProfile: Important, Your existing Outlook profile will be wiped!!" -ForegroundColor Red
		If (Test-path $ExportJobRoot) {
			Write-host "Backup-OutlookProfile: Searching for backups ($ExportJobRoot) ..please wait"  -ForegroundColor Yellow
			$outBackupJobs = Get-childitem "$ExportJobRoot" | Sort-Object 
			If ($outBackupJobs) {
				$i=0	
				Foreach ($outJob in $outBackupJobs) {
					$i++				
					$outJobImportPath = $outJob.FullName
					Write-Host "Press '$i' to Import $outJobImportPath"								
				}
				Write-Host "Press 'c' to cancel (exit)"	
				$outJobSelection = Read-Host "Please make a selection"
				
				If ($outJobSelection -eq "c" ) {
					Write-Host "Backup-OutlookProfile: Import cancelled.. exiting" -ForegroundColor Red								
					break
				}
				
				if ($outJobSelection -notmatch '\d+') {
					Write-Host "Backup-OutlookProfile: Invalid selection.. exiting" -ForegroundColor Red								
					break
				}
				
				#if ($outJobSelection -ge $i) {
					#Write-Host "Backup-OutlookProfile: Backup doesn't exist $outJobSelection = ($i).. exiting" -ForegroundColor Red								
					#break
				#}
				
				$theIndex =($outJobSelection - 1)						
				$outJobImportPath = $outBackupJobs[$theIndex].FullName

				# Decompress	
				if ($outJobImportPath -like '*.7z') {
					Write-Host "Backup-OutlookProfile: Found compressed backup. Extracting.. please wait" -ForegroundColor Yellow
					$command = "$7zEXEPath\7za.exe x ""$outJobImportPath"" ""-o$7zJobTempPath"" -r"					
					Invoke-Expression -Command:$command						
					$7zJobImportFiles = Get-ChildItem -Path $7zJobTempPath
					$outJobImportPath = "$7zJobTempPath\$7zJobImportFiles"
				}
								
				$outAppDataLocal = "$profilePath\AppData\Local\Microsoft\Outlook\"
				$outAppDataRoaming = "$profilePath\AppData\Roaming\Microsoft\Outlook\"
				$userSignatures = "$profilePath\AppData\Roaming\Microsoft\Signatures\"
				$outProfileRegPath = "HKCU:\Software\Microsoft\Office\" + $outVerNumber + ".0\Outlook\Profiles"
				$outAppDataLocalDst = "$profilePath\AppData\Local\Microsoft\"
				$outAppDataRoamingDst = "$profilePath\AppData\Roaming\Microsoft\"
				$userSignaturesDst = "$profilePath\AppData\Roaming\Microsoft\"	
				$outAppDataLocalSrc = "$outJobImportPath\Local\Outlook"
				$outAppDataRoamingSrc = "$outJobImportPath\Roaming\Outlook"
				$userSignaturesSrc = "$outJobImportPath\Signatures\Signatures"
				$outProfileRegFile = "$outJobImportPath\HKCU-Outlook-Profile.reg"		
				
				Write-Host "Backup-OutlookProfile: Using backup $outJobImportPath"
				Write-Host "Backup-OutlookProfile: Cleaning up existing profiles ...please wait" -ForegroundColor Yellow
				
				If (Test-path $outAppDataLocal) {
					Remove-Item $outAppDataLocal -Force -Recurse | Out-Null
				}
				If (Test-path $outAppDataRoaming) {
					Remove-Item $outAppDataRoaming -Force -Recurse | Out-Null
				}
				If (Test-path $userSignatures) {
					Remove-Item $userSignatures -Force -Recurse | Out-Null
				}
				If (Test-path $outProfileRegPath) {
					Remove-Item $outProfileRegPath -Force -Recurse | Out-Null
				}
				
				Write-Host "Backup-OutlookProfile: Starting import process ...please wait" -ForegroundColor Yellow
				
				Copy-Item -Path "$outAppDataLocalSrc" -Destination "$outAppDataLocalDst" -Recurse
				Copy-Item -Path "$outAppDataRoamingSrc" -Destination "$outAppDataRoamingDst" -Recurse
				Copy-Item -Path "$userSignaturesSrc" -Destination "$userSignaturesDst" -Recurse
				$command = "cmd.exe /C reg import ""$outProfileRegFile"""				
				Invoke-Expression -Command:$command

				Write-host "Backup-OutlookProfile: Cleaning up temp backup files...please wait" -ForegroundColor Yellow	
				If (Test-path $7zJobTempPath) {
					Remove-Item -Path $7zJobTempPath -Force -Recurse | Out-Null	
				}
				
			} else {
				Write-host "Backup-OutlookProfile: Error, No backups found in ($BackupPath)"  -ForegroundColor Red
			}
		} else {
			Write-host "Backup-OutlookProfile: Error, Invalid backup path! ($BackupPath)"  -ForegroundColor Red
		}				
			
	}
	
	"List" {
		Write-host "Backup-OutlookProfile: Listing Outlook Profile Backups"
		If (Test-path $ExportJobRoot) {
			Write-host "Backup-OutlookProfile: Searching for backups ($ExportJobRoot) ..please wait"  -ForegroundColor Yellow
			$outBackupJobs = Get-childitem "$ExportJobRoot" | Sort-Object 
			If ($outBackupJobs) {
				$i=0	
				Foreach ($outJob in $outBackupJobs) {
					$i++				
					$outJobImportPath = $outJob.FullName
					Write-Host "[$i][$outJobImportPath]"					
				}
			}
		}
	}
}
$ExportFinished = (Get-Date).ToString()
Write-host "Backup-OutlookProfile: $BackupMode Completed $ExportFinished"

