@echo off
SET workingdir=%cd%
Powershell -ExecutionPolicy ByPass -File .\Backup-OutlookProfile.ps1 -BackupPath D: -BackupMode Export  -Compress Yes -Temp Yes
pause




