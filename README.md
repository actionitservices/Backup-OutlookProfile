
# Backup-OutlookProfile

File Name: 	Backup-OutlookProfile.ps1  
Authors: 	Neil Hennessy (Action IT Services)  
Contact: 	info@actionitservices.com.au  
Requires: 	PowerShell Version 5.1  

**Description**

Recover from .OST corruption, restore Autocomplete contacts and perform Outlook profile migration to new end points.  
  
The following script will export and import Outlook profiles for Windows.   
- Outlook profile registry keys are included.   
- Outlook AppData and Signature folders are included.  
- Any existing Outlook profiles will be wiped using the import process.  
- The script will create a folder OutlookProfileBackups to the BackupPath folder.  
- Default run does not use the %TEMP% folder or compression.  
- The import process will scan BackupPath and ask to select which backup to import via number selection.   
  
Use and edit the parramanters in the .cmd files for quick or scheduled running.  
  
Export-OutlookProfile.cmd  
Import-OutlookProfile.cmd  
List-OutlookProfile.cmd  
  
**BackupMode**

Export: (Export Outlook profile to folder or zip)  
Import: (Import or Export Profile to foldert or zip)  
List:	(List of export's availble for import)  
  
**Compress**

Yes: (Use 7-zip to compress export data into single 7z archive)  
No: (No compression. Copy directly to export location)  

**Temp**

Yes: (Copy export data to %TEMP% folder before zipping, then copy to final export location. Use this option to improve export speeds for compression tasks on flash storage)  
No: (Copy directly to export location, then perform zip)  

**Examples**

1. Export to D: Drive directly, no compression, no %TEMP%.  
	.\Backup-OutlookProfile.ps1 -BackupMode Export -BackupPath D:  

2. Export to D: Drive, compressing to 7z, using %TEMP% for compression.  
	.\Backup-OutlookProfile.ps1 -BackupMode Export -BackupPath D: -Compress Yes -Temp Yes  

2.  Import from D: Drive  
	.\Backup-OutlookProfile.ps1 -BackupMode Import -BackupPath D:  

4. Export to D: Drive, compressing to 7z, using %TEMP% for compression.  
	.\Backup-OutlookProfile.ps1 -BackupMode List -BackupPath D:  


**Notes**

When using the import process and the Windows profile name and/or path has changed, a Junction point to the old name is required for Outlook to locate the imported .OST files.  
E.g.	Username/profile path when running the export was Terry.Tester, however the new Windows profile path/name is TerryTester.  
		Create a Junction point using the the following PS:   
		New-Item -ItemType Junction -Path "C:\Users\Terry.Tester" -Value "C:\Users\TerryTester"  
		
		
**Version History**

+ Version   : 0.1 (01/07/18) - Initial Script
+ Version   : 0.2 (01/07/18) - Bug fixes
+ Version   : 0.3 (03/10/18) - Improved input error checking and changed log line for script complete.
+ Version   : 0.4 (03/10/18) - Added 7-Zip compression features
+ Version   : 0.5 (04/10/18) - Fixed profile folder name retrieval, where it does not match the username. e.g. domain.user.
+ Version   : 0.6 (08/10/18) - Fixed removing temp folder when compressing the backup. Added optional switch to use %TEMP% to compress/decompress 7-zip backups.
+ Version   : 0.7 (08/10/18) - Added List mode to display list of backups and sorted the backups in Import and List modes.
+ Version   : 0.8 (08/10/18) - Fixed regex on Import selection to allow higher than 10.
+ Version   : 0.9 (08/10/18) - PowerShell Version Check
