# ImagingLogging

This set of scripts is intended to be used in logging time on the Scanco scanners, generate scan times, and generate the csv files required for upload to iLabs. 

mapNetworkDrives will automatically map ortho network drives if they aren't already. The username and password must be set with currently active credentials.

DXASync.cmd and VisionSync.cmd - these scripts are currently run as a scheduled Windows process every night. Together, they mirror the local Faxitron database at "J:\Silva's Lab\P30 Core Center\Faxitron Backup"  While this should be an automated process, running them manually once in a while won't hurt anything, and the Windows task scheduler should occasionally be checked up on.

ImagingLabRecordKeepingIlabs - this script is run, and interacted with by Scanco scanner users, to connect Scanco machine records with iLabs information. The user should simply follow the prompts, filling in the user and PI information with their wustl email addresses.

vivaCTTimeCollector and microCTTimeCollector - these scripts generate a VMS script from the RSQHEADERTEMPLATE.COM template, put that generated VMS script on to the target CT system, and run it. That generates the time information on a per scan basis by reading metadata from the rsq files. It then gathers that data and puts it in a location that is looked for by the billing compilation script, so don't change where they go! Or if you do, make sure to check all the scripts in this project to change them the same way. If Ortho changes the way they provide network storage this may become necessary; if this happens, I recommend changing the storage location of everything to DiskTower.

rsqGet.bat is a batch file that will automatically launch MATLAB and run the time collector scripts, if you set the path to the .m files correctly. It isn't required but is convenient for setting up a Windows Task Scheduler task.
