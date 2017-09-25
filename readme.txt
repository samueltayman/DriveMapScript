****************************************************************************************
*
* The following is documentation on using the MapNetworkDrives.vbs script.
*
* All of the information stated in this document is for user knowledge and 
* requirements necessary for the script to run at 100% efficiency. 
*
* Intended use as a logon script through Active Directory and Group Policy
*
* NOTE: Users must be given write access to the scripts logon folder in order for
* error reporting to be enabled
*
****************************************************************************************

MAP NETWORK DRIVES EXECUTION REQUIREMENTS

CSV File holding data in the following format:
----------------------------------------------------------------------------------------
Column 1: Type (User or Group)
Column 2: Name (User name or Group name)
Column 3: Share (Location of share folder, ie. \\Server\Share) 
Column 4: Drive Letter (Drive letter for network share, ie. X:)

CSV can either be made in EXCEL or in a text document being saved as a CSV file, without UTF-8
Text file CSV's should be formatted as Type,Name,Share,Drive Letter

No extra lines or spaces should be used if a text file format is chosen.

CSV file MUST be named DriveMap or the file will not be seen by the script. 
----------------------------------------------------------------------------------------

****************************************************************************************

CSV and VBS files MUST be stored in the Scripts\Logon folder on the domain server. 

The following steps show how to set up the script for Group Policy:
----------------------------------------------------------------------------------------
Step 1: Access the Group Policy Management console on the domain server that will
distribute the logon script

Step 2: Right-Click the domain and select "Create a GPO in this domain, and Link it here..." option

Step 3: Name the policy and click Okay

Step 4: Under the "Security Filtering" Snap-in select "Add" and add "Domain Users" 

Step 5: Right-Click the GPO and select "Edit"

Step 6: Go to User Configuration -----> Policies -----> Windows Settings -----> Scripts(Logon/Logoff) -----> Logon

Step 7: Select "Add" then "Browse", if the MapNetworkDrives script and DriveMap CSV file are not located in the displayed folder, 
drag the files into this folder then select the MapNetworkDrives vbs file and select "Open" then "OK" and finally "OK"

The script should now run whenever a user logs on to their PC from this point on
-----------------------------------------------------------------------------------------

*****************************************************************************************

EXCEPTIONS AND ERROR LOGGING INFORMATION

The MapNetworkDrives script does support error logging. Error logs are 
created during script runtime and are user specific.

If multiple logon attempts by one user are made on the same day the 
error logs are re-written with the most recent attempt and saved.

Error logs are stored in the same location as the VBS and CSV files and
it is recommended that a shortcut to the location is made for easy access 
and monitoring of log files. It is also recommended that log files
be erased every one to two weeks. 

