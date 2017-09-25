'************************************************************************
'*				Script: MapNetworkDrives
'*					
'*				Date Created: 6-27-2017
'*	
'*					Version: 0.1
'*
'*		This script maps user network drives using the appropriate
'*		user identification including group membership and checks
'*		for duplicate drive associations
'*		
'*		All drives are unmapped during execution then the proper drives
'*		are mapped. Error logs are created for each user upon execution
'*		of the script as well
'*
'*
'*		Created By: Samuel Tayman && Danilo Symonette
'*
'************************************************************************

'On Error Resume Next

'************************************************************************
'*Declare global variables used for processing within script
'*
'*
'************************************************************************
Dim colDrives 				Rem: Stores Network Drives 
Dim EndNotice				Rem: Stores the concatenated end notice for user
Dim GroupMembershipArray	        Rem: Array to hold GroupMembership
Dim GroupMembership			Rem: Holds split Group Membership to contain only included groups
Dim CurrentUser				Rem: Current logged on user 
Dim objUser					Rem: Object to retrieve user data
Dim objNetwork				Rem: Variable for network object 
Dim file					Rem: Variable to store the csv file 
Dim dataArray				Rem: Array stores the csv data 
Dim fso						Rem: FileSystemObject created to open csv file 
Dim logFile					Rem: Log file to store errors that occur
Dim Wsch					Rem: Windows Script Host object
Dim logFileName				Rem: Name for the error log file
Dim strExtensionsToDelete	Rem: List of extensions being deleted
Dim strFolder				Rem: Folder to delete files from
Dim objFSO					Rem: Second fso object for file deletion
Dim MaxAge					Rem: Max age of files allowed
Dim IncludeSubFolders		Rem: Go into subfolders of current directory

'************************************************************************
'*Declare Constants for file use
'*
'*
'************************************************************************
Const ForReading = 1		Rem: Constant for reading the csv file only
Const ForWriting = 2		Rem: Constant for writing to the log file

'************************************************************************
'*Set global variables
'*
'*
'************************************************************************
Rem: Set the WSH 
Set Wsch = WScript.CreateObject("Wscript.Shell")
Rem: Set the FileSystemObject 
Set fso = CreateObject("Scripting.FileSystemObject")
Rem: Set the file using FileSystemObject to the appropriate csv file
Set file = fso.OpenTextFile("DriveMap.csv", ForReading)
Rem: Set network object for network use
Set objNetwork = WScript.CreateObject("WScript.Network")
Rem: Set user object to retrieve active directory information
Set objUser = CreateObject("ADSystemInfo")
Rem: Sets the collection of Drives mapped to the computer
Set colDrives = objNetwork.EnumNetworkDrives
Rem: Set CurrentUser to the logged on user's login credentials
Set CurrentUser = GetObject("LDAP://" & objUser.UserName)
Rem: Directory to delete files from
strFolder = Wsch.CurrentDirectory
Rem: Option to delete files from subfolders
includeSubfolders = TRUE
Rem: Set FSO object to delete old log files
Set objFSO = CREATEOBJECT("Scripting.FileSystemObject")
Rem: A comma separated list of file extensions. Files with extensions provided in the list below will be deleted
strExtensionsToDelete = "txt,log"
Rem: Max File Age (in Days).  Files older than this will be deleted.
maxAge = 1

' ************************************************************
Rem: Create log file name for the current date
logFileName = Wsch.CurrentDirectory & "\" & ConvertDate(Now) & "-" & objNetwork.UserName & ".log" 
Rem: Set the logging file to use 
If(fso.FileExists(logFileName)) Then
	Set logFile = fso.OpenTextFile(logFileName, ForWriting)
Else
	Set logFile = fso.CreateTextFile(logFileName, True)
End If
'************************************************************************
'*Main body of script, removes all currently mapped network drives
'*makes log file edits, and maps drives to user according to 
'*CSV file stored on the local server 
'*
'*
'************************************************************************

Rem: Removes all mapped drives from the computer
For i = 0 to colDrives.Count-1 Step 2

	If(colDrives.Item(i) > "") Then
		objNetwork.RemoveNetworkDrive colDrives.Item(i)
	End If
Next

EndNotice = "Thank you for choosing Peerless Tech Solutions as your IT Support." & VbCrLf

arrGroups = CurrentUser.memberOf

logFile.WriteLine "Logged User: " & objNetwork.UserName

'************************************************************************
'* Function to convert date of YYYYMMDD format
'* This function is used by the logging script to set the
'* date value on the log file name.
'* NOTE: Calls to this function are not logged in the log file
'************************************************************************
Function ConvertDate(ByVal dtdDate)
    ' Convert valid date to different format
    Dim strYear
    Dim strMonth
    Dim strDay
    strYear = Year(dtdDate)
    ' Append 0 to months less than 2 digits
    If Month(dtdDate) < 10 Then
        strMonth = "0" & Month(dtdDate)
    Else
        strMonth = Month(dtdDate)
    End If
    ' Append 0 to days less than 2 digits
    If Day(dtdDate) < 10 Then
        strDay = "0" & Day(dtdDate)
    Else
        strDay = Day(dtdDate)
    End If
    ' Return new date
    ConvertDate = strYear & "_" & strMonth & "_" & strDay
End Function

'************************************************************************
'* Function that deletes the old log files from the 
'* server at the end of the script. Recursively goes through
'* subfolders in directory to delete old files as well.
'*
'*
'************************************************************************

Function DeleteFiles(BYVAL strDirectory,BYVAL strExtensionsToDelete,BYVAL maxAge,includeSubFolders)
	Dim objFolder, objSubFolder, objFile
	Dim strExt

	Set objFolder = objFSO.GetFolder(strDirectory)
	For EACH objFile in objFolder.Files
		For EACH strExt in SPLIT(UCASE(strExtensionsToDelete),",")
			If RIGHT(UCASE(objFile.Path),LEN(strExt)+1) = "." & strExt Then
				If objFile.DateLastModified < (NOW - MaxAge) Then
					objFile.Delete
					EXIT FOR
				END IF
			END IF
		NEXT
	NEXT	
	If includeSubFolders = TRUE THEN 'Recursively delete in folders
		FOR EACH objSubFolder in objFolder.SubFolders
			DeleteFiles objSubFolder.Path,strExtensionsToDelete,maxAge, includeSubFolders
		NEXT
	END IF
End Function


'************************************************************************
'* Function to map drives based on Group membership
'* This function is used depending on the values stored within
'* the client CSV file.
'*
'*
'************************************************************************

Function GroupDrive
	If IsEmpty(arrGroups) Then
				
	Rem: If the user is part of only a single group, this code snippet runs and drives are mapped
	ElseIf (TypeName(arrGroups) = "String") Then
				
		If(Err.Number <> 13) Then
			GroupMembershipArray = Split(CurrentUser.MemberOf, "=")
			GroupMembership = Replace(GroupMembershipArray(1), ",OU", "")
			GroupMembership = Replace(GroupMembership, ",DC", "")
			GroupMembership = Replace(GroupMembership, ",CN", "")
			GroupMembershipArray = GroupMembership
	
		End If
				
		Rem: Map drives based on group membership of user
		If(StrComp(dataArray(1),GroupMembership,1) = 0) Then
		objNetwork.MapNetworkDrive dataArray(3), dataArray(2), 0
		Rem: Error Handling method, compare error with known error strings and report fix or support request to user
		If(Err.Number <> 0) Then	
			If(Err.Number = -2147024811) Then
				'Concatenate the EndNotice to properly display error messages and 
				'completed functions at completion of script
				logFile.WriteLine "Failed Drive Mapping, Drive Already In Use: DRIVE - " & dataArray(3) & VbCrLf & "Error: " & Err.Description & VbCrLf & "Error Number: " & Err.Number
			End If
		End If
		
		If(Err.Number <> 0 And Err.Number <> -2147024811) Then
			logFile.WriteLine "Error: " & Err.Description & VbCrLf & "Error Number: " & Err.Number
		End If
			
	End If
				
	Else
			
		Rem: If user has multiple group memberships, go through each group and map drives with associated group
		For Each strGroup In arrGroups
				
					
			If(Err.Number <> 13) Then
				GroupMembershipArray = Split(strGroup, "=")
				GroupMembership = Replace(GroupMembershipArray(1), ",OU", "")
				GroupMembership = Replace(GroupMembership, ",DC", "")
				GroupMembership = Replace(GroupMembership, ",CN", "")
				GroupMembershipArray = GroupMembership
				
			End If
					
			If(StrComp(dataArray(1),GroupMembership,1) = 0) Then
				objNetwork.MapNetworkDrive dataArray(3), dataArray(2), 0
			End If
		Next
				
	End If


End Function

'************************************************************************
'* Function to map drives based on User membership
'* This function is used depending on the values stored within
'* the client CSV file.
'*
'*
'************************************************************************

Function UserDrive
	If(StrComp(objNetwork.UserName, dataArray(1), 1) = 0) Then
		objNetwork.MapNetworkDrive dataArray(3), dataArray(2), 0
	End If
	
	
End Function	

'************************************************************************
'*Read the CSV file from server and map network drives based on
'*if a user is part of a group or if the user has a personal 
'*mapped drive
'*
'*NOTE: Errors are handled throughout the function code, not all
'*errors are handled however updates should be made to include more 
'*possible error outcomes
'*
'*
'************************************************************************


Rem: Loop goes through file and maps each group with their appropriate drive
Do While not file.AtEndOfStream
	dataArray = Split(file.ReadLine, ",")
		If(StrComp(dataArray(0),"Group",1) = 0) Then 
			GroupDrive
			
		ElseIf(StrComp(dataArray(0), "User", 1) = 0) Then
			UserDrive

		End If
		If(StrComp(dataArray(0), "",1) = 1) Then 
			If(StrComp(dataArray(3), "",1) = 0) Then
				Call Err.Raise(65535, "Network Path", "Drive letter is not defined.")
			End If
		End If
		
		If(Err.Number <> 0) Then
			If(Err.Number = -2147024811) Then
				'Concatenate the EndNotice to properly display error messages and 
				'completed functions at completion of script
				logFile.WriteLine "Error: " & Err.Number & VbCrLf & "Failed Drive Mapping, Drive Already In Use: DRIVE - " & dataArray(3) & VbCrLf & "Error: " & Err.Description & VbCrLf & "Error Number: " & Err.Number
			ElseIf(Err.Number = -2147024843) Then
				logFile.WriteLine "Error: " & Err.Number & VbCrLf & "Network path not found, Path: " & dataArray(2) & ". Correct CSV file to resolve this error."
			ElseIf(Err.Number = -2147023696) Then
				logFile.WriteLine "Error: " & Err.Number & VbCrLf & "Invalid drive letter input in CSV file, Value: " & dataArray(3) & ". Correct CSV file to resolve this error."
			ElseIf(Err.Number = 65535) Then
				logFile.WriteLine "Error has occurred, no drive letter has been assigned for the following: " & VbCrLf & "Type: " & dataArray(0) & VbCrLf & "Name: " & dataArray(1) & VbCrLf & "Path: " & dataArray(2)
			Else
				logFile.WriteLine "Error: " & Err.Description & VbCrLf & "Error Number: " & Err.Number
			End If
		End If
	Err.Clear
	
Loop

'EndNotice = EndNotice & VbCrLf & VbCrLf & "For any of your IT Support needs contact Peerless Tech Solutions at" & VbCrLf & "Phone#: 301-539-4227" & VbCrLf & "Email@: support@getpeerless.com"

'CreateObject("WScript.Shell").Popup EndNotice, 3

Rem: Close file after use
file.close

Rem: Delete old log files from server
DeleteFiles strFolder,strExtensionsToDelete, maxAge, includeSubFolders