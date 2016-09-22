Option Explicit

'Process arguments
Dim SourceFolder:	SourceFolder	= Wscript.Arguments.Item(0) ' This is the root folder we're going to recursively search
Dim Folder2Archive:	Folder2Archive	= Wscript.Arguments.Item(1) ' This is the name of the folder we're looking for under the root folder

'Helper objects
Dim fso:	set fso = CreateObject("Scripting.FileSystemObject")

WScript.Echo "Searching for folders named " & Folder2Archive & " under folder " & SourceFolder
WScript.Echo ""

'Test for existence of SourceFolder, exit if it doesn't exist.
If not fso.FolderExists(SourceFolder) Then
	WScript.Echo "ERROR: " & SourceFolder & " does not exist!"
	Set fso = Nothing
	WScript.Quit -1
End If

Dim aFolderTree() : ReDim aFolderTree(-1)
Dim iRet

' Recurse through folder structure and return an array of folder names
iRet = BuildFolderTree(SourceFolder, aFolderTree, fso)
If iRet <> 0 Then
	On Error Resume Next
	Err.Raise iRet
	WScript.Echo "ERROR: " & Err.Number & vbCRLF & Err.Description
	Err.Clear
	On Error Goto 0
	WScript.Quit iRet
End If

Dim iCount
Dim iFolderCount:	iFolderCount = 0	' Count number of found archive folders

For iCount = LBound(aFolderTree) to UBound(aFolderTree)
	If IsMatchingFolder(aFolderTree(iCount), Folder2Archive) Then
		iFolderCount = iFolderCount + 1
	End If
Next

WScript.Echo ""
WScript.Echo "Completed searching through " & UBound(aFolderTree) + 1 & " folders."
WScript.Echo ""
WScript.Echo "Found " & iFolderCount & " " & Folder2Archive & " folders."
WScript.Echo ""

Dim sFolder
Dim fFolder
Dim fSubFolder
Dim fFile
Dim iArchivedFileCount
Dim iDaysToKeep:		iDaysToKeep = 30	' TODO: Consider parameterising this at some point...

For iCount = LBound(aFolderTree) to UBound(aFolderTree)
	If IsMatchingFolder(aFolderTree(iCount), Folder2Archive) Then
		sFolder = aFolderTree(iCount)
		WScript.Echo "Archiving files in folder: " & sFolder
		Set fFolder = fso.GetFolder(sFolder)
		iArchivedFileCount = 0

		For Each fFile In fFolder.Files
			If LCase(Right(fFile.Name, 4)) <> ".zip" And  LCase(Right(fFile.Name, 3)) <> ".gz" Then
				' Archive all non-zip files (since they're the archive files!)
				ArchiveFile sFolder, fFile.Name, fso
				iArchivedFileCount = iArchivedFileCount + 1
			End If
		Next

		WScript.Echo "	Archived " & iArchivedFileCount & " files."
		WScript.Echo "Archiving subfolders in folder: " & sFolder
		iArchivedFileCount = 0

		' Archive all subfolders
		For Each fSubFolder In fFolder.SubFolders
			ArchiveFolder sFolder, fSubFolder.Name, fso
			iArchivedFileCount = iArchivedFileCount + 1
		Next

		WScript.Echo "	Archived " & iArchivedFileCount & " folders."
		
		' Delete zip files older than X days.
		WScript.Echo "Deleting archived files older than " & iDaysToKeep & " days in folder: " & sFolder
		For Each fFile In fFolder.Files
			If LCase(Right(fFile.Name, 4)) = ".zip" Or  LCase(Right(fFile.Name, 3)) = ".gz" Then
				If DateDiff("d", fFile.DateLastModified, Date) > iDaysToKeep Then
					Wscript.Echo "Deleting file " & fFile.Name
					fso.DeleteFile(fFile)
				End If
			End If
		Next
	End If
Next

Set fso = Nothing

WScript.Echo ""
WScript.Echo "Done."
WScript.Echo ""

'
' Helper Functions
'

Function IsMatchingFolder(sFolderTreeItem, Folder2Match)

	Dim bMatches: bMatches = False

	If Right(LCase(sFolderTreeItem), Len(Folder2Match)) = LCase(Folder2Match) Then
		bMatches = True
	End If

	IsMatchingFolder = bMatches
End Function

' recurse the passed sRootValid folder returning an array of folders
Function BuildFolderTree(sRootValid, aFolder, fso)
	'errorless return 0
	'else return error number first encountered
	'aFolder return preserveed as the state just before the error encountered.

	BuildFolderTree = 0
	On Error Resume Next

	Dim oFolder:	Set oFolder = fso.GetFolder(sRootValid)
	Dim iFound: 	iFound = oFolder.SubFolders.Count

	If iFound <> 0 Then
		Dim iRef:   iRef   = UBound(aFolder)
		Dim iCount: iCount = 0
		ReDim Preserve aFolder(iRef + iFound)
		For Each oSubFolder in oFolder.SubFolders
			iCount = iCount + 1
			aFolder(iRef + iCount) = oSubFolder.Path
		Next
	End If

	Dim iRetDyn
	
	If oFolder.SubFolders.Count <> 0 Then
		Dim oSubFolder
		For Each oSubFolder in oFolder.SubFolders
			WScript.Echo "	Searching subfolder " & oSubFolder.Path
			iRetDyn = BuildFolderTree(oSubFolder.Path, aFolder, fso)
		Next
	End If
	
	Set oFolder = Nothing
	
	If iRetDyn <> 0 Then
		BuildFolderTree = iRetDyn
	ElseIf Err.Number <> 0 Then
		BuildFolderTree = Err.Number
		Err.clear
	End If
	On Error Goto 0
End Function

Function ArchiveFolder(sFolder, sFolderName, fso)
	
	Dim sSourceFolder:	sSourceFolder	= sFolder & "\" & sFolderName
	Dim fFolder
	Dim sZipFileName
	Dim sZipFile
	
	If fso.FolderExists(sSourceFolder) Then
		set fFolder = fso.GetFolder(sSourceFolder)
		
		if fFolder.Files.Count > 0 Then
			' Folder contains files, so zip the folder
			sZipFileName	= GenerateZipFileName(sFolder, sFolderName, fso)
			sZipFile		= sFolder & "\" & sZipFileName
			
			WScript.Echo "      Adding: " & sFolderName & " To: " & sZipFileName
			WindowsZip sSourceFolder, sZipFile, fso
		End If

		WScript.Echo "    Deleting: " & sFolderName
		On Error Resume Next
		fso.DeleteFolder sSourceFolder
		if err.number <> 0 Then
			WScript.Echo "ERROR: " & Err.Number & vbCRLF & Err.Description
			WScript.Quit err.number
		End If
		
		On Error Goto 0
	End If
End Function

Function ArchiveFile(sFolder, sFileName, fso)
	
	Dim sSourceFile:	sSourceFile		= sFolder & "\" & sFileName
	Dim sZipFileName:	sZipFileName	= GenerateZipFileName(sFolder, sFileName, fso)
	Dim sZipFile:		sZipFile		= sFolder & "\" & sZipFileName
	
	WScript.Echo "      Adding: " & sFileName & " To: " & sZipFileName
	WindowsZip sSourceFile, sZipFile, fso

	WScript.Echo "    Deleting: " & sFileName

	On Error Resume Next
	fso.DeleteFile sSourceFile
	if err.number <> 0 Then
		WScript.Echo "ERROR: " & Err.Number & vbCRLF & Err.Description
		WScript.Quit err.number
	End If
	
	On Error Goto 0

End Function

Function GenerateZipFileName(sFolder, sFileOrFolder2Zip, fso)

	Dim sFileOrFolderPath:	sFileOrFolderPath = sFolder & "\" & sFileOrFolder2Zip
	Dim fFile
	Dim fFolder
	Dim datLastModified
	
	if fso.FileExists(sFileOrFolderPath) Then
		set fFile = fso.GetFile(sFileOrFolderPath)
		datLastModified = fFile.DateLastModified
	ElseIf fso.FolderExists(sFileOrFolderPath) Then
		set fFolder = fso.GetFolder(sFileOrFolderPath)
		datLastModified = fFolder.DateLastModified
	End If
	
	Dim sDay:				sDay    = Day(datLastModified)
	Dim sMonth:				sMonth  = Month(datLastModified)
	Dim sYear:				sYear   = Year(datLastModified)

	If Len(sDay) = 1 Then
		sDay = "0" & sDay
	End If
	
	If Len(sMonth) = 1 Then
		sMonth = "0" & sMonth
	End If

	GenerateZipFileName = sYear & sMonth & sDay & ".zip"
End Function

'
' Zip/Unzip Helper Functions
'
' These are courtesy of stack overflow:
'	http://stackoverflow.com/questions/30211/can-windows-built-in-zip-compression-be-scripted
'
Function WindowsUnZip(sUnzipFileName, sUnzipDestination, fsoUnzip)
	
	If Not fsoUnzip.FolderExists(sUnzipDestination) Then
		fsoUnzip.CreateFolder(sUnzipDestination)
	End If

	With CreateObject("Shell.Application")
       .NameSpace(sUnzipDestination).Copyhere .NameSpace(sUnzipFileName).Items
	End With

End Function

Function WindowsZip(sFile, sZipFile, fsoZip)

	If Not fsoZip.FileExists(sZipFile) Then
		NewZip sZipFile, fsoZip
	End If

	Dim appZip:			Set appZip = CreateObject("Shell.Application")
	Dim sZipFileCount:	sZipFileCount = appZip.NameSpace(sZipFile).items.Count

	Dim arrFile:		arrFile   = Split(sFile, "\")
	Dim sFileName:		sFileName = (arrFile(Ubound(arrFile)))

	'listfiles
	Dim sDupe:			sDupe = False
	Dim sFileNameInZip

	For Each sFileNameInZip In appZip.NameSpace(sZipFile).items
		If LCase(sFileName) = LCase(sFileNameInZip) Then
			sDupe = True
			Exit For
		End If
	Next

	If Not sDupe Then
		appZip.NameSpace(sZipFile).Copyhere sFile

		'Keep script waiting until Compressing is done
		On Error Resume Next
		Do Until sZipFileCount < appZip.NameSpace(sZipFile).Items.Count
			Wscript.Sleep(300)
		Loop
		On Error GoTo 0
	End If
End Function

Sub NewZip(sNewZip, fsoNewZip)
	
	Dim fsoNewZipFile:	Set fsoNewZipFile = fsoNewZip.CreateTextFile(sNewZip)
	fsoNewZipFile.Write "PK" & Chr(5) & Chr(6) & String(18, 0)
	fsoNewZipFile.Close

	Wscript.Sleep(500)
End Sub
