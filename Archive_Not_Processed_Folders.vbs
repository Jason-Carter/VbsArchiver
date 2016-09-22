option explicit

Dim arrFolders
Dim iFolder
Dim sFolder
Dim fFolder
Dim objFSO
Dim fFile

Set objFSO = CreateObject("Scripting.FileSystemObject")

'
' Use underscore/carriage return to create multiple folders to process over several lines, e.g.
'
' arrFolders = Array(	"\\<servername>\<sharename>$\<pathname1>\<pathname2>\<pathnameX>\NotProcessed", _
'			"\\<servername>\<sharename>$\<pathname1>\<pathname2>\<pathnameX>\NotProcessed")
'
arrFolders = Array("\\<servername>\<sharename>$\<pathname1>\<pathname2>\<pathnameX>\NotProcessed")

for iFolder = LBound(arrFolders) to UBound(arrFolders)
	
	sFolder = arrFolders(iFolder)
	WScript.Echo "Archiving files in folder: " & sFolder
	set fFolder = objFSO.GetFolder(sFolder)
	
	for each fFile in fFolder.Files
		ArchiveFile sFolder, fFile.Name
	next

	WScript.Echo "Folder files archived: " & sFolder
	WScript.Echo ""
next

'
' Helper Functions
'

Function ArchiveFile(sFolder, sFileName)
	
	Dim sSourceFile
	Dim sArchiveFolder
	Dim sArchiveFile
	Dim sZipFile
	
	if InStr(sFileName, "_TRD_") > 0 then
		
		sSourceFile    = sFolder & "\" & sFileName
		sArchiveFolder = sFolder & "\" & ConvertFileNameToDateBasedName(sFileName)
		sArchiveFile   = sArchiveFolder & "\" & sFileName
		sZipFile       = sArchiveFolder & ".zip"
		
		WScript.Echo "      Adding: " & sSourceFile
		WScript.Echo "          To: " & sZipFile
		
		WindowsZip sSourceFile, sZipFile

		' Once zipped, move this to a folder (TODO: delete instead)
		'if not objFSO.FolderExists(sArchiveFolder) Then
		'	WScript.Echo "    Creating: " & sArchiveFolder
		'	objFSO.CreateFolder(sArchiveFolder)
		'end if
		'
		'WScript.Echo "   Moving To: " & sArchiveFile
		'objFSO.MoveFile sSourceFile, sArchiveFile
		
		WScript.Echo "    Deleting: " & sSourceFile
		objFSO.DeleteFile sSourceFile
	end if
	
End Function

Function ConvertFileNameToDateBasedName(sFileName)
	
	Dim arrFile
	Dim sYear
	Dim sMonth
	Dim sDay
	
	' Example filename:
	'
	'		XXXYYYZZZ_TRD_04_04_2014.txt
	'
	arrFile =  split(sFileName, "_")
	sDay    = arrFile(2)
	sMonth  = arrFile(3)
	sYear   = Left(arrFile(4), 4) ' the .txt part is removed
	
	ConvertFileNameToDateBasedName = sYear & sMonth & sDay

End Function

'
' Zip/Unzip Helper Functions
'
' These are courtesy of stack overflow:
'	http://stackoverflow.com/questions/30211/can-windows-built-in-zip-compression-be-scripted
'
Function WindowsUnZip(sUnzipFileName, sUnzipDestination)
	
	Dim fsoUnzip
	
	Set fsoUnzip = CreateObject("Scripting.FileSystemObject")
 
	If Not fsoUnzip.FolderExists(sUnzipDestination) Then
		fsoUnzip.CreateFolder(sUnzipDestination)
	End If

	With CreateObject("Shell.Application")
       .NameSpace(sUnzipDestination).Copyhere .NameSpace(sUnzipFileName).Items
	End With

	Set fsoUnzip = Nothing
End Function

Function WindowsZip(sFile, sZipFile)

	Dim fsoZip
	Dim appZip
	Dim sZipFileCount
	Dim arrFile
	Dim sFileName
	Dim sDupe
	Dim sFileNameInZip

	Set fsoZip = CreateObject("Scripting.FileSystemObject")

	If Not fsoZip.FileExists(sZipFile) Then
		NewZip(sZipFile)
	End If

	Set appZip = CreateObject("Shell.Application")
	sZipFileCount = appZip.NameSpace(sZipFile).items.Count

	arrFile = Split(sFile, "\")
	sFileName = (arrFile(Ubound(arrFile)))

	'listfiles
	sDupe = False
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

Sub NewZip(sNewZip)
	
	Dim fsoNewZip
	Dim fsoNewZipFile
	
	Set fsoNewZip  = CreateObject("Scripting.FileSystemObject")

	Set fsoNewZipFile = fsoNewZip.CreateTextFile(sNewZip)
	fsoNewZipFile.Write "PK" & Chr(5) & Chr(6) & String(18, 0)
	fsoNewZipFile.Close

	Set fsoNewZip = Nothing
	Wscript.Sleep(500)
End Sub
