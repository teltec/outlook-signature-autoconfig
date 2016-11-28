'
' Filesystem module
'
' Author: Jardel Weyrich <jardel@teltecsolutions.com.br>
'

Option Explicit

Function FolderExists(ByVal fullpath)
	Dim ofs
	Set ofs = WScript.CreateObject("Scripting.FileSystemObject")
	FolderExists = ofs.FolderExists(fullpath)
	Set ofs = Nothing
End Function

Function FileExists(ByVal fullpath)
	Dim ofs
	Set ofs = WScript.CreateObject("Scripting.FileSystemObject")
	FileExists = ofs.FileExists(fullpath)
	Set ofs = Nothing
End Function

Function CreateFolderRecursive(ByVal fullpath)
	CreateFolderRecursive = True

	Dim ofs
	Set ofs = WScript.CreateObject("Scripting.FileSystemObject")
	If ofs.FolderExists(fullpath) Then
		Set ofs = Nothing
		Exit Function
	End If
	
	Dim isUncPath: isUncPath = StartsWith(fullpath, "\\")
	
	Dim parts
	Dim path
	
	If isUncPath Then
		parts = Split(Mid(fullpath, 3), "\") ' Skip the initial "\\"
		path = "\\" ' Start the path with "\\"
	Else
		parts = Split(fullpath, "\")
		path = ""
	End If

	Dim dir
	For Each dir In parts
		If path <> "" And path <> "\\" Then path = path & "\"
		If path <> "\\" Then
			path = path & dir
			On Error Resume Next
			If Not ofs.FolderExists(path) Then
				ofs.CreateFolder(path)
			End If
			If Err.Number <> 0 Then
				LogError("Cannot access/create folder " & path)
				Err.Clear
				Set ofs = Nothing
				CreateFolderRecursive = False
				Exit Function
			End If
			On Error Goto 0
		Else
			path = path & dir
		End If
		'LogDebug("path=" & path & " dir=" & dir)
	Next
	
	Set ofs = Nothing
End Function
