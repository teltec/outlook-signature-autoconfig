'
' Log module
'
' Author: Jardel Weyrich
'

Option Explicit

Function IsRunningOnConsole()
	Dim argv0: argv0 = LCase(WScript.FullName) ' Most likely returns "c:\windows\system32\cscript.exe"
	IsRunningOnConsole = EndsWith(argv0, "\cscript.exe")
End Function

Sub Log(ByVal streamType, ByVal message)
	Const StdOut = 1
	Const StdErr = 2
	If streamType <> StdOut And streamType <> StdErr Then
		' VBScript Run-time Errors - https://msdn.microsoft.com/en-us/library/xe43cc8d.aspx
		Call Err.Raise(5000, "log.vbs:Log", "streamType argument must be one of: StdOut, StdErr")
	End If
	
	Dim fso, stream
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set stream = fso.GetStandardStream(streamType)
	
	If IsRunningOnConsole() Then
		stream.WriteLine message
	Else
		' TODO(jweyrich): Write to a log file?
	End If
	
	Set stream = Nothing
	Set fso = Nothing
End Sub

Sub LogDebug(ByVal message)
	Const StdOut = 1
	Log StdOut, "DEBUG: " & message
End Sub

Sub LogInfo(ByVal message)
	Const StdOut = 1
	Log StdOut, "INFO: " & message
End Sub

Sub LogError(ByVal message)
	Const StdErr = 2
	Log StdErr, "ERROR: " & message
End Sub
