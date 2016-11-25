'
' String module
'
' Author: Jardel Weyrich <jardel@teltecsolutions.com.br>
'

Option Explicit

Function IsNullOrEmptyStr(ByVal str)
	Select Case VarType(str)
		' vbEmpty			0		Empty (uninitialized)
		Case vbEmpty
			IsNullOrEmptyStr = True
		' vbNull			1		Null (no valid data)
		Case vbNull
			IsNullOrEmptyStr = True
		' vbString			8		String
		Case vbString
			IsNullOrEmptyStr = (Len(str) = 0)
		Case Else
			' VBScript Run-time Errors - https://msdn.microsoft.com/en-us/library/xe43cc8d.aspx
			Call Err.Raise(5000, "string.vbs:IsNullOrEmptyStr", "argument is not of type vbEmpty, vbNull, or vbString")
	End Select
End Function

Function StartsWith(ByVal str, ByVal prefix)
	StartsWith = Left(str, Len(prefix)) = prefix
End Function

Function EndsWith(ByVal str, ByVal prefix)
	EndsWith = Right(str, Len(prefix)) = prefix
End Function

Function ReplaceRange(original, first, length, replacement)
	ReplaceRange = Left(original, first) & replacement & Mid(original, first + length + 1)
End Function
