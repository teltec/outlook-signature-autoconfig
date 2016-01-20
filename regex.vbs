'
' Regex module
'
' Author: Jardel Weyrich
'

Option Explicit

Function RegexApplySingleLine(ByVal singleLineInput, ByVal pattern)
	' To debug the RegEx you can use https://myregextester.com/
	Dim re
	Set re = new RegExp
	re.Global = True
	re.Pattern = pattern
	
	Dim matches
	Set matches = re.Execute(singleLineInput)
	'LogDebug("matches.Count=" & matches.Count & " matches(0)=" & matches(0))

	Dim result: result = Array()
	
	If matches.Count > 0 Then
		Dim match
		Set match = matches(0)
		'LogDebug("match.SubMatches.Count=" & match.SubMatches.Count)
		
		ReDim result(match.SubMatches.Count)
		
		Dim i
		For i = 0 to match.SubMatches.Count - 1
			result(i) = match.SubMatches(i)
		Next
	End If
	
	RegexApplySingleLine = result
End Function

' Match the `value` against the `regexPattern`, and if there's a match, return the capture group informed by `captureGroupIndex`.
' Otherwise, return Null.
Function RegexCaptureSingleLine(ByVal value, ByVal regexPattern, ByVal captureGroupIndex)
	Dim captureGroupArray: captureGroupArray = RegexApplySingleLine(value, regexPattern)
	If UBound(captureGroupArray) > captureGroupIndex Then ' Found a match and has at least `captureGroupIndex` capture groups.
		RegexCaptureSingleLine = captureGroupArray(captureGroupIndex) ' Get the informed capture group by index
		Exit Function
	End If
	RegexCaptureSingleLine = Null
End Function
