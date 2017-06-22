'
' Template module
'
' Author: Jardel Weyrich <jardel@teltecsolutions.com.br>
'

Option Explicit

'----------------------------------------------------------------------------------------

Class TemplateTag
	Public TagName
	Public Offset
	Public Length

	Sub Class_Initialize()
	End Sub

	Sub Class_Terminate()
	End Sub
	
	Function Init(tagName, offset, length)
		Me.TagName = tagName
		Me.Offset = offset
		Me.Length = length
		Set Init = Me
	End Function
End Class

'----------------------------------------------------------------------------------------

Class IfTag
	Public Tag
	Public VariableName

	Sub Class_Initialize()
		Me.VariableName = vbEmpty
	End Sub

	Sub Class_Terminate()
		Set Me.Tag = Nothing
	End Sub
	
	Function Init(offset, length, variableName)
		Set Me.Tag = New TemplateTag.Init("if", offset, length)
		Me.VariableName = variableName
		Set Init = Me
	End Function
End Class

'----------------------------------------------------------------------------------------

Class EndIfTag
	Public Tag
	
	Sub Class_Initialize()
	End Sub

	Sub Class_Terminate()
		Set Me.Tag = Nothing
	End Sub
	
	Function Init(offset, length)
		Set Me.Tag = new TemplateTag.Init("endif", offset, length)
		Set Init = Me
	End Function
End Class

'----------------------------------------------------------------------------------------

Class ConditionBlock
	Public Context
	Public OpenTag
	Public CloseTag
	
	Sub Class_Initialize()
	End Sub

	Sub Class_Terminate()
		Set Me.Context = Nothing
		Set Me.OpenTag = Nothing
		Set Me.CloseTag = Nothing
	End Sub
	
	' context - Dictionary
	' openTag - IfTag
	' CloseTag - EndIfTag
	Function Init(ByRef context, ByRef openTag)
		Set Me.Context = context
		Set Me.OpenTag = openTag
		Set Init = Me
	End Function
	
	Function Close(ByRef closeTag)
		Set Me.CloseTag = closeTag
		Set Close = Me
	End Function
	
	Function Evaluate()
		Dim variableName: variableName = Me.OpenTag.VariableName
		Dim hasVariableName: hasVariableName = VarType(variableName) <> vbEmpty
		
		If Not hasVariableName Then
			Evaluate = False
			Exit Function
		End If
		
		Dim variableValue: variableValue = Me.Context.Item(variableName)
		If VarType(variableValue) = vbEmpty Then
			Call Err.Raise(5000, "ConditionBlock.Evaluate", "Undefined variable: " + variableName)
		End If
		
		'LogDebug("variableValue=" + CStr(variableValue))
		Evaluate = CBool(variableValue)	
	End Function
End Class

'----------------------------------------------------------------------------------------

' Arguments:
'   templateStr - String containing the template data.
' 	contextDict - Dictionary containing all variables to be used during the template processing.
' Return:
'	Processed template
Function ProcessTemplateTags(ByRef templateStr, ByRef contextDict)
	Const vbTextCompare = 1
	
	If IsNullOrEmptyStr(templateStr) Then
		ProcessTemplateTags = ""
		Exit Function
	End If
	
	Dim partialTemplateStr: partialTemplateStr = ""
	Dim rendered: rendered = ""
	
	'
	' Find tags
	'
	Dim templateSize: templateSize = Len(templateStr)
	Dim currentPosition: currentPosition = 1
	Dim insideConditionBlock: insideConditionBlock = False
	Dim conditionBlock: Set conditionBlock = Nothing
	Dim conditionResult: conditionResult = False
	
	Do While currentPosition < templateSize
		Dim startPosition: startPosition = InStr(currentPosition, templateStr, "{%", vbTextCompare)
		If startPosition = 0 Then
			' No more template tags to process.
			partialTemplateStr = Mid(templateStr, currentPosition, templateSize - currentPosition)
			'LogDebug("partialTemplateStr[no_more_tags] = " + partialTemplateStr)
			rendered = rendered + partialTemplateStr
			Exit Do
		End If
		
		Dim endPosition: endPosition = InStr(startPosition, templateStr, "%}", vbTextCompare)
		If endPosition = 0 Then
			LogError("Missing '%}'" & vbCrlf)
			Exit Do
		End If
		
		partialTemplateStr = Mid(templateStr, currentPosition, startPosition - currentPosition)
		'LogDebug("partialTemplateStr[found_tag] = " + partialTemplateStr)
		
		endPosition = endPosition + 2 ' +2 because of "%}"
		
		currentPosition = endPosition
		
		Dim tagStr: tagStr = Mid(templateStr, startPosition, endPosition - startPosition) 
		'LogDebug("Tag = " + tagStr)
		
		'
		' Parse tag
		'
		' To debug the RegEx you can use https://myregextester.com/
		Dim re
		Set re = new RegExp
		re.Global = True
		re.Multiline = False
		re.Pattern = "{%\s*(?:(.+?)\s+(.+?)?|(.+?))\s*%}" ' keyword <space(s)> argument | keyword
		
		Dim matches
		Set matches = re.Execute(tagStr)
		'LogDebug("matches.Count=" & matches.Count)
		
		Dim match
		For Each match In matches
			Dim matchOffset: matchOffset = startPosition + match.FirstIndex
			Dim matchLength: matchLength = match.Length
			'LogDebug("match.SubMatches.Count=" + CStr(match.SubMatches.Count))
			'Dim submatch
			'For Each submatch In match.SubMatches
			'	LogDebug("SubMatch=" + submatch)
			'Next
			Dim tagName: tagName = IIf(match.SubMatches.Count > 0, LCase(match.SubMatches(0)), vbEmpty)
			'LogDebug("tagName=" + tagName)
			Dim variableName: variableName = IIf(match.SubMatches.Count > 1, match.SubMatches(1), vbEmpty)
			Dim hasVariableName: hasVariableName = VarType(variableName) <> vbEmpty
			'LogDebug("variableName=" + variableName + " hasVariableName=" + CStr(hasVariableName))
			Dim variableValue: variableValue = IIf(hasVariableName, contextDict.Item(variableName), vbEmpty)
			Dim hasVariableValue: hasVariableValue = variableValue <> vbEmpty
			'LogDebug("variableValue=" + CStr(variableValue) + " hasVariableValue=" + CStr(hasVariableValue))
			
			'LogDebug("MATCH:" _
			'	+ " matchOffset=" + CStr(matchOffset) _
			'	+ ", matchLength=" + CStr(matchLength) _
			'	+ ", tagName=" + tagName _
			'	+ ", variableName=" + IIf(hasVariableName, CStr(variableName), "<vbEmpty>") _
			'	+ ", variableValue=" + IIf(hasVariableValue, CStr(variableValue), "<vbEmpty>") _
			')
			
			Dim tmpTag
			
			Select Case tagName
				Case "if"
					If insideConditionBlock Then
						Call Err.Raise(5000, "ProcessTemplateTags", "The processor doesn't supported nested conditions yet.")
					End If
					
					'LogDebug("partialTemplateStr(@if) = " + partialTemplateStr)
					rendered = rendered + partialTemplateStr

					Set tmpTag = New IfTag.Init(matchOffset, matchLength, variableName)
					Set conditionBlock = New ConditionBlock.Init(contextDict, tmpTag)
					conditionResult = conditionBlock.Evaluate()
					insideConditionBlock = True
				Case "endif"
					If Not insideConditionBlock Then
						Call Err.Raise(5000, "ProcessTemplateTags", "Found endif without previous if.")
					End If
					Set tmpTag = New EndIfTag.Init(matchOffset, matchLength)
					conditionBlock.Close(tmpTag)
					insideConditionBlock = False
					If conditionResult Then
						REM LogDebug("Starts: "+Cstr(conditionBlock.OpenTag.Tag.Offset + conditionBlock.OpenTag.Tag.Length))
						REM LogDebug("Length: "+Cstr(conditionBlock.CloseTag.Tag.Offset - conditionBlock.OpenTag.Tag.Offset - conditionBlock.OpenTag.Tag.Length))
						REM partialTemplateStr = Mid(templateStr, _
						REM 	(conditionBlock.OpenTag.Tag.Offset + conditionBlock.OpenTag.Tag.Length), _
						REM 	(conditionBlock.CloseTag.Tag.Offset - (conditionBlock.OpenTag.Tag.Offset - conditionBlock.OpenTag.Tag.Length)) _
						REM )
						'LogDebug("partialTemplateStr(@endif) = " + partialTemplateStr)
						rendered = rendered + partialTemplateStr
					End If
				Case Else
					If insideConditionBlock And conditionResult Then
						' Process tag.
					Else
						' Do NOT process tag.
					End If
			End Select
		Next
	Loop
	
	If insideConditionBlock Then
		Call Err.Raise(5000, "ProcessTemplateTags", "Missing an endif?")
	End If
	
	ProcessTemplateTags = rendered
End Function

'----------------------------------------------------------------------------------------

' Arguments:
'   templateStr - String containing the template data.
' 	contextDict - Dictionary containing all variables to be used during the template processing.
' Return:
'	Processed template
Function BindTemplateVariables(ByRef templateStr, ByRef contextDict)
	' To debug the RegEx you can use https://myregextester.com/
	Dim re
	Set re = new RegExp
	re.Global = True
	re.Multiline = True
	re.Pattern = "{{(.+?)}}"
	
	Dim matches
	Set matches = re.Execute(templateStr)
	'LogDebug("matches.Count=" & matches.Count)

	Dim result: result = Array()
	
	Dim displacement: displacement = 0
	
	Dim match
	For Each match In matches
		Dim matchOffset: matchOffset = match.FirstIndex
		Dim matchLength: matchLength = match.Length
		Dim variableName: variableName = match.SubMatches(0)
		Dim variableValue: variableValue = contextDict.Item(variableName)
		Dim variableValueLength: variableValueLength = Len(variableValue)
		
		'LogDebug("displacement = " + CStr(displacement) _
		'	+ ", matchOffset = " + CStr(matchOffset) _
		'	+ ", matchLength = " + CStr(matchLength) _
		'	+ ", variableName = " + variableName _
		'	+ ", variableValue = " + CStr(variableValue) _
		'	+ ", variableValueLength = " + CStr(variableValueLength) _
		')
		
		templateStr = ReplaceRange(templateStr, matchOffset + displacement, matchLength, variableValue)
		displacement = displacement + (variableValueLength - matchLength)
	Next
	
	BindTemplateVariables = templateStr
End Function

'----------------------------------------------------------------------------------------

Function ReadTemplateFromFile(templateFilePath)
	If Not FileExists(templateFilePath) Then
		LogError("Template file does not exist - " & templateFilePath & vbCrlf)
		Exit Function
	End If
	
	Const ForReading = 1
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFile: Set objFile = objFSO.OpenTextFile(templateFilePath, ForReading)
	Dim templateStr: templateStr = objFile.ReadAll
	objFile.Close
	Set objFile = Nothing
	Set objFSO = Nothing
	
	ReadTemplateFromFile = templateStr
End Function

'----------------------------------------------------------------------------------------

Function ReadTemplateFromFileUTF8(templateFilePath)
	If Not FileExists(templateFilePath) Then
		LogError("Template file does not exist - " & templateFilePath & vbCrlf)
		Exit Function
	End If

	' ADODB.Stream file I/O constants
	Const template_adTypeBinary = 1
	Const template_adTypeText   = 2

	Dim objStream: Set objStream = CreateObject("ADODB.Stream")
	objStream.Open
	objStream.Type = template_adTypeText
	objStream.Position = 0
	' Use UTF-8 so that accents/diacritics actually work.
	objStream.CharSet = "utf-8"
	' Read file
	objStream.LoadFromFile(templateFilePath)
	Dim templateStr: templateStr = objStream.ReadText()
	
	' Close it.
	objStream.Close
	Set objStream = Nothing
	
	ReadTemplateFromFileUTF8 = templateStr
End Function

'----------------------------------------------------------------------------------------

' Arguments:
'   templateStr - String containing the template data.
' 	contextDict - Dictionary containing all variables to be used during the template rendering.
' Return:
'	Rendered template
Function RenderTemplate(ByRef templateStr, ByRef contextDict)
	Dim rendered: rendered = templateStr
	rendered = ProcessTemplateTags(rendered, contextDict)
	rendered = BindTemplateVariables(rendered, contextDict)
	RenderTemplate = rendered
End Function

'----------------------------------------------------------------------------------------
