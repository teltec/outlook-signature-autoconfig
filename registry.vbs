'
' Registry module
'
' Author: Jardel Weyrich
'

Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Const REG_INVALID = -1
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7

Function RegGetValueType(ByVal defKey, ByVal subKeyName, ByVal valueName)
	Dim strComputer: strComputer = "."
	Dim objReg: Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	
	Dim arrValueNames()
	Dim arrValueTypes()
	
	objReg.EnumValues defKey, subKeyName, arrValueNames, arrValueTypes

	Dim found: found = False
	If Not IsNull(arrValueNames) Then
		Dim i
		'For i = 0 To UBound(arrValueNames)
		'	LogDebug("arrValueNames[" & i & "] = " & arrValueNames(i))
		'Next
		For i = LBound(arrValueNames) To UBound(arrValueNames)
			If arrValueNames(i) = valueName Then
				found = True
				Exit For
			End If
		Next
	End If
	
	RegGetValueType = IIf(found, arrValueTypes(i), REG_INVALID)
	
	Set objReg = Nothing
End Function

Function RegGetValue(ByVal defKey, ByVal subKeyName, ByVal valueName)
	Dim strComputer: strComputer = "."
	Dim objReg: Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
		
	Dim valueData: valueData = Empty
	Dim valueDataStr: valueDataStr = ""
	Dim valueType: valueType = RegGetValueType(defKey, subKeyName, valueName)
	Select Case valueType
		Case REG_SZ
			objReg.GetStringValue defKey, subKeyName, valueName, valueData
			valueDataStr = valueData
		'Case REG_EXPAND_SZ
		'	objReg.GetExpandedStringValue defKey, subKeyName, valueName, valueData
		Case REG_BINARY
			objReg.GetBinaryValue defKey, subKeyName, valueName, valueData
			valueDataStr = RegBinaryToString(valueData)
		'Case REG_DWORD
		'	objReg.GetDWORDValue defKey, subKeyName, valueName, valueData
		'Case REG_MULTI_SZ
		'	objReg.GetMultiStringValue defKey, subKeyName, valueName, valueData
		Case REG_INVALID
			Call Err.Raise(5000, "registry.vbs:RegGetValue", "Invalid valueType")
		Default
			Call Err.Raise(5000, "registry.vbs:RegGetValue", "Unhandled valueType")
	End Select

	Set objReg = Nothing
	
	RegGetValue = valueDataStr
End Function

Function RegSetValue(ByVal defKey, ByVal subKeyName, ByVal valueName, ByRef valueDataStr, ByVal valueType)
	Dim strComputer: strComputer = "."
	Dim objReg: Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	
	Select Case valueType
		Case REG_SZ
			objReg.SetStringValue defKey, subKeyName, valueName, valueDataStr
		'Case REG_EXPAND_SZ
		'	objReg.SetExpandedStringValue defKey, subKeyName, valueName, valueDataStr
		Case REG_BINARY
			objReg.SetBinaryValue defKey, subKeyName, valueName, RegStringToBinary(valueDataStr)
		'Case REG_DWORD
		'	objReg.SetDWORDValue defKey, subKeyName, valueName, valueDataStr
		'Case REG_MULTI_SZ
		'	objReg.SetMultiStringValue defKey, subKeyName, valueName, valueDataStr
		Case REG_INVALID
			Call Err.Raise(5000, "registry.vbs:RegGetValue", "Invalid valueType")
		Default
			Call Err.Raise(5000, "registry.vbs:RegGetValue", "Unhandled valueType")
	End Select

	Set objReg = Nothing
End Function

Function RegBinaryToString(ByRef value())
	Dim result: result = ""
	
	Dim i
	For i = LBound(value) To UBound(value)
		If value(i) <> 0 Then result = result & Chr(value(i)) 
	Next
	
	RegBinaryToString = result
End Function

Function RegStringToBinary(ByVal value)
	Dim result: result = ""
	Dim length: length = Len(value)
		
	Dim i
	For i = 1 to length
		Dim ascii: ascii = Asc(Mid(value, i, 1))
		result = IIf(i = 1, ascii & ",00", result & "," & ascii & ",00")
	Next
	
	result = result & ",00,00"
	
	RegStringToBinary = Split(result, ",")
End Function

Function RegStringToBinaryStr(ByVal value)
	Dim result: result = ""
	Dim length: length = Len(value)
	
	Dim i
	For i = 1 to length
		result = result & Hex(AscW(Mid(value, i, 1))) & ",00,"
	Next
	
	result = result & "00,00"

	RegStringToBinaryStr = result
End Function
