'
' Configure signatures for Microsoft Outlook 2003/2007/2010/2013/1016.
'
' Author: Jardel Weyrich <jardel@teltecsolutions.com.br> 
' Date  : 20/01/2016 
'
' Parts of this code are originally from http://www.outlookcode.com/codedetail.aspx?id=821
'

Option Explicit

'----------------------------------------------------------------------------------------

Sub Include(ByVal filePath)
	Const ForReading = 1
	Dim currentScriptPath: currentScriptPath = Wscript.ScriptFullName
	Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
	Dim currentScriptFile: Set currentScriptFile = fso.GetFile(currentScriptPath)
	Dim currentScriptDirectory: currentScriptDirectory = fso.GetParentFolderName(currentScriptFile)
	Set currentScriptFile = Nothing
	Dim file: Set file = fso.OpenTextFile(currentScriptDirectory & "\" & filePath & ".vbs", ForReading)
	Dim fileData: fileData = file.ReadAll
	file.Close
	Set file = Nothing
	Set fso = Nothing
	ExecuteGlobal fileData
End Sub

'----------------------------------------------------------------------------------------

Include "app_config"
Include "common"
Include "log"
Include "string"
Include "array"
Include "filesystem"
Include "registry"

'----------------------------------------------------------------------------------------

' Outlook versions
Const gVersionOutlook2003 = "11.0"
Const gVersionOutlook2007 = "12.0"
Const gVersionOutlook2010 = "14.0"
Const gVersionOutlook2013 = "15.0"
Const gVersionOutlook2016 = "16.0"

'----------------------------------------------------------------------------------------

' Change the registry key that specifies in which directory the signatures are stored.
Function ConfigureSignatureLocation(ByVal argDirectory, ByVal outlookVersion)
	Dim regKey: regKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & outlookVersion & "\Common\General\Signatures"
	Dim objShell: Set objShell = CreateObject("WScript.Shell")
	objShell.RegWrite regKey, argDirectory, "REG_SZ"
	Set objShell = Nothing
End Function

'----------------------------------------------------------------------------------------

Function GetLocalSignatureDirectory(ByVal argPastaAssinatura)
	Dim objShell: Set objShell = CreateObject("WScript.Shell")
	
	Dim appData: appData = ObjShell.ExpandEnvironmentStrings("%APPDATA%")
	Dim targetDirectory: targetDirectory = appData & "\Microsoft\" & argPastaAssinatura & "\"
	
	Set objShell = Nothing
	Set appData = Nothing

	GetLocalSignatureDirectory = targetDirectory	
End Function

'----------------------------------------------------------------------------------------

' Create the local directory to store the signature file(s).
' REFERENCE: https://support.microsoft.com/en-us/kb/2691977
Function CreateLocalSignatureDirectory(ByVal argPastaAssinatura)
	Dim targetDirectory: targetDirectory = GetLocalSignatureDirectory(argPastaAssinatura)
	
	Dim objFolder: Set objFolder = CreateObject("Scripting.FileSystemObject")
	
	If Not (objFolder.FolderExists(targetDirectory)) Then
		' LogDebug("Creating directory " & targetDirectory)
		objFolder.CreateFolder(targetDirectory)
	End If

	Set objFolder = Nothing

	CreateLocalSignatureDirectory = targetDirectory	
End Function

'----------------------------------------------------------------------------------------

Function GetProfileRegKeyPath(ByVal emailAddress, ByVal strProfile, ByVal outlookVersion)	
	Dim strKeyPath
	Dim defaultProfileKeyPath
	
	' REFERENCE: https://social.technet.microsoft.com/Forums/office/en-US/f9fd782e-ecdb-41e1-b00a-0b4b2cfd7d32/outlook-signature-registry-settings?forum=outlook
	' Outlook Outlook 2003, 2007, 2010
	If (outlookVersion = gVersionOutlook2003) or (outlookVersion = gVersionOutlook2007) or (outlookVersion = gVersionOutlook2010) Then
		strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\"
		defaultProfileKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\"
	Else ' Outlook 2013, 2016
		strKeyPath = "Software\Microsoft\Office\" & outlookVersion & "\Outlook\Profiles\"
		defaultProfileKeyPath = "Software\Microsoft\Office\" & outlookVersion & "\Outlook\"
	End If
	
	Const HKEY_CURRENT_USER = &H80000001
	
	'LogDebug("strKeyPath = '" & strKeyPath & "'")
	'LogDebug("defaultProfileKeyPath = '" & defaultProfileKeyPath & "'")
	'LogDebug("strProfile(passed) = '" & strProfile & "'")
	
	Dim strComputer: strComputer = "."
	Dim objReg: Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

	' Retrieve the strProfile if it's Null or Empty.
	' TODO(jweyrich): Should we Trim() the string first?
	If IsNullOrEmptyStr(strProfile) Then
		objReg.GetStringValue HKEY_CURRENT_USER, defaultProfileKeyPath, "DefaultProfile", strProfile
	End If

	'LogDebug("strProfile(read) = '" & strProfile & "'")

	' TODO(jweyrich): Should we Trim() the string first?
	If IsNullOrEmptyStr(strProfile) Then
		LogError("Could not find any profile for Outlook " & outlookVersion & " - Aborting.")
		Set objReg = Nothing
		GetProfileRegKeyPath = Null
		Exit Function
	End If
	
	strKeyPath = strKeyPath & strProfile & "\9375CFF0413111d3B88A00104B2A6676"

	'LogDebug("strKeyPath = " & strKeyPath)

	' Retrieve profiles.
	Dim arrProfileKeys
	objReg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrProfileKeys
	
	Dim foundMatchingAccount: foundMatchingAccount = False
	Dim tempKey
	Dim subkey
	For Each subkey In arrProfileKeys	
		tempKey = strKeyPath & "\" & subkey
		'LogDebug("tempKey = " & tempKey)
		
		Dim valueType
		valueType = RegGetValueType(HKEY_CURRENT_USER, tempKey, "Account Name")
		'LogDebug("'Account Name' registry type is " & valueType)
		If valueType = REG_INVALID Then Exit Function
		Dim accountName: accountName = RegGetValue(HKEY_CURRENT_USER, tempKey, "Account Name")
		'LogDebug("'Account Name' = " & accountName)
		
		If accountName = emailAddress Then
			foundMatchingAccount = True
			Exit For
		End If
	Next
		
	If Not foundMatchingAccount Then
		LogError("Could not find a matching account configured for '" & emailAddress & "' on Outlook " & outlookVersion)
		Set objReg = Nothing
		GetProfileRegKeyPath = Null
		Exit Function
	End If
	
	LogInfo("Found a matching account for '" & emailAddress & "' on Outlook " & outlookVersion)

	GetProfileRegKeyPath = tempKey
	
	Set objReg = Nothing
End Function

'----------------------------------------------------------------------------------------

Function GetDefaultSignatureName(ByVal emailAddress, ByVal strProfile, ByVal outlookVersion)
	GetDefaultSignatureName = Null
	
	Dim profileKeyPath: profileKeyPath = GetProfileRegKeyPath(emailAddress, strProfile, outlookVersion)
	If IsNullOrEmptyStr(profileKeyPath) Then Exit Function
	
	' TODO: Test with versions 2003, 2007, 2010
	Dim valueType
	valueType = RegGetValueType(HKEY_CURRENT_USER, profileKeyPath, "New Signature")
	If valueType = REG_INVALID Then Exit Function
	Dim result: result = RegGetValue(HKEY_CURRENT_USER, profileKeyPath, "New Signature")
	
	GetDefaultSignatureName = result
End Function

'----------------------------------------------------------------------------------------

' Configure default signature
' REFERENCE: https://support.microsoft.com/en-us/kb/2691977
Function SetDefaultSignatureName(ByVal emailAddress, ByVal strSigName, ByVal strProfile, ByVal outlookVersion)
	SetDefaultSignatureName = False
	
	Dim profileKeyPath: profileKeyPath = GetProfileRegKeyPath(emailAddress, strProfile, outlookVersion)
	If IsNullOrEmptyStr(profileKeyPath) Then Exit Function
	
	Dim value: value = strSigName
	
	' TODO: Test with versions 2003, 2007, 2010
	
	Dim valueType
	
	valueType = RegGetValueType(HKEY_CURRENT_USER, profileKeyPath, "New Signature")
	If valueType = REG_INVALID Then valueType = REG_BINARY
	Call RegSetValue(HKEY_CURRENT_USER, profileKeyPath, "New Signature", value, valueType)
	
	valueType = RegGetValueType(HKEY_CURRENT_USER, profileKeyPath, "Reply-Forward Signature")
	If valueType = REG_INVALID Then valueType = REG_BINARY
	Call RegSetValue(HKEY_CURRENT_USER, profileKeyPath, "Reply-Forward Signature", value, valueType)
	
	SetDefaultSignatureName = True
End Function

'----------------------------------------------------------------------------------------

' Copy signature files from remote location to local
Function CopyRemoteSignatureFiles(ByRef objUser, ByVal signatureLocalDirectoryName)
	CopyRemoteSignatureFiles = False
	
	' Local directory to store signature files.
	Dim localDirectoryPath: localDirectoryPath = CreateLocalSignatureDirectory(signatureLocalDirectoryName)		
	
	Dim signatureFolder: signatureFolder = gConfigSignaturesSourceLocation + "\" + objUser.samAccountName
	If Not FolderExists(signatureFolder) Then
		LogError("Signature folder '" & signatureFolder & "' does not exist for user '" & objUser.samAccountName & "'")
		Exit Function
	End If
	
	Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
	Dim baseFolder: Set baseFolder = fso.GetFolder(signatureFolder)
	' LogDebug("baseFolder.Path = " & baseFolder.Path)
	
	' Copy all files from remote directory
	Dim file
	For Each file In baseFolder.Files
		Dim remoteSignatureFilePath: remoteSignatureFilePath = baseFolder.Path + "\" + file.Name
		' LogDebug("remoteSignatureFilePath = " & remoteSignatureFilePath)

		Dim localSignatureFilePath: localSignatureFilePath = localDirectoryPath + file.Name
		' LogDebug("localSignatureFilePath = " & localSignatureFilePath)

		' Check if the signature file exists
		If Not fso.FileExists(remoteSignatureFilePath) Then
			LogError("File " & remoteSignatureFilePath & " does not exist... aborting.")
			Set baseFolder = Nothing
			Exit Function
		End If

		' Copy signature file(s) from network share to local directory.
		On Error Resume Next
		fso.CopyFile remoteSignatureFilePath, localSignatureFilePath 
		If Err.Number <> 0 Then
			LogError("Failed to copy file to " & localSignatureFilePath)
			Err.Clear
			Set baseFolder = Nothing
			Set fso = Nothing
			CopyRemoteSignatureFiles = False
			Exit Function
		End If
		On Error Goto 0
		LogInfo("Copied signature to '" & localSignatureFilePath & "'")
	Next
	
	Set baseFolder = Nothing
	Set fso = Nothing
	
	CopyRemoteSignatureFiles = True
End Function

'----------------------------------------------------------------------------------------

' We use the remote signatures location because we want only the recent/valid signatures.
' The local signatures location might be poluted with dozens of old signatures.
Function GetAvailableSignatureNames(objUser)
	Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
	Dim baseFolder: Set baseFolder = fso.GetFolder(gConfigSignaturesSourceLocation + "\" + objUser.samAccountName)
	
	If baseFolder.Files.Count = 0 Then
		' Return nothing! It's important, because we check the result somewhere else.
		Set baseFolder = Nothing
		Set fso = Nothing
		Exit Function
	End If
	
	Dim result: result = Array()
	ReDim result(baseFolder.Files.Count)
	
	Dim i: i = 0
	Dim file
	For Each file In baseFolder.Files
		result(i) = fso.GetBaseName(file)
		'LogDebug("GetAvailableSignatureNames: result[" & i & "] = " & result(i))
		i = i + 1
	Next
	
	Set baseFolder = Nothing
	Set fso = Nothing
	
	GetAvailableSignatureNames = result
End Function

'----------------------------------------------------------------------------------------

Function DecideWhichSignatureToUse(ByVal current, ByRef available())
	DecideWhichSignatureToUse = current
	If ArraySize(available) = 0 Then
		Exit Function
	End If
	
	If IsNullOrEmptyStr(current) Then
		DecideWhichSignatureToUse = available(0)
		Exit Function
	End If

	Dim atIndex: atIndex = ArrayFind(available, current)
	If atIndex < 0 Then
		DecideWhichSignatureToUse = available(0)
	End If
End Function

'----------------------------------------------------------------------------------------

Function ConfigureOutlookDefaultSignature(ByRef objUser, ByVal signatureLocalDirectoryName, ByVal outlookVersion)
	ConfigureOutlookDefaultSignature = False
		
	' Configura Local Assinatura no registro para pasta TeltecSolutions
	Call ConfigureSignatureLocation(signatureLocalDirectoryName, outlookVersion)
	
	Dim currentDefaultSignatureName: currentDefaultSignatureName = GetDefaultSignatureName(objUser.mail, "", outlookVersion)
	LogInfo("The current default signature is '" & currentDefaultSignatureName & "'")
	
	Dim availableSignaturesNames: availableSignaturesNames = GetAvailableSignatureNames(objUser)
	If ArraySize(availableSignaturesNames) = 0 Then
		LogInfo("There are no available signatures for '" & objUser.mail & "'")
	End If
	
	Dim newDefaultSignatureName: newDefaultSignatureName = DecideWhichSignatureToUse(currentDefaultSignatureName, availableSignaturesNames)
	LogInfo("Decided to use '" & newDefaultSignatureName & "' as default signature")
	
	REM Dim arrEmailAccountDomain: arrEmailAccountDomain = Split(objUser.mail, "@")
	REM Dim strMailDomain: strMailDomain = arrEmailAccountDomain(1)
	REM Select Case strMailDomain
	REM 	Case "teltecsolutions.com.br"
	REM 		Call SetDefaultSignatureName(objUser.mail, newDefaultSignatureName, "", outlookVersion)
	REM 	Case "teltecnetworks.com.br"
	REM 		Call SetDefaultSignatureName(objUser.mail, newDefaultSignatureName, "", outlookVersion)
	REM End Select
	
	Call SetDefaultSignatureName(objUser.mail, newDefaultSignatureName, "", outlookVersion)
	
	LogInfo("Changed the current default signature to '" & newDefaultSignatureName & "'")
	
	ConfigureOutlookDefaultSignature = True
End Function

'----------------------------------------------------------------------------------------

Sub Main()
	Dim objSysInfo
	Set objSysInfo = CreateObject("ADSystemInfo")
	Dim strUser: strUser = objSysInfo.UserName
	Set objSysInfo = Nothing

	Dim objUser
	Set objUser = GetObject("LDAP://" & strUser)
	
	Dim signatureLocalDirectoryName: signatureLocalDirectoryName = "TeltecSolutions"
	
	Dim copiedSignatures: copiedSignatures = CopyRemoteSignatureFiles(objUser, signatureLocalDirectoryName)
	If Not copiedSignatures Then
		Exit Sub
	End If
	
	Call ConfigureOutlookDefaultSignature(objUser, signatureLocalDirectoryName, gVersionOutlook2003)
	Call ConfigureOutlookDefaultSignature(objUser, signatureLocalDirectoryName, gVersionOutlook2007)
	Call ConfigureOutlookDefaultSignature(objUser, signatureLocalDirectoryName, gVersionOutlook2010)
	Call ConfigureOutlookDefaultSignature(objUser, signatureLocalDirectoryName, gVersionOutlook2013)
	Call ConfigureOutlookDefaultSignature(objUser, signatureLocalDirectoryName, gVersionOutlook2016)

	' LogInfo("Done!")
End Sub

'----------------------------------------------------------------------------------------

Main()