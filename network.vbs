'
' Network module
'
' Author: Jardel Weyrich <jardel@teltecsolutions.com.br>
'

Option Explicit

Function GetCurrentUsername()
	Dim objNetwork
	Set objNetwork = CreateObject("WScript.Network")
	GetCurrentUsername = objNetwork.UserName
	Set objNetwork = Nothing
End Function

Function GetCurrentDomainName()
	Dim objNetwork
	Set objNetwork = CreateObject("WScript.Network")
	GetCurrentDomainName = objNetwork.UserDomain
	Set objNetwork = Nothing
End Function

Function GetUserDN(ByVal un, ByVal dn)
	Dim obj
	Set obj = CreateObject("NameTranslate")
	obj.init 1, dn
	obj.set 3, dn & "\" & un
	GetUserDN = obj.Get(1)
	Set obj = Nothing
End Function

REM ' Retrieve the username from the logged in user.
REM Function GetUsername()
REM 	GetUsername = ""
REM 	Dim strComputer: strComputer = "."
REM 
REM 	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
REM 	Set colItems = objWMIService.ExecQuery("Select * From Win32_ComputerSystem")
REM 
REM 	For Each objItem in colItems
REM 		Dim userNameParts: userNameParts = Split(objItem.UserName, "\")
REM 		GetUsername = userNameParts(1)
REM 	Next
REM End Function