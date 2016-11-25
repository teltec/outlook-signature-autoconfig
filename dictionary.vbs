'
' Dictionary module
'
' Author: Jardel Weyrich <jardel@teltecsolutions.com.br>
'

Option Explicit

Class Dictionary
	Public List

	Sub Class_Initialize()
		Set List = CreateObject("Scripting.Dictionary")
	End Sub

	Sub Class_Terminate()
		Set List = Nothing
	End Sub

	Function Append(key, value) 
		List.Add CStr(key), value 
		Append = value
	End Function

	Function Item(key)
		If List.Exists(CStr(key)) Then
			Item = List(CStr(key))
		Else
			Item = vbEmpty
		End If
	End Function
End Class