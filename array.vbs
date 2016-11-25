'
' Array module
'
' Author: Jardel Weyrich <jardel@teltecsolutions.com.br>
'

Option Explicit

Function IsArray(ByRef value)
	If ((VarType(value) And vbArray) = vbArray) Then
		IsArray = True
	Else
		IsArray = False
	End If
End Function

Function IsArrayOf(ByRef value, ByVal valueType)
	If ((VarType(value) And vbArray) = vbArray) And ((VarType(value) And valueType) = valueType) Then
		IsArrayOf = True
	Else
		IsArrayOf = False
	End If
End Function

Function ArraySize(ByRef array())
	ArraySize = 0

	If VarType(array) = vbNull Or VarType(array) = vbEmpty Then
		Exit Function
	End If
	
	If Not IsArray(array) Then
		LogDebug("ArraySize: the parameter is not an array (array=" & array & ")")
		Exit Function
	End If
	
	ArraySize = UBound(array) + 1
End Function

Function ArrayCountNotEmptyOrNull(ByRef array())
	Dim i, count: count = 0
	For i = LBound(array) To UBound(array)
		If Not (IsEmpty(array(i)) Or IsNull(array(i))) Then
			count = count + 1
		End If
	Next
	
	ArrayCountNotEmptyOrNull = count
End Function

Function ArrayFind(ByRef array(), ByRef value)
	ArrayFind = -1
	Dim found: found = False
	Dim i
	For i = LBound(array) To UBound(array)
		If array(i) = value Then
			found = True
			ArrayFind = i
			Exit For
		End If
	Next
End Function

' Code from http://www.4guysfromrolla.com/webtech/032800-1.shtml#postadlink
Class DynamicArray
	Private data_
	Private size_
	Private used_

	Private Sub Class_Initialize()
		ReDim data_(7) ' Valid indices ragen from 0 to 7
		size_ = 8
		used_ = 0
	End Sub
	
	Public Property Get HasItems()
		HasItems = used_ > 0
	End Property
	
	Public Property Get Size()
		Size = size_
	End Property
	
	Public Property Get Used()
		Used = used_
	End Property

	Public Property Get Items()
		Items = data_
	End Property
	
	Public Property Get StartIndex()
		StartIndex = 0
	End Property

	Public Property Get EndIndex()
		EndIndex = used_
	End Property
	
	Public Property Get Item(ByVal index)
		If index < 0 Then Call Err.Raise(5000, "DynamicArray.(Get Item)", "Index out of range (< 0)")
		If index >= used_ Then Call Err.Raise(5000, "DynamicArray.(Get Item)", "Index out of range (>= used_)")

		Item = data_(index)
	End Property

	Public Property Let Item(ByVal index, ByRef value)
		If index < 0 Then Call Err.Raise(5000, "DynamicArray.(Let Item)", "Index out of range (< 0)")
		If index >= size_ Then Call Err.Raise(5000, "DynamicArray.(Let Item)", "Index out of range (>= size_)")

		data_(index) = value
	End Property
	
	Public Sub Resize(ByVal newSize)
		If newSize = size_ Then Exit Sub ' Same size
		
		If (newSize > size_) Then ' Increasing size
			Redim Preserve data_(newSize)
			size_ = newSize
		Else ' Decreasing size
			If newSize < used_ Then
				Call Err.Raise(5000, "DynamicArray.Resize", "newSize can't be smaller than used_")
			End If
			
			Redim Preserve data_(newSize)
			size_ = newSize
		End If
	End Sub
	
	' Add to the end
	Public Sub Add(ByRef value)
		If size_ <= used_ Then
			Resize(size_ * 2) ' Amortized resize
		End If
		
		data_(used_) = value
		used_ = used_ + 1
	End Sub
	
	' Remove from the end
	Public Sub Remove()
		If index < 0 Then Call Err.Raise(5000, "DynamicArray.Remove", "Index out of range (< 0)")
		If index >= used_ Then Call Err.Raise(5000, "DynamicArray.Remove", "Index out of range (>= used_)")	
		
		data_(used_ - 1) = Empty
	End Sub
End Class