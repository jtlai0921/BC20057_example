a = 91
s = "是質數"
For i = 2 to (a - 1)
	If a Mod i = 0 Then
		s = "不是質數"
		Exit For
	End If
Next
MsgBox CStr(a) & s
