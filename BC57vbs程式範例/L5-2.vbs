a = 91
s = "�O���"
For i = 2 to (a - 1)
	If a Mod i = 0 Then
		s = "���O���"
		Exit For
	End If
Next
MsgBox CStr(a) & s
