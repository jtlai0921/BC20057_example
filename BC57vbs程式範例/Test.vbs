' �N��ܥX�⪺�����ܼƪ�l�ơC
Dim te(2)
te(0) = "���Y"
te(1) = "�ŤM"
te(2) = "��"

' �N�έp�ϥΪ̳ӳ��ƪ��ܼƪ�l�ơC

' �N�üƪ��ؤl��l��

' ��ܱҰʰT���C
MsgBox "�q���C�� Ver.1.00 by H.Yazawa"

' ����5���C
For i = 1 To 5
	' ��J�ϥΪ̭n�X�����C
	user = CInt(InputBox("0:���Y�B1:�ŤM�B2:��"))

	'�ϥζüƨӨM�w�q���X�����C
	computer = CInt(Rnd * 2)

	'�s�@�r�����ܨϥΪ̥X�����C
	If user = computer Then
		MsgBox s & "�K�K����!"
	ElseIf computer = (user + 1) Mod 3 Then
		MsgBox s & "�K�K�ϥΪ̳ӧQ!"
		kachi = kachi + 1
	Else
		MsgBox s & "�K�K�q���ӧQ!"
	End If
Next

'��ܨϥΪ̪��ӳ��ơC
MsgBox " �ϥΪ̪��ӳ���:" & kachi
