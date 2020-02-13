' 將表示出手的種類變數初始化。
Dim te(2)
te(0) = "石頭"
te(1) = "剪刀"
te(2) = "布"

' 將統計使用者勝場數的變數初始化。

' 將亂數的種子初始化

' 顯示啟動訊息。
MsgBox "猜拳遊戲 Ver.1.00 by H.Yazawa"

' 比賽5次。
For i = 1 To 5
	' 輸入使用者要出的拳。
	user = CInt(InputBox("0:石頭、1:剪刀、2:布"))

	'使用亂數來決定電腦出的拳。
	computer = CInt(Rnd * 2)

	'製作字串來顯示使用者出的拳。
	If user = computer Then
		MsgBox s & "……平手!"
	ElseIf computer = (user + 1) Mod 3 Then
		MsgBox s & "……使用者勝利!"
		kachi = kachi + 1
	Else
		MsgBox s & "……電腦勝利!"
	End If
Next

'顯示使用者的勝場數。
MsgBox " 使用者的勝場數:" & kachi
