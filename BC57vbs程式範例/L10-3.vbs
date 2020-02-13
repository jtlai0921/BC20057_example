k = InputBox("請輸入加密鍵")
kagi = CInt(k)
bun1 = InputBox("請輸入明文或密文")
bun2 = ""
For i = 1 To Len(bun1)
    moji = Mid(bun1, i, 1)
    bun2 = bun2 & Chr(Asc(moji) Xor kagi)
Next
MsgBox bun2
