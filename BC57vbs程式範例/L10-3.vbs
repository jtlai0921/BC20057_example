k = InputBox("�п�J�[�K��")
kagi = CInt(k)
bun1 = InputBox("�п�J����αK��")
bun2 = ""
For i = 1 To Len(bun1)
    moji = Mid(bun1, i, 1)
    bun2 = bun2 & Chr(Asc(moji) Xor kagi)
Next
MsgBox bun2
