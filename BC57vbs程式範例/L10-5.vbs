Dim kagi(2)
kagi(0) = 3
kagi(1) = 4
kagi(2) = 5
bun1 = InputBox("�п�J����αK��")
bun2 = ""
For i = 1 To Len(bun1)
    moji = Mid(bun1, i, 1)
    bun2 = bun2 & Chr(Asc(moji) Xor kagi((i - 1) Mod 3))
Next
MsgBox bun2
