angobun = InputBox("�п�J�[�K�᪺���")
hirabun = ""
For i = 1 To Len(angobun)
    moji = Mid(angobun, i, 1)
    hirabun = hirabun & Chr(Asc(moji) - 3)
Next
MsgBox hirabun
