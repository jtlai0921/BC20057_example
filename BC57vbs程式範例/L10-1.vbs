hirabun = InputBox("½Ð¿é¤J¤å¦r¡C")
angobun = ""
For i = 1 To Len(hirabun)
    moji = Mid(hirabun, i, 1)
    angobun = angobun & Chr(Asc(moji) + 3)
Next
MsgBox angobun
