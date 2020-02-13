angobun = InputBox("請輸入加密後的資料")
hirabun = ""
For i = 1 To Len(angobun)
    moji = Mid(angobun, i, 1)
    hirabun = hirabun & Chr(Asc(moji) - 3)
Next
MsgBox hirabun
