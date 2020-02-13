angobun = InputBox("請輸入加密後的資料")
hirabun = ""
For kagi = 1 To 9
    hirabun = hirabun & "加密鍵" & CStr(kagi) & ":"
    For i = 1 To Len(angobun)
        moji = Mid(angobun, i, 1)
        hirabun = hirabun & Chr(Asc(moji) Xor kagi)
    Next
    hirabun = hirabun & Chr(&HD)
Next
MsgBox hirabun
