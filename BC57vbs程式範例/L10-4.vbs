angobun = InputBox("�п�J�[�K�᪺���")
hirabun = ""
For kagi = 1 To 9
    hirabun = hirabun & "�[�K��" & CStr(kagi) & ":"
    For i = 1 To Len(angobun)
        moji = Mid(angobun, i, 1)
        hirabun = hirabun & Chr(Asc(moji) Xor kagi)
    Next
    hirabun = hirabun & Chr(&HD)
Next
MsgBox hirabun
