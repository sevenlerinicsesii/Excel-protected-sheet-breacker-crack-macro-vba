Sub PasswordBreaker()

'if you know the lenght of pass you can only chose and try the number of car

    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    On Error Resume Next
    'if pass 1 car -> max 45sec (with lenovo 8gb ram) if you have better computer you can even solv with better delay)
    For i = 32 To 126
        ActiveSheet.Unprotect Chr(i)
        If ActiveSheet.ProtectContents = False Then
            MsgBox "One usable password is " & Chr(i)
            Exit Sub
        End If
    Next
    'if pass 2 car -> max 30min
    For i = 32 To 126
        For j = 32 To 126
            ActiveSheet.Unprotect Chr(i) & Chr(j)
            If ActiveSheet.ProtectContents = False Then
                MsgBox "One usable password is " & Chr(i) & Chr(j)
                Exit Sub
            End If
        Next
    Next
    'if pass 3 car -> max 25hours
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k)
                If ActiveSheet.ProtectContents = False Then
                    MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k)
                    Exit Sub
                End If
            Next
        Next
    Next
    'if pass 4 car -> max 48days
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For l = 32 To 126
                    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l)
                    If ActiveSheet.ProtectContents = False Then
                        MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l)
                        Exit Sub
                    End If
                Next
            Next
        Next
    Next
    'if pass 5 car -> max 6years
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For l = 32 To 126
                    For m = 32 To 126
                        ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m)
                        If ActiveSheet.ProtectContents = False Then
                            MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m)
                            Exit Sub
                        End If
                    Next
                Next
            Next
        Next
    Next
    'if pass 6 car -> max 274years
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For l = 32 To 126
                    For m = 32 To 126
                        For i1 = 32 To 126
                            ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1)
                            If ActiveSheet.ProtectContents = False Then
                                MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1)
                                Exit Sub
                            End If
                        Next
                    Next
                Next
            Next
        Next
    Next
    'if pass 7 car -> max 12648years
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For l = 32 To 126
                    For m = 32 To 126
                        For i1 = 32 To 126
                            For i2 = 32 To 126
                                ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2)
                                If ActiveSheet.ProtectContents = False Then
                                    MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2)
                                    Exit Sub
                                End If
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
    'if pass 8 car -> max 581 833 years
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For l = 32 To 126
                    For m = 32 To 126
                        For i1 = 32 To 126
                            For i2 = 32 To 126
                                For i3 = 32 To 126
                                    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3)
                                    If ActiveSheet.ProtectContents = False Then
                                        MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3)
                                        Exit Sub
                                    End If
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
    'if pass 9 car -> max 26.7 millions years
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For l = 32 To 126
                    For m = 32 To 126
                        For i1 = 32 To 126
                            For i2 = 32 To 126
                                For i3 = 32 To 126
                                    For i4 = 32 To 126
                                        ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4)
                                        If ActiveSheet.ProtectContents = False Then
                                            MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4)
                                            Exit Sub
                                        End If
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
    'if pass 10 car -> max 1.23 billion years
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For l = 32 To 126
                    For m = 32 To 126
                        For i1 = 32 To 126
                            For i2 = 32 To 126
                                For i3 = 32 To 126
                                    For i4 = 32 To 126
                                        For i5 = 32 To 126
                                            ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5)
                                            If ActiveSheet.ProtectContents = False Then
                                                MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5)
                                                Exit Sub
                                            End If
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
    'if pass 11 car -> max 56.6 billions years
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For l = 32 To 126
                    For m = 32 To 126
                        For i1 = 32 To 126
                            For i2 = 32 To 126
                                For i3 = 32 To 126
                                    For i4 = 32 To 126
                                        For i5 = 32 To 126
                                            For i6 = 32 To 126
                                                ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6)
                                                If ActiveSheet.ProtectContents = False Then
                                                    MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6)
                                                    Exit Sub
                                                End If
                                            Next
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
    'if pass 12 car -> max 2605 billions years
    For i = 32 To 126
        For j = 32 To 126
            For k = 32 To 126
                For l = 32 To 126
                    For m = 32 To 126
                        For i1 = 32 To 126
                            For i2 = 32 To 126
                                For i3 = 32 To 126
                                    For i4 = 32 To 126
                                        For i5 = 32 To 126
                                            For i6 = 32 To 126
                                                For n = 32 To 126
                                                    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                    If ActiveSheet.ProtectContents = False Then
                                                        MsgBox "One usable password is " & Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                        Exit Sub
                                                    End If
                                                Next
                                            Next
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next

End Sub

