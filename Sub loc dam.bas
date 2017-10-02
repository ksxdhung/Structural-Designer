Private Sub CommandButton1_Click()
i = 5
While Sheet1.Cells(i, "D") <> ""
i = i + 1
Wend

j = 6
While Sheet1.Cells(j, "M") <> ""
j = j + 1
Wend

For i = 4 To i - 1
    For j = 6 To j - 1
    If Sheet1.Cells(i, "D") = Sheet1.Cells(j, "M") Then
        Sheet1.Cells(i, "F") = Sheet1.Cells(j, "N")
        Sheet1.Cells(i, "G") = Sheet1.Cells(j, "O")
    End If
    
    If Sheet1.Cells(i, "E") = Sheet1.Cells(j, "M") Then
        Sheet1.Cells(i, "H") = Sheet1.Cells(j, "N")
        Sheet1.Cells(i, "I") = Sheet1.Cells(j, "O")
    End If
    
    Next j
    
Next i

End Sub

Private Sub CommandButton2_Click()
i = 5
k = 4
Dim arr As Variant
While Sheet1.Cells(i, "C") <> ""
    If Sheet1.Cells(i, "K") <> "" Then
        arr = Split(Sheet1.Cells(i, "K"), "+", , vbTextCompare)
        For Each dam In arr
        Sheet2.Cells(k, 2) = dam
        k = k + 1
        Next
    End If
Sheet2.Range("B" & k - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
i = i + 1
Wend
Sheet2.Activate
End Sub

Private Sub CommandButton4_Click()
'xac dinh phuong cua dam
i = 5
While Sheet1.Cells(i, "D") <> ""
    If Sheet1.Cells(i, 6) - Sheet1.Cells(i, 8) = 0 Then
        Sheet1.Cells(i, 10) = "Phuong Y"
    Else
        If Sheet1.Cells(i, 7) - Sheet1.Cells(i, 9) = 0 Then
        Sheet1.Cells(i, 10) = "Phuong X"
        Else
        Sheet1.Cells(i, 10) = "Phuong Xien"
        End If
    End If
i = i + 1
Wend

' Noi 2 dam lien tuc
For j = 4 To i - 1
    s = Sheet1.Cells(j, 3)
    For k = 4 To i - 1
        If Sheet1.Cells(j, 5) = Sheet1.Cells(k, 4) And Sheet1.Cells(j, 10) = Sheet1.Cells(k, 10) And k <> j Then s = s & "+" & Sheet1.Cells(k, 3)
    Next k
    
    Sheet1.Cells(j, 11) = s
Next j

'noi cac dam lien tuc tu cac doan 2 dam
Dim s1, s2 As String
For h = 4 To i - 1
    s1 = CStr(Sheet1.Cells(h, 11))
    For k = 4 To i - 1
        s2 = CStr(Sheet1.Cells(k, 11))
        If Text1(s1) = Text2(s2) And k <> h And Text3(Text1(s1), s2) <> "" Then
            Sheet1.Cells(k, 11) = ""
            s1 = s1 & Text3(Text1(s1), s2)
        End If
        Sheet1.Cells(h, 11) = s1
    Next k
Next h
    
For h = 4 To i - 1
    s1 = CStr(Sheet1.Cells(h, 11))
    For k = 4 To i - 1
        s2 = CStr(Sheet1.Cells(k, 11))
        If Text1(s1) = Text2(s2) And k <> h And Text3(Text1(s1), s2) <> "" Then
            Sheet1.Cells(k, 11) = ""
            s1 = s1 & Text3(Text1(s1), s2)
        End If
        Sheet1.Cells(h, 11) = s1
    Next k
Next h

'xoa bo cac doan dam thua
For h = 4 To i - 1
    If Sheet1.Cells(h, 11) <> "" Then
        s1 = CStr(Sheet1.Cells(h, 11))
        m = Len(s1)
        p = 0
        For k = 4 To i - 1
            s2 = CStr(Sheet1.Cells(k, 11))
            For n = 1 To Len(s2)
                If Text1(s2) = s1 And h <> k Then p = p + 1
            Next n
        Next k
        If p > 0 Then Sheet1.Cells(h, 11) = ""
    End If
Next h
MsgBox ("Join successfully!")
End Sub
