

'1. Ham lay ten dam cuoi cung

Public Function Text1(ByVal s As String)
    For i = 1 To Len(s)
        If Mid(s, i, 1) = "+" Then
        m = i
        End If
    Next i
    Text1 = Right(s, Len(s) - m)
End Function

'2. Ham lay ten dam dau tien
Public Function Text2(ByVal s As String)
    m = 0
    For i = 1 To Len(s)
        If Mid(s, i, 1) = "+" Then
        m = i
        Exit For
        End If
    Next i
    If m = 0 Then Text2 = s Else Text2 = Left(s, i - 1)
End Function

'3. Ham lay phan con lai cua chuoi khac voi chuoi khac

Function Text3(ByVal s1 As String, ByVal s2 As String)
i = Len(s1)
j = Len(s2)
m = 0
For k = 1 To j
    If Mid(s2, k, i) = s1 Then m = k
Next
If m <> 0 Then s3 = Left(s2, m - 1) & Right(s2, j - m - i + 1) Else s3 = ""
Text3 = s3
End Function
