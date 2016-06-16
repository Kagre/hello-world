'Option Compare Database
Option Explicit

Private Sub ToBase64_tst(): Debug.Print ToBase64("Man123"): End Sub
Public Function ToBase64(ByVal str As String) As String
    Const Mask11 As Long = &HFC
    Const Mask12 As Long = &H3
    Const Mask21 As Long = &HF0
    Const Mask22 As Long = &HF
    Const Mask31 As Long = &HC0
    Const Mask32 As Long = &H3F
    
    Dim Alphabet As Variant
    Alphabet = Array( _
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", _
        "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
        "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", _
        "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", _
        "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/")
    
    Dim A As Long, b As Long, c As Long
    Dim e As Long, F As Long, g As Long, H As Long
    
    Dim i As Long, sel As Long
    sel = 0
    For i = 1 To Len(str)
        Select Case sel
         Case 0
            A = Asc(Mid(str, i, 1))
            e = (A And Mask11) \ 4
            F = (A And Mask12) * 16
            ToBase64 = ToBase64 & Alphabet(e)
            sel = 1
         Case 1
            b = Asc(Mid(str, i, 1))
            F = F Or ((b And Mask21) \ 16)
            g = (b And Mask22) * 4
            ToBase64 = ToBase64 & Alphabet(F)
            sel = 2
         Case 2
            c = Asc(Mid(str, i, 1))
            g = g Or ((c And Mask31) \ 64)
            H = c And Mask32
            ToBase64 = ToBase64 & Alphabet(g) & Alphabet(H)
            sel = 0
        End Select
    Next i
    
    Select Case sel
     Case 1
        ToBase64 = ToBase64 & Alphabet(F)
     Case 2
        ToBase64 = ToBase64 & Alphabet(g)
    End Select
End Function

Private Sub FromBase64_tst(): Debug.Print FromBase64("TWFu"): End Sub
Public Function FromBase64(s As String) As String
    Const Mask21 As Long = &H30
    Const Mask22 As Long = &HF
    Const Mask31 As Long = &H3C
    Const Mask32 As Long = &H3
    
    Dim Alphabet As String
    Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    
    Dim A As Long, b As Long, c As Long
    Dim e As Long, F As Long, g As Long, H As Long
    
    Dim i As Long, at As Long, sel As Long
    sel = 0
    For i = 1 To Len(s)
        at = InStr(Alphabet, Mid(s, i, 1))
        Do Until at <> 0
            i = i + 1
            at = InStr(Alphabet, Mid(s, i, 1))
        Loop
        at = at - 1
        Select Case sel
         Case 0
            e = at
            A = e * 4
            sel = 1
         Case 1
            F = at
            A = A Or ((F And Mask21) \ 16)
            b = (F And Mask22) * 16
            FromBase64 = FromBase64 & Chr(A)
            sel = 2
         Case 2
            g = at
            b = b Or ((g And Mask31) \ 4)
            c = (g And Mask32) * 64
            FromBase64 = FromBase64 & Chr(b)
            sel = 3
         Case 3
            H = at
            c = c Or H
            FromBase64 = FromBase64 & Chr(c)
            sel = 0
        End Select
    Next i
    
    Select Case sel
     Case 1
        FromBase64 = FromBase64 & Chr(A)
     Case 2
        FromBase64 = FromBase64 & Chr(b)
     Case 3
        FromBase64 = FromBase64 & Chr(c)
    End Select
End Function
