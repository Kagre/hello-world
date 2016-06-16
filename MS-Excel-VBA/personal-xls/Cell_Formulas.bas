Public Function IsFunction(ref As Range) As Boolean
    IsFunction = (Left(ref.Cells.Formula, 1) = "=")
End Function

Public Function JoinRange(ToJoin As Range, Optional Delimiter As String = ",") As String
    Dim c As Range, start As Boolean
    start = True
    For Each c In ToJoin.Cells
    If Not IsEmpty(c) Then
    If c.Value <> "" Then
        JoinRange = IIf(start, "", JoinRange & Delimiter) & c.Value
        start = False
    End If: End If
    Next c
End Function

Public Function bKey(ByVal K1 As Range, ByVal K2 As Range, Optional ByVal sep As String = ":") As String
    bKey = UCase(Trim(K1.Value)) & sep & UCase(Trim(K2.Value))
End Function

Public Function xXOR(a As Boolean, B As Boolean) As Boolean: xXOR = a Xor B: End Function

Public Function PowerMod(ByVal base As Long, ByVal exponent As Long, ByVal modulus As Long) As Long
    PowerMod = 1
    base = base Mod modulus
    Do While exponent > 0
        If exponent Mod 2 = 1 Then PowerMod = (PowerMod * base) Mod modulus
        exponent = exponent \ 2
        base = (base * base) Mod modulus
    Loop
End Function

Public Function LogMod(ByVal base As Long, ByVal target As Long, ByVal modulus As Long) As Long
    Dim val As Long: val = 1
    Dim exponent As Long: exponent = 1
    base = base Mod modulus
    val = base
    Do While exponent <= modulus And val <> target
'        Debug.Print val & " --> " & (val * base) Mod modulus
        val = (val * base) Mod modulus
        If val = base Then
            LogMod = -1
            Exit Function
        End If
        exponent = exponent + 1
    Loop
    LogMod = IIf(exponent <= modulus, exponent, -1)
End Function

Public Function FindA(ByVal n As Long) As Long
    For FindA = 2 To n
        If PowerMod(FindA, n - 1, n) = 1 Then Exit For
    Next FindA
    If PowerMod(FindA, n - 1, n) <> 1 Then FindA = -1
End Function
