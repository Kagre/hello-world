Option Private Module

Private Function Quote(ByVal S As String) As String: Quote = """" & S & """": End Function

' Lambda concept derived from: http://www4.ncsu.edu/~cjhazard/projects/spreadsheet_programming.html
'  I've used %1, %2, ... to signify argument replacements because $ could go wrong with Sheet1!$A$1+$1
Private Function Lambda(Funct As String, ParamArray Params()) As String
    Lambda = Funct
    Dim i As Long, j As Long, p As String
    j = UBound(Params) + 1
    For i = UBound(Params) To LBound(Params) Step -1
        p = IIf(VarType(Params(i)) = vbString, Quote(Params(i)), Params(i))
        Lambda = Replace(Lambda, "%" & j, p)
        j = j - 1
    Next i
    Lambda = Replace(Lambda, "Chr(", "Char(")
    
#If False Then
' Depricated code retained incase I need to replace something other than Chr
'  which doesn't have a spreadsheet function equivelent.
    Dim c As String, ch As String, at As Long, pCount As Long
    Dim lft As String, rht As String
    at = InStr(UCase(Lambda), "CHR(")
    Do While at > 0
        lft = Left(Lambda, at - 1)
        at = at + 4
        pCount = 1
        ch = "("
        Do Until pCount = 0
            c = Mid(Lambda, at, 1)
            If c = "(" Then pCount = pCount + 1
            If c = ")" Then pCount = pCount - 1
            ch = ch & c
            at = at + 1
        Loop
        rht = Mid(Lambda, at)
        Lambda = lft & Quote(Chr(Evaluate(ch))) & rht
        at = InStr(UCase(Lambda), "CHR(")
    Loop
#End If
End Function

'By arbitrary convention I've assigned the # after the Lambda variable name
'  to let the caller know not to pass a Lambda function with more than # arguments
Public Sub LambdaSelectionByVal(Lambda1 As String)
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    Dim calc As XlCalculation: calc = Application.Calculation: Application.Calculation = xlCalculationManual
    On Error GoTo RestApp
    
    Dim i As Long, rowCount As Long, FirstRow As Long
    Dim sel As Range, ar As Range, col As Range, c As Range
    Dim s1 As String
    Set sel = Selection
    
    For Each ar In sel.Areas
        rowCount = ar.Columns(1).Cells.count
        FirstRow = ar.Cells(1).Row - 1
        For Each col In ar.Columns
            For i = 1 To rowCount
                Set c = col.Cells(i)
                If IsEmpty(c) Then i = c.End(xlDown).Row - FirstRow
                If i > rowCount Then Exit For
                
                Set c = col.Cells(i)
                If Not (c.Rows.Hidden Or c.Columns.Hidden) Then
                    s1 = c.Value
                    c.Value = Evaluate(Lambda(Lambda1, s1))
                End If
            Next i
        Next col
    Next ar
   
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    On Error GoTo 0
    Exit Sub

RestApp:
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    On Error GoTo 0
    Resume
End Sub

Public Sub LambdaRangeByVal(Lambda2 As String)
    Dim r As Range, ar As Range, col As Range: Set r = Selection
    Dim ct As Long
    Dim toRng As Range, s1Rng As Range, s2Rng As Range
    
    ct = 1
    For Each ar In r.Areas
        For Each col In ar.Columns
            Select Case ct
             Case 1
                Set toRng = col
                ct = 2
             Case 2
                Set s1Rng = col
                ct = 3
             Case 3
                Set s2Rng = col
                ct = 4
             Case Else
                Exit For
            End Select
        Next col
    Next ar
    
    If toRng.Rows.count < 2 Then
        toRng.Parent.Activate
        Set toRng = Application.InputBox("Select Destination Range", Type:=8)
    End If
    If ct < 3 Then Set s1Rng = Application.InputBox("Select Source 1 Range", Type:=8)
    If ct < 4 Then Set s2Rng = Application.InputBox("Select Source 2 Range", Type:=8)
    
    Dim MinCells As Long
    MinCells = toRng.Cells.count
    MinCells = IIf(MinCells <= s1Rng.Cells.count, MinCells, s1Rng.Cells.count)
    MinCells = IIf(MinCells <= s2Rng.Cells.count, MinCells, s2Rng.Cells.count)
    
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    Dim calc As XlCalculation: calc = Application.Calculation: Application.Calculation = xlCalculationManual
    On Error GoTo RestApp
    Dim s1 As String, s2 As String, i As Long
    For i = 1 To MinCells
        
        s1 = s1Rng.Cells(i).Value
        s2 = s2Rng.Cells(i).Value
        toRng.Cells(i).Value = Evaluate(Lambda(Lambda2, s1, s2))
    Next i
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    On Error GoTo 0
    Exit Sub

RestApp:
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    On Error GoTo 0
    Resume
End Sub

Private Sub Jargon_tst(): MsgBox Jargon(): End Sub
Public Function Jargon() As String
    alphabet = "abcdefghijklmnopqrstuvwxyz ¥ƒµßá"
    L = Len(alphabet)
    Randomize
    For i = 1 To 10
        Jargon = Jargon & Mid(alphabet, Int((L * Rnd) + 1), 1)
    Next i
End Function

Private Sub JargonRange(r As Range)
    r.Select
    LambdaSelectionByVal "Jargon()"
End Sub
