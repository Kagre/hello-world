'Copyright 2016 Gregory Kaiser
'
'This file is part of my random-code libarary.
'
'My random-code library is free software: you can redistribute it and/or modify
'it under the terms of the GNU Lesser General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'My random-code library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public License
'along with my random-code library.  If not, see <http://www.gnu.org/licenses/>.

Dim ColWidths() As Double 'used in copy/paste column widths

'converts Selected range to Proper Case _
 Ex: "Proper Case String"
Public Sub ToProper()
    LambdaSelectionByVal "PROPER(%1)"
End Sub

'Converts selected range to upper case _
 Ex: "UPPER CASE STRING"
Public Sub ToUpper()
    LambdaSelectionByVal "UPPER(%1)"
End Sub

'Converts selected range to lower case _
 Ex: "lower case string"
Public Sub ToLower()
    LambdaSelectionByVal "LOWER(%1)"
End Sub

Public Sub TrimSelection()
    LambdaSelectionByVal "TRIM(%1)"
End Sub

Public Sub RTrimSelection()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    Dim calc As XlCalculation: calc = Application.Calculation: Application.Calculation = xlCalculationManual
    On Error GoTo RestApp
    
    Dim i As Long, rowCount As Long, FirstRow As Long
    Dim col As Range
    
    rowCount = Selection.Columns(1).Cells.count
    FirstRow = Selection.Cells(1).Row - 1
    For Each col In Selection.Columns
        For i = 1 To rowCount
            If IsEmpty(col.Cells(i)) Then i = col.Cells(i).End(xlDown).Row - FirstRow
            If i > rowCount Then Exit For
            col.Cells(i).Value = RTrim(col.Cells(i).Value)
        Next i
    Next col
   
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

Public Sub formulaTransposeSelection()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    Dim calc As XlCalculation: calc = Application.Calculation: Application.Calculation = xlCalculationManual
    On Error GoTo RestApp
    
    Dim m As Long, c As Range
    Set c = Selection.Cells(1, 1)
    m = Selection.Rows.count
    If m < Selection.Columns.count Then m = Selection.Columns.count
    Dim i As Long, j As Long, swp As String
    For j = 1 To m - 1
        For i = 0 To m - j - 1
            swp = c.Offset(j + i, j - 1).Formula
            c.Offset(j + i, j - 1).Formula = c.Offset(j - 1, j + i).Formula
            c.Offset(j - 1, j + i).Formula = swp
        Next i
    Next j
    
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    c.Resize(Selection.Columns.count, Selection.Rows.count).Select
    On Error GoTo 0
    Exit Sub

RestApp:
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    On Error GoTo 0
    Resume
End Sub

'clears cells with error or empty string values
Public Sub clearJunk()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    
    Dim i As Long, rowCount As Long, FirstRow As Long
    Dim col As Range
    
    rowCount = Selection.Columns(1).Cells.count
    FirstRow = Selection.Cells(1).Row - 1
    For Each col In Selection.Columns
        For i = 1 To rowCount
            If IsEmpty(col.Cells(i)) Then i = col.Cells(i).End(xlDown).Row - FirstRow
            If i > rowCount Then Exit For
            If IsError(col.Cells(i).Value) Then
                col.Cells(i).ClearContents
            ElseIf col.Cells(i).Value = "" Then
                col.Cells(i).ClearContents
            ElseIf Trim(col.Cells(i).Value) = "" Then
                col.Cells(i).ClearContents
            End If
        Next i
    Next col
    
    Application.ScreenUpdating = scrn
End Sub

'clears cells with values of 0
Public Sub clearZeros()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    
    Dim i As Long, rowCount As Long, FirstRow As Long
    Dim col As Range
    
    rowCount = Selection.Columns(1).Cells.count
    FirstRow = Selection.Cells(1).Row - 1
    For Each col In Selection.Columns
        For i = 1 To rowCount
            If IsEmpty(col.Cells(i)) Then i = col.Cells(i).End(xlDown).Row - FirstRow
            If i > rowCount Then Exit For
            If IsNumeric(col.Cells(i).Value) Then
            If col.Cells(i).Value = 0 Then
                col.Cells(i).ClearContents
            End If: End If
        Next i
    Next col
    
    Application.ScreenUpdating = scrn
End Sub

Public Sub Touch()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    Dim calc As XlCalculation: calc = Application.Calculation: Application.Calculation = xlCalculationManual
    On Error GoTo RestApp
    
    Dim i As Long, rowCount As Long, FirstRow As Long
    Dim col As Range
    
    rowCount = Selection.Columns(1).Cells.count
    FirstRow = Selection.Cells(1).Row - 1
    For Each col In Selection.Columns
        For i = 1 To rowCount
            If IsEmpty(col.Cells(i)) Then i = col.Cells(i).End(xlDown).Row - FirstRow
            If i > rowCount Then Exit For
            If Left(col.Cells(i).Formula, 1) <> "=" Then col.Cells(i).Value = col.Cells(i).Value
        Next i
    Next col
    
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

Public Sub TextTouch()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    Dim calc As XlCalculation: calc = Application.Calculation: Application.Calculation = xlCalculationManual
    On Error GoTo RestApp
    
    Dim str As String
    Dim i As Long, rowCount As Long, FirstRow As Long
    Dim col As Range
    
    rowCount = Selection.Columns(1).Cells.count
    FirstRow = Selection.Cells(1).Row - 1
    For Each col In Selection.Columns
        For i = 1 To rowCount
            If IsEmpty(col.Cells(i)) Then i = col.Cells(i).End(xlDown).Row - FirstRow
            If i > rowCount Then Exit For
            str = col.Cells(i).Value
            col.Cells(i).NumberFormat = "@"
            col.Cells(i).Value = str
        Next i
    Next col
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

Public Sub ValueTouch()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    Dim calc As XlCalculation: calc = Application.Calculation: Application.Calculation = xlCalculationManual
    On Error GoTo RestApp
    
    Dim str As String
    Dim i As Long, rowCount As Long, FirstRow As Long
    Dim ar As Range, col As Range
    
    For Each ar In Selection.Areas
        rowCount = ar.Columns(1).Cells.count
        FirstRow = ar.Cells(1).Row - 1
        For Each col In ar.Columns
            For i = 1 To rowCount
                If IsEmpty(col.Cells(i)) Then i = col.Cells(i).End(xlDown).Row - FirstRow
                If i > rowCount Then Exit For
                col.Cells(i).Value = col.Cells(i).Value
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

'over a range, copies the value from the cell above the current one if the current cell is empty
Public Sub FillDown()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    Dim calc As XlCalculation: calc = Application.Calculation: Application.Calculation = xlCalculationManual
    On Error GoTo RestApp
    
    Dim i As Long, j As Long, rowCount As Long, FirstRow As Long
    Dim col As Range, fromR As Range: Set fromR = Selection
    
    rowCount = fromR.Columns(1).Cells.count
    FirstRow = fromR.Cells(1).Row - 1
    For Each col In fromR.Columns
        For i = 1 To rowCount
            If IsEmpty(col.Cells(i)) Then
                j = col.Cells(i).End(xlDown).Row - FirstRow - 1
                If j > rowCount Then j = rowCount
                Application.Range(col.Cells(i), col.Cells(j)).Value = col.Cells(i - 1).Value
                i = j
            ElseIf Not IsEmpty(col.Cells(i + 1)) Then
                i = col.Cells(i).End(xlDown).Row - FirstRow
            End If
SkipI:  Next i
    Next col
    fromR.Select
    
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    On Error GoTo 0
    Exit Sub

RestApp:
    If Err = 1004 Then
        i = i + 1
        Resume SkipI
    End If
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    On Error GoTo 0
    Resume
End Sub

'over a range, copies from the cell above the current one if the current cell is empty
Public Sub CopyDown()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
    Dim calc As XlCalculation: calc = Application.Calculation: Application.Calculation = xlCalculationManual
    On Error GoTo RestApp
    
    Dim i As Long, j As Long, rowCount As Long, FirstRow As Long
    Dim col As Range, fromR As Range: Set fromR = Selection
    
    rowCount = fromR.Columns(1).Cells.count
    FirstRow = fromR.Cells(1).Row - 1
    For Each col In fromR.Columns
        For i = 1 To rowCount
            If IsEmpty(col.Cells(i)) Then
                j = col.Cells(i).End(xlDown).Row - FirstRow - 1
                If j > rowCount Then j = rowCount
                col.Cells(i - 1).Copy Application.Range(col.Cells(i), col.Cells(j))
                i = j
            ElseIf Not IsEmpty(col.Cells(i + 1)) Then
                i = col.Cells(i).End(xlDown).Row - FirstRow
            End If
SkipI:  Next i
    Next col
    Application.CutCopyMode = False
    fromR.Select
    
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    On Error GoTo 0
    Exit Sub

RestApp:
    If Err = 1004 Then
        i = i + 1
        Resume SkipI
    End If
    Application.Calculation = calc: Application.Calculate
    Application.ScreenUpdating = scrn
    On Error GoTo 0
    Resume
End Sub

Public Sub setNormalToCell()
    With ActiveWorkbook.Styles("Normal").Font
        .name = ActiveCell.Font.name
        .Size = ActiveCell.Font.Size
        .Bold = False
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .Strikethrough = False
        .ThemeColor = 2
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub

'Copy the column widths and hidden columns from the active worksheet
Public Sub copyColWidths()
    Dim r As Range, i As Long, L As Long
    Set r = ActiveSheet.Cells(1, 1)
    L = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.count).Column
    ReDim ColWidths(L - 1)
    For i = 0 To L - 1
        ColWidths(i) = r.Offset(0, i).ColumnWidth
    Next i
End Sub

'Paste the column widths to the active worksheet
Public Sub pasteColWidths()
    Application.ScreenUpdating = False
    On Error GoTo SoftExit
    Dim ws As Worksheet: Set ws = ActiveSheet
    For i = LBound(ColWidths) To UBound(ColWidths)
        If ColWidths(i) <= 0 Then ws.Columns(i + 1).Hidden = True Else ws.Columns(i + 1).ColumnWidth = ColWidths(i)
    Next i
SoftExit:
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub
Private Sub outColWidths()
    Application.ScreenUpdating = False
    On Error GoTo SoftExit
    Dim ws As Worksheet: Set ws = ActiveSheet
    For i = LBound(ColWidths) To UBound(ColWidths)
        Selection.Value = Selection.Value & ", " & ColWidths(i)
    Next i
SoftExit:
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub


Public Sub ScrubSelection()
    LambdaSelectionByVal "ScrubText(%1)"
    Exit Sub
End Sub

Public Function ScrubText(Text As String) As String
    Dim i As Long, T As String, a As Long
    For i = 1 To Len(Text)
        T = Mid(Text, i, 1)
        a = AscW(T)
        If 31 < a And a < 128 Then ScrubText = ScrubText & T
    Next i
End Function

Public Function StrippedText(Text As String) As String
    Dim i As Long, T As String, a As Long
    For i = 1 To Len(Text)
        T = Mid(Text, i, 1)
        a = AscW(T)
        If a > 128 Or a < 31 Then StrippedText = StrippedText & T
    Next i
End Function

Public Sub UnWrapText(): Selection.WrapText = False: End Sub
Public Sub removeAllLinks(): ActiveSheet.Cells.Hyperlinks.Delete: End Sub
Public Function DAdd(S As String, N As Long, D As Date) As Date: DAdd = DateAdd(S, N, D): End Function
Public Function IsFunction(ref As Range) As Boolean: IsFunction = (Left(ref.Cells.Formula, 1) = "="): End Function
Public Function Quote(S As String) As String: Quote = """" & S & """": End Function

Public Sub ChartLineWidth()
    Dim c As Chart, S As Series
    On Error Resume Next
    For Each co In ActiveSheet.ChartObjects()
        Set c = co.Chart
        For Each S In c.SeriesCollection
            S.Format.Line.Weight = 1
            S.MarkerSize = 2
        Next S
    Next co
    On Error GoTo 0
End Sub

Public Sub ChartPointSize()
    Dim c As Chart, S As Series
    On Error GoTo ErrHandler
    Set c = ActiveSheet
    On Error Resume Next
    For Each S In c.SeriesCollection
        S.MarkerSize = 2
    Next S
    GoTo EOSb
    
SkipSpreadSheet:
    For Each co In ActiveSheet.ChartObjects()
        Set c = co.Chart
        For Each S In c.SeriesCollection
            S.MarkerSize = 2
    Next S: Next co

EOSb: On Error GoTo 0
    Exit Sub
ErrHandler:
    On Error Resume Next
    Resume SkipSpreadSheet
End Sub

Public Sub ShelfSelection()
    Dim v As Variant
    Dim c As Range, r As Range, S As Range
    Application.ScreenUpdating = False
    Set S = Selection
    For Each r In S.Columns
        v = Empty
        For Each c In r.Cells
            If v = c.Value Then
                c.ClearContents
            Else: v = c.Value
            End If
        Next c
    Next r
    Application.ScreenUpdating = True
End Sub
