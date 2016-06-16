Private Sub cellFormatInfo()
    Const NL As String = vbNewLine
    Const QO As String = """"
    Dim c As Range, msg As String
    Set c = ActiveCell
    msg = ""
    msg = msg & "Font.Name: " & c.Font.Name & NL
    msg = msg & "Font.Color: " & c.Font.Color & NL
    msg = msg & "Font.ColorIndex: " & c.Font.ColorIndex & NL
    msg = msg & "Font.Size: " & c.Font.Size & NL
    If c.Font.Bold Then msg = msg & "Font.Bold: Yes" & NL
    If c.Font.Italic Then msg = msg & "Font.Italic: Yes" & NL
    msg = msg & "HorizontalAlignment: "
    Select Case c.HorizontalAlignment
      Case 1
        msg = msg & "Default=1"
      Case xlCenter
        msg = msg & "xlCenter"
      Case xlDistributed
        msg = msg & "xlDistributed"
      Case xlJustify
        msg = msg & "xlJustify"
      Case xlLeft
        msg = msg & "xlLeft"
      Case xlRight
        msg = msg & "xlRight"
      Case Else
        msg = msg & "Unknown=" & c.HorizontalAlignment
    End Select
    msg = msg & NL
    msg = msg & "VerticalAlignment: "
    Select Case c.VerticalAlignment
      Case xlCenter
        msg = msg & "xlCenter"
      Case xlDistributed
        msg = msg & "xlDistributed"
      Case xlJustify
        msg = msg & "xlJustify"
      Case xlBottom
        msg = msg & "xlBottom"
      Case xlTop
        msg = msg & "xlTop"
      Case Else
        msg = msg & "Unknown=" & c.VerticalAlignment
    End Select
    msg = msg & NL
    msg = msg & "Interior.Color: " & c.Interior.Color & NL
    msg = msg & "Interior.ColorIndex: " & c.Interior.ColorIndex & NL
    msg = msg & "NumberFormat: " & c.NumberFormat & NL
    msg = msg & "Borders.ColorIndex: " & c.Borders.ColorIndex & NL
    msg = msg & "Borders.Weight: " & c.Borders.Weight & NL
    msg = msg & "Borders.LineStyle: " & c.Borders.LineStyle & NL
    MsgBox msg
End Sub

Private Sub getSelectedColumnWidths()
    Dim c As Range, o As Range
    Set o = Selection.Cells(1)
    For Each c In Selection.Columns
        o.Value = o.Value & ", " & c.ColumnWidth
    Next c
End Sub

Private Sub ConditionalFormatExamples()
    Dim QO As String: QO = Chr(34)
    Dim fc As FormatCondition
    If ActiveCell.FormatConditions.count < 3 Then
        Set fc = ActiveCell.FormatConditions.Add(xlExpression, , "=$O3=" & QO & "New" & QO)
        With fc.Interior
            .Pattern = xlSolid
            .ColorIndex = 4
        End With
    End If
End Sub

Private Sub GetRGB_tst(): MsgBox GetRGB(ActiveCell.Interior.Color): End Sub
Function GetRGB(Color As Long) As String
    Dim r As Long, G As Long, B As Long
    Dim num As String: num = "0123456789ABCDEF"
    r = Color \ 256 ^ 0 Mod 256
    GetRGB = Mid(num, r \ 16 + 1, 1) & Mid(num, r Mod 16 + 1, 1)
    G = Color \ 256 ^ 1 Mod 256
    GetRGB = GetRGB & Mid(num, G \ 16 + 1, 1) & Mid(num, G Mod 16 + 1, 1)
    B = Color \ 256 ^ 2 Mod 256
    GetRGB = GetRGB & Mid(num, B \ 16 + 1, 1) & Mid(num, B Mod 16 + 1, 1)
End Function

Private Function getLuma(r As Integer, G As Integer, B As Integer) As Double
    Const Wr As Double = 0.299
    Const Wb As Double = 0.114
    Const Wg As Double = 1 - Wr - Wb
    Const Umax As Double = 0.436
    Const Vmax As Double = 0.615
    Dim U As Double, V As Double
    getLuma = Wr * r + Wg * G + Wb * B
    U = (Umax * (B - getLuma)) / (1 - Wb)
    V = (Vmax * (r - getLuma)) / (1 - Wr)
End Function

Sub ChartLineWidth()
    Dim c As Chart, s As Series
    For Each co In ActiveSheet.ChartObjects()
        Set c = co.Chart
        For Each s In c.SeriesCollection
            s.Format.Line.Weight = 1
        Next s
    Next co
End Sub
