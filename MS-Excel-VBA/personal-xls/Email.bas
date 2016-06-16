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

'Requires Ref: Microsoft Outlook 15.0 Object Library

Private Sub MakeEmail_tst()
    MakeEmail True, "", "Test Email", ""
End Sub

Private Function MakeEmail(MarkAutoSend As Boolean, EmailTemplate As String, EmailSubject As String, FileAttachments As Variant) As Outlook.MailItem
    Dim OlApp As Outlook.Application
    On Error GoTo Fail
    Set OlApp = GetObject(, "Outlook.Application")
    Dim obj As Object, em As Outlook.MailItem, TemplateName As String
    
    If EmailTemplate = "" Then
        Set obj = OlApp.CreateItem
    Else
        Set obj = OlApp.CreateItemFromTemplate(TemplateName)
    End If
    Set em = obj
    
    em.Subject = EmailSubject & Format(Now(), " m/d")
    em.HTMLBody = RangetoHTML(ActiveSheet.UsedRange)
    
    If VarType(FileAttachments) = vbString Then
        If FileAttachments <> "" Then em.Attachments.Add FileAttachments
    ElseIf IsArray(FileAttachments) Then
        Dim i As Long
        For i = LBound(FileAttachments) To UBound(FileAttachments)
            em.Attachments.Add FileAttachments(i)
        Next i
    End If
    
    If MarkAutoSend Then em.Categories = "AutoSend" & IIf(Len(em.Categories) > 0, ", ", "") & em.Categories
    
    em.Save
    Set MakeAuditEmail = em
EOFn: On Error GoTo 0
    Exit Function
Fail:
    If Err = 429 Then
        Set OlApp = CreateObject("Outlook.Application")
        Resume Next
    ElseIf Err = -2147287037 Then
        'ToUpdate -- template missing
    End If
    MsgBox Err.Description, vbCritical, "Ignored Error"
    Resume EOFn
End Function
