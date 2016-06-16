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

Option Explicit
'Office VBA Module: WinSCP
' * uses Excel specific code, can be updated with light modification
' * uses Windows 7 path structure
'     ToDo - private const lSep as string = "\"
'     ToDo - private const sSep as string = "/"
' * requires Read/Write permissions to BatchPath folder
Private Const QO As String = """"
Private Const NL As String = vbNewLine
Private Const SecondsPerDay As Double = 24# * 60# * 60#
Private Const BytesPerKB As Long = 2 ^ 10
Private Const BytesPerMB As Long = 2 ^ 20
Private Const BytesPerGB As Long = 2 ^ 30
Private Const BytesPerTB As Long = 2 ^ 40

'Ideally you'd put the Batch files in the same folder as WinSCP
Private Const BatchPath As String = "C:\Program Files (x86)\WinSCP\"
Private Const BatchGet As String = "WinSCPgetFile.bat"
Private Const BatchGetFlag As String = "WinSCPgetFile.active"
Private Const BatchGetErr As String = "WinSCPgetFile.error"
Private Const BatchSend As String = "WinSCPsendFile.bat"
Private Const BatchSendFlag As String = "WinSCPsendFile.active"
Private Const BatchSendErr As String = "WinSCPsendFile.error"

Private Function quote(s As String) As String: quote = QO & s & QO: End Function
Private Sub PauseSec(Optional ByVal Seconds As Double = 1#)
    Dim wait As Date
    wait = Now() + Seconds / SecondsPerDay
    Do Until Now() > wait
        DoEvents
    Loop
End Sub

'Notes:
' * Contains Excel specific code
' * Contains OS and Local specific code
' * Requires Read/Write permissions to the BatchPath folder
'Arguments:
' FileToSend - fully qualified file path
'   Ex: "C:\sendme.txt"
' DestinationPath - relative path to where file should be saved
'   Ex: ".\subfolder\"
' Connection - the name of the WinSCP connection to use
'   Ex: "MySavedSession"
'Returns:
' true  - if successful
' false - otherwise
Public Function sFTPsendFile(ByVal FileToSend As String, ByVal DestinationPath As String, ByVal Connection As String) As Boolean
    sFTPsendFile = False
    Dim fts As String, dp As String, con As String
    fts = FileToSend: dp = DestinationPath: con = Connection
    
    'Argument validation
    'Implementation note: these were not done in an elseif structure
    ' so that all of the validations are performed
    If Len(Trim(fts)) = 0 Then
        'ToDo - implement more advanced file picker dialog
        Dim FDlg As FileDialog
        Set FDlg = Application.FileDialog(msoFileDialogFilePicker)
        FDlg.Title = "Select a file to sFTP via WinSCP"
        FDlg.Filters.Clear
        FDlg.Filters.Add "Any File", "*.*"
        FDlg.Filters.Add "MS-Excel Spreadsheet", "*.xls,*.xlsx,*.xlsm,*.xlsb"
        FDlg.Filters.Add "Text File", "*.txt,*.csv"
        FDlg.AllowMultiSelect = False
        FDlg.InitialView = msoFileDialogViewList
        FDlg.InitialFileName = ThisWorkbook.Path 'MS-Access: CurrentProject.Path
        If FDlg.Show Then fts = FDlg.SelectedItems(1)
        If Len(Trim(fts)) = "" Then Exit Function
    End If
    'Non-Excel Implementation: update Application.PathSeparator to local specific string "\"
    Dim at As Long: at = InStrRev(fts, Application.PathSeparator)
    If at = 0 Then
        'ToDo - adapt path defaulting to local and preference
        fts = ThisWorkbook.Path & Application.PathSeparator & fts
        at = Len(ThisWorkbook.Path) + 1
    End If
    Dim UNCpath As Boolean: UNCpath = (InStr(fts, "\\") > 0)
    Dim fs As Object: Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(Trim(fts)) Then Exit Function
    If fs.FileExists(BatchPath & BatchSendFlag) Then
        'ToDo - write code to wait for other instance to finish
        '     - include timeout feature so it doesn't wait forever
        Exit Function
    End If
    If fs.FileExists(BatchPath & BatchSendErr) Then
        'ToDo - write code to wait for the other instance to finish
        '     - include timeout feature so it doesn't wait forever
        Exit Function
    End If
    If Len(Trim(dp)) = 0 Then
        'ToDo - implement more advanced missing destionation path routine
        dp = InputBox("Please supply the destination path for the file:" & NL & _
                       fts & NL & _
                       "For Example: ./subfolder/")
        If Len(Trim(dp)) = 0 Then Exit Function
    End If
    If Len(Trim(con)) = 0 Then
        'ToDo - implement more advanced connection picker
        con = InputBox("" & _
           "Please specify which WinSCP connection to use for sending:" & NL & _
           " File: " & quote(fts) & NL & _
           " Dest: " & quote(dp))
        If Len(Trim(con)) = 0 Then Exit Function
    End If
    
    fts = Trim(fts): dp = Trim(dp): con = Trim(con)
    
    Dim Path As String, File As String
    Path = Left(fts, at)
    File = Mid(fts, at + 1)
    
    'Since the shell command sent to the OS is executed in parallel with this
    ' VBA code I use the FileSystem as a semaphore to signal when the file
    ' trnasfer has completed.
    Dim f As Object
    Set f = fs.CreateTextFile(BatchPath & BatchSendFlag)
    f.Write "Temparary semaphore file used to signal upload status."
    f.Close
    
    Dim sz As Variant
    Set f = fs.GetFile(fts)
    sz = f.Size
    Set f = Nothing
    
    'Run the send batch file with the verified parameters
    Shell BatchPath & BatchSend & " " & _
        quote(Path) & " " & _
        quote(File) & " " & _
        quote(dp) & " " & _
        quote(con) & " " & _
        IIf(UNCpath, "YES", "NO")
    
    'Standard time doubling delay routine
    'This allows the FileSystem a chance to catch up
    Dim delay As Double
    delay = sz / (2 * BytesPerTB) 'Assume transfer rate of 2TB/sec and work down from there
    Do While fs.FileExists(BatchPath & BatchSendFlag)
        PauseSec IIf(delay < 1, 1, delay)
        If delay > SecondsPerDay Then Exit Function
        delay = delay * 2
    Loop
    
    If fs.FileExists(BatchPath & BatchSendErr) Then
        PauseSec 5 'give the FileSystem time to finish building the file
        fs.DeleteFile BatchPath & BatchSendErr
        Exit Function
    End If
    
    sFTPgetFile = True
End Function


'Notes:
' * Contains Excel specific code
' * Contains OS and Local specific code
' * Requires Read/Write permissions to the BatchPath folder
'Arguments:
' FileToGet - Server-side relative path to file
'   Ex: "./subfolder/getMe.txt"
' SaveToPath - absolute local path to where file should be saved
'   Ex: "C:\myFolder\"
' Connection - the name of the WinSCP connection to use
'   Ex: "MySavedSession"
'Returns:
' true  - if successful
' false - otherwise
Public Function sFTPgetFile(ByVal FileToGet As String, ByVal SaveToPath As String, ByVal Connection As String) As Boolean
    sFTPgetFile = False
    Dim ftg As String, sp As String, con As String
    ftg = FileToGet: sp = SaveToPath: con = Connection
    
    'Argument Validation
    If Len(Trim(ftg)) = 0 Then
        'ToDo -- implement a more advanced missing file routine
        ftg = InputBox("Please supply the relative path and name for the file to download" & NL & _
                       " For Example: " & quote("./SubFolder/toGet.txt"))
        If Len(Trim(ftg)) = 0 Then Exit Function
    End If
    'WinSCP uses the / path separator regardless of local
    Dim at As Long: at = InStrRev(ftg, "/")
    If Len(Trim(sp)) = 0 Then
        'ToDo -- implement a more advanced save folder picker
        Dim FldrDlg As FileDialog
        Set FldrDlg = Application.FileDialog(msoFileDialogFolderPicker)
        FldrDlg.AllowMultiSelect = False
        FldrDlg.InitialView = msoFileDialogViewList
        FldrDlg.InitialFileName = ThisWorkbook.Path & "\." 'MS-Access: CurrentProject.Path
        If FldrDlg.Show Then sp = FDlg.SelectedItems(1)
        If Len(Trim(sp)) = 0 Then Exit Function
    End If
    Dim UNCpath As Boolean: UNCpath = (InStr(sp, "\\") > 0)
    Dim fs As Object: Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(Trim(sp)) Then Exit Function
    If Len(Trim(con)) = 0 Then
        'ToDo - implement more advanced connection picker
        con = InputBox("Please specify which WinSCP connection to use for getting the file:" & NL & ftg)
        If Len(Trim(con)) = 0 Then Exit Function
    End If
    
    ftg = Trim(ftg): sp = Trim(sp): con = Trim(con)
    
    Dim Path As String, File As String
    Path = IIf(at > 0, Left(ftg, at), "./")
    File = Mid(ftg, at + 1)
    
    'Since the shell command sent to the OS is executed in parallel with this
    ' VBA code I use the FileSystem as a semaphore to signal when the file
    ' trnasfer has completed.
    Dim f As Object
    Set f = fs.CreateTextFile(BatchPath & BatchGetFlag)
    f.Write "Temparary semaphore file used to signal download status."
    f.Close
    Set f = Nothing
    
    'Run the send batch file with the verified parameters
    Shell BatchPath & BatchGet & " " & _
        quote(Path) & " " & _
        quote(File) & " " & _
        quote(sp) & " " & _
        quote(con) & " " & _
        IIf(UNCpath, "YES", "NO")
    
    'Standard time doubling delay routine
    'This allows the FileSystem a chance to catch up
    Dim delay As Double: delay = 1
    Do While fs.FileExists(BatchPath & BatchGetFlag)
        PauseSec delay
        If delay > SecondsPerDay Then Exit Function
        delay = delay * 2
    Loop
    
    If fs.FileExists(BatchPath & BatchGetErr) Then
        PauseSec 5 'give the FileSystem time to finish building the file
        fs.DeleteFile BatchPath & BatchGetErr
        Exit Function
    End If
    
    sFTPgetFile = True
End Function

'ToDo - Change macro to public once it has been written
Private Sub sFTPsendFiles()
    'ToDo - generate custom WinSCP script file to run, that has multiple
    '  put "toSend.txt" lines in it.
End Sub
