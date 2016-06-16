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

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function GetLocalTime Lib "kernel32" _
    (lpSystemTime As SYSTEMTIME) As Long

Public Function MilliNow() As Date
    Const conv As Double = 86400000# '= 24 hours/day * 60 min/hour * 60 sec/min * 1000 millisec/sec
    Dim st As SYSTEMTIME, e As Long
    e = GetLocalTime(st)
    MilliNow = CDate(st.wMonth & "/" & st.wDay & "/" & st.wYear) _
        + TimeValue(st.wHour & ":" & st.wMinute & ":" & st.wSecond) _
        + CDate(st.wMilliseconds / conv)
End Function

Public Function MilliStr(t As Date, Optional frmt As String = "") As String
    Dim mill As Double
    mill = (t - Fix(t)) * 24        '24 hours/day
    mill = (mill - Fix(mill)) * 60  '60 min/hour
    mill = (mill - Fix(mill)) * 60  '60 sec/min
    mill = mill - Fix(mill)         'remove sec
    
    MilliStr = Format(t, IIf(frmt = "", "mm/dd/yy hh:mm:ss", frmt)) & IIf(InStr(frmt, "ss") > 0, Format(mill, ".0000"), "")
End Function

Private Sub MilliNow_tst()
    Dim out As String, val As Date, mil As Double
    GregRibbon.MilliTime.MilliNow
    val = MilliNow
    out = Format(val, "m/d/yy hh:mm:ss")
    mil = (val - Fix(val)) * 24
    mil = (mil - Fix(mil)) * 60
    mil = (mil - Fix(mil)) * 60
    mil = mil - Fix(mil)
    out = out & Format(mil, ".0000")
    MsgBox out
End Sub

Private Sub Wait2()
    Dim d As Double, L As Long
    L = 1
    d = 0
    Do Until d > 18
       d = d + 1 / L
       L = L + 1
    Loop
End Sub

Private Sub Wait2_tst()
    Dim beg As Date, fin As Date, milb As Double, milf As Double
    beg = MilliNow()
    Wait2
    fin = MilliNow()
    
    milb = (beg - Fix(beg)) * 24
    milb = (milb - Fix(milb)) * 60
    milb = (milb - Fix(milb)) * 60
    'milb = milb - Fix(milb)
    
    milf = (fin - Fix(fin)) * 24
    milf = (milf - Fix(milf)) * 60
    milf = (milf - Fix(milf)) * 60
    'milf = milf - Fix(milf)
    
    MsgBox (milf - milb)
End Sub
