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

'This is where the VBA code gets mapped to the button
Public Sub PersonalRibbonRouter(Ctrl As IRibbonControl)
    Select Case Ctrl.id
    
    'Personal buttons
     Case "Meaningful_ID_01_Bttn"
        Module1.MySub1
     Case "Meaningful_ID_02_Bttn"
        Module1.MySub2
     Case "Meaningful_ID_03_Bttn"
        Module2.MySub1
     Case "Meaningful_ID_04_Bttn"
        Module2.MySub2
     Case "Meaningful_ID_05_Bttn"
        Module3.MySub1
     Case "Meaningful_ID_06_Bttn"
        Module3.MySub2

     Case Else
        MsgBox "Ribbon Button code not mapped." & Chr(10) & _
            "Update macro: Personal.xlam>RibbonX>PersonalRibbonButtonRouter('" & Ctrl.id & "')", vbCritical
    End Select
End Sub

Public Sub thisRibbonInitalize()
    MsgBox "personal ribbon init"
End Sub

