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

