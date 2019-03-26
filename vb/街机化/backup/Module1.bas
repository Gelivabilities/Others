Attribute VB_Name = "Module1"
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Sub keyR()
    Const KeyDown = -32767
    
    For KeyCode = 65 To 65
        If GetAsyncKeyState(KeyCode) = KeyDown Then LogKey KeyCode
    Next
End Sub
Private Sub LogKey(KeyCode)
    
    Const VK_F1 = &H70
    Const VK_F24 = &H87
    

    If KeyCode >= VK_F1 And KeyCode <= VK_F24 Then

      Select Case KeyCode
        Case 65:  Label2.Caption = Str(Int(Label2.Caption) + 1)
    End Select
    
    End If
    
End Sub
