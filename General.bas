Attribute VB_Name = "Module1"
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long
Public Sub formdrag(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.hWnd, &HA1, 2, 0&)
End Sub

