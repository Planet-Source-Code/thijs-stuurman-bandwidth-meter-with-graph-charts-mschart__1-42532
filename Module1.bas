Attribute VB_Name = "Module1"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, _
ByVal y As Long, ByVal CX As Long, ByVal CY As Long, _
ByVal wFlags As Long) As Long
'constants for SetWindowPos API
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Public Sub SetOnTop(ByVal hwnd As Long, ByVal bSetOnTop As Boolean)
Dim lR As Long
If bSetOnTop Then
lR = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
Else
lR = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End If
End Sub

