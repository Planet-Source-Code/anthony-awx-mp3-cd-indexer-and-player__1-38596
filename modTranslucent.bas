Attribute VB_Name = "modTranslucent"
Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hWnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hWnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Private Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hWnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long



Public Function TranslucentForm(frm As Form, _
                TranslucenceLevel As Byte) As Boolean
  

    On Error Resume Next
    
    SetWindowLong frm.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes frm.hWnd, 0, TranslucenceLevel, _
                               LWA_ALPHA
    TranslucentForm = Err.LastDllError = 0

End Function

