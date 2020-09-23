Attribute VB_Name = "modFade"
Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Dim i As Integer

Public Sub FadeForm(Frm As Form, Level As Byte)
    On Error Resume Next
    Dim msg As Long
    
    msg = GetWindowLong(Frm.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetWindowLong Frm.hwnd, GWL_EXSTYLE, msg
    SetLayeredWindowAttributes Frm.hwnd, 0, Level, LWA_ALPHA
End Sub

Public Sub AnimFadeOpen(Frm As Form, Optional Step As Byte = 3)
    FadeForm Frm, 0
    Frm.Show
        
    Do Until i = 255
        If i + Step >= 255 Then
            FadeForm Frm, 255
            Exit Do
        End If
            
        i = i + Step
        FadeForm Frm, CByte(i)
        DoEvents
    Loop
End Sub

Public Sub AnimFadeClose(Frm As Form, Optional Step As Byte = 5)
    i = 255
        
    Do Until i = 0
        If i - Step <= 0 Then
            FadeForm Frm, 0
            Exit Do
        End If
            
        i = i - Step
        FadeForm Frm, CByte(i)
        DoEvents
    Loop
End Sub
