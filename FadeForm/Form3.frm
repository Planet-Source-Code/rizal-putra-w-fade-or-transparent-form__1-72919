VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   480
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FadeForm Me, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AnimFadeClose Me
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    
    Const FadeStep = 3
    
    Do Until i = 255
        If i + FadeStep >= 255 Then
            FadeForm Me, 255
            Exit Do
        End If
            
        i = i + FadeStep
        FadeForm Me, CByte(i)
        DoEvents
    Loop
    
    Timer1.Enabled = False
End Sub
