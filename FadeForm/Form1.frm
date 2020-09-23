VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Open with Timer (Show vbModal)"
      Height          =   435
      Left            =   983
      TabIndex        =   2
      Top             =   2160
      Width           =   2715
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Standard"
      Height          =   435
      Left            =   983
      TabIndex        =   1
      Top             =   1320
      Width           =   2715
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      Left            =   173
      Max             =   255
      TabIndex        =   0
      Top             =   480
      Value           =   255
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    'use the "timer" for fade counter, if the form show in vbModal
    Form3.Show vbModal
End Sub

Private Sub HScroll1_Change()
    FadeForm Me, HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    FadeForm Me, HScroll1.Value
End Sub
