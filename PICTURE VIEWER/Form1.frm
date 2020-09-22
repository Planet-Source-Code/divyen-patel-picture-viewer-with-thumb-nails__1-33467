VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   7020
   ClientLeft      =   5205
   ClientTop       =   1515
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9390
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   0
      TabIndex        =   1
      Top             =   6720
      Width           =   9135
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6735
      Left            =   9120
      Max             =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub HScroll1_Change()
Image1.Left = -HScroll1.Value
End Sub













Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu MDIForm1.pmenu
    End If
    
End Sub

Private Sub VScroll1_Change()
Image1.Top = -VScroll1.Value
End Sub

