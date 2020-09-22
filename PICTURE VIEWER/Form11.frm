VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Slide Show"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7500
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   240
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   120
      Top             =   120
      Width           =   11655
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
Timer1.Interval = Form10.Slider1.Value * 1000
i = 0

End Sub

Private Sub Timer1_Timer()

    Form11.Caption = "Slide Show ..." & Form10.List1.List(i)
    Image1.Picture = LoadPicture(Form10.List1.List(i))
    i = i + 1
    
    If i = Form10.List1.ListCount Then
     i = 0
    End If
    
    
End Sub
