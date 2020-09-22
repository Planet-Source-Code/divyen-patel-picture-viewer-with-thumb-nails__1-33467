VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Thumbnail Size"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5445
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   630
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1111
      _Version        =   393216
      Min             =   1
      SelStart        =   1
      TickStyle       =   2
      Value           =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   3000
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   5292
      _Version        =   393216
      Orientation     =   1
      Min             =   1
      SelStart        =   1
      TickStyle       =   2
      Value           =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Width"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Height"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   840
      Top             =   840
      Width           =   3000
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Form13.Hide
End Sub

Private Sub Form_Load()
Form13.Left = 3000
Form13.Top = 700

Unload Form12

Slider1.Value = 6
Slider2.Value = 5
End Sub

Private Sub Slider1_Change()
Image1.Height = Slider1.Value * 300
End Sub

Private Sub Slider1_Click()
Image1.Height = Slider1.Value * 300
End Sub


Private Sub Slider2_Change()
Image1.Width = Slider2.Value * 300
End Sub

Private Sub Slider2_Click()
Image1.Width = Slider2.Value * 300
End Sub
