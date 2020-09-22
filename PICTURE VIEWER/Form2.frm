VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open File"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Enter File Path"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   190
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo a
    If Len(Form2.Text1.Text) = 0 Then
        GoTo a:
    End If
    
    Form1.Image1.Picture = LoadPicture(Form2.Text1.Text)
    Form1.Top = 0
       Form1.Left = folder_view.Width + 200
       Form1.Height = 7000
    Form1.HScroll1.Max = 0
Form1.VScroll1.Max = 0

If Form1.Image1.Width > Form1.HScroll1.Width Then
    Form1.HScroll1.Max = Form1.Image1.Width - Form1.HScroll1.Width
End If

If Form1.Image1.Height > Form1.VScroll1.Height Then
    Form1.VScroll1.Max = Form1.Image1.Height - Form1.VScroll1.Height
End If

Form1.HScroll1.LargeChange = 1000
Form1.VScroll1.LargeChange = 1000

Form1.HScroll1.SmallChange = 500
Form1.VScroll1.SmallChange = 500
MDIForm1.StatusBar1.Panels(3).Text = folder_view.File1.Path & "\" & folder_view.File1.FileName

Form2.Hide
    
Exit Sub

a:
    MsgBox "Invalid file or File not Found", vbOKOnly
    
End Sub
