VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename File"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Filename without Extention"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ext As String
Dim fo As New FileSystemObject

Private Sub Command1_Click()
If Len(folder_view.Dir1.Path) <> 3 Then

Path = folder_view.Dir1.Path & "\" & folder_view.File1.FileName
dpath = folder_view.Dir1.Path & "\"

fo.MoveFile Path, dpath & Text1.Text & ext

Unload Form5

folder_view.File1.Refresh
Else
Path = folder_view.Dir1.Path & folder_view.File1.FileName
dpath = folder_view.Dir1.Path

fo.MoveFile Path, dpath & Text1.Text & ext

Unload Form5

folder_view.File1.Refresh

End If

End Sub

Private Sub Command2_Click()
Unload Form5
End Sub

Private Sub Form_Load()
folder_view.Refresh
Form5.Refresh
Text1.Text = Mid(folder_view.File1.FileName, 1, (Len(folder_view.File1.FileName) - 4))
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
ext = Mid(folder_view.File1.FileName, (Len(folder_view.File1.FileName) - 3), Len(folder_view.File1.FileName))

End Sub

