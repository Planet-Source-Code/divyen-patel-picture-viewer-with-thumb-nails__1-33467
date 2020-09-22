VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open Directory"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6600
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Folder Path"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   195
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo a
folder_view.Drive1.Drive = Text1.Text
folder_view.Dir1.Path = Text1.Text
Exit Sub

a:
    MsgBox "Invalid Folder Path...", vbOKOnly
End Sub
