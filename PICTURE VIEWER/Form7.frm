VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Destination Folder"
   ClientHeight    =   3195
   ClientLeft      =   2040
   ClientTop       =   1905
   ClientWidth     =   6015
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Select the Destination folder where you want to Move Selected  Picture"
      Height          =   735
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   5895
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fo As New FileSystemObject

Private Sub Command1_Click()

If Len(folder_view.File1.FileName) <> 0 Then
If Len(folder_view.Dir1.Path) <> 3 Then

If Len(Label3.Caption) <> 3 Then

    fo.MoveFile folder_view.Dir1.Path & "\" & folder_view.File1.FileName, Form7.Label3.Caption & "\" & folder_view.File1.FileName

    folder_view.File1.Refresh
Else
    fo.MoveFile folder_view.Dir1.Path & "\" & folder_view.File1.FileName, Form7.Label3.Caption & folder_view.File1.FileName
    folder_view.File1.Refresh
End If

Else

If Len(Label3.Caption) <> 3 Then
    fo.MoveFile folder_view.Dir1.Path & folder_view.File1.FileName, Form7.Label3.Caption & "\" & folder_view.File1.FileName
Else
    fo.MoveFile folder_view.Dir1.Path & folder_view.File1.FileName, Form7.Label3.Caption & folder_view.File1.FileName
End If

folder_view.File1.Refresh

End If

Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
Unload Form7

End Sub

Private Sub Command2_Click()
Unload Form7

End Sub

Private Sub Dir1_Change()
Label3.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Label3.Caption = Dir1.Path
End Sub

