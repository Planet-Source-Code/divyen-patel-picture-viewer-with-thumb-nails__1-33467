VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Destination Folder"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6180
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Select the Destination folder where you want to copy Selected  Picture"
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fo As New FileSystemObject

Private Sub Command1_Click()

If Len(folder_view.File1.FileName) <> 0 Then
If Len(folder_view.Dir1.Path) <> 3 Then

If Len(Label3.Caption) <> 3 Then

    fo.CopyFile folder_view.Dir1.Path & "\" & folder_view.File1.FileName, Form6.Label3.Caption & "\" & folder_view.File1.FileName
    folder_view.File1.Refresh
Else
    fo.CopyFile folder_view.Dir1.Path & "\" & folder_view.File1.FileName, Form6.Label3.Caption & folder_view.File1.FileName
    folder_view.File1.Refresh
End If

Else

If Len(Label3.Caption) <> 3 Then
    fo.CopyFile folder_view.Dir1.Path & folder_view.File1.FileName, Form6.Label3.Caption & "\" & folder_view.File1.FileName
Else
    fo.CopyFile folder_view.Dir1.Path & folder_view.File1.FileName, Form6.Label3.Caption & folder_view.File1.FileName
End If

folder_view.File1.Refresh

End If

Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If

Unload Form6

End Sub

Private Sub Command2_Click()
Unload Form6

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
