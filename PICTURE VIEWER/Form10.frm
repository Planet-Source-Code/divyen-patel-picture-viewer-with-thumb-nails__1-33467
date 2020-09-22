VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slide Show Settings"
   ClientHeight    =   6465
   ClientLeft      =   2760
   ClientTop       =   1275
   ClientWidth     =   6870
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   6870
   Begin VB.Frame Frame1 
      Caption         =   "Delay ( In Seconds )"
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   3735
      Begin MSComctlLib.Slider Slider1 
         Height          =   390
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   688
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Show"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   4320
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
      Begin VB.ListBox List1 
         Height          =   3180
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Add Pictures in the Listbox and then Press Show Button"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
If List1.ListCount <> 0 Then
    Form11.Show
Else
    MsgBox "You have not Selected any Picture ...", vbInformation
    
End If

End Sub

Private Sub Command2_Click()
If Len(File1.FileName) <> 0 Then

If Len(Dir1.Path) > 3 Then
    List1.AddItem Dir1.Path & "\" & File1.FileName
ElseIf Len(Dir1.Path) = 3 Then
    List1.AddItem Dir1.Path & File1.FileName
End If
End If

End Sub

Private Sub Command3_Click()
  If List1.ListIndex = -1 Then
        MsgBox "You Have not selected any item"
    Else
        List1.RemoveItem (List1.ListIndex)
    End If
End Sub

Private Sub Command4_Click()
Unload Form10
End Sub

Private Sub Dir1_Change()

    File1.Path = Dir1.Path
    Image1.Picture = LoadPicture(none)
    Image1.Stretch = True
    Image1.Height = 1935
    Image1.Width = 2655
    
End Sub

Private Sub Drive1_Change()
On Error GoTo a

Dir1.Path = Drive1.Drive

Exit Sub

a:
    MsgBox "Drive not Ready"
    Drive1.Drive = "c:\"

End Sub

Private Sub File1_Click()
Image1.Stretch = True
Image1.Height = 1935
Image1.Width = 2655

If Len(Dir1.Path) <> 3 Then
    Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
Else
    Image1.Picture = LoadPicture(Dir1.Path & File1.FileName)
End If

If Image1.Height > Image1.Picture.Height Then
    Image1.Stretch = False
End If

If Image1.Width > Image1.Picture.Width Then
    Image1.Stretch = False
End If


End Sub

Private Sub Form_Load()
File1.Pattern = "*.bmp;*.jpg;*.gif;*.wmf"
Form10.Top = 1300
Form10.Left = 2700
Drive1.Drive = "c:\"
Dir1.Path = "c:\"
End Sub

