VERSION 5.00
Begin VB.Form folder_view 
   BorderStyle     =   0  'None
   Caption         =   "Select Picture"
   ClientHeight    =   6960
   ClientLeft      =   2445
   ClientTop       =   240
   ClientWidth     =   2340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   30
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   30
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00808000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00808000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3210
      Left            =   0
      Pattern         =   "*.bmp;*.gif;*.jpg;*.tga;*.wmf"
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
End
Attribute VB_Name = "folder_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Public path_file As String



Private Sub Dir1_Change()
File1.Path = Dir1.Path
MDIForm1.StatusBar1.Panels(4).Text = File1.ListCount & "Object(s)"
Form1.Image1.Picture = LoadPicture(none)
On Error Resume Next
    Unload Form12
Exit Sub



End Sub

Private Sub Drive1_Change()

On Error GoTo a
  Dir1.Path = Drive1.Drive
Exit Sub
a:
X = MsgBox("Device Not Ready", vbInformation, "Error...")
Dir1.Path = "c:\"
Drive1.Drive = "c:"
    
End Sub

Private Sub File1_Click()
Unload Form4
Form1.Image1.Height = 7000
Form1.Image1.Width = 9255

'Form1.Image1.Stretch = True
path_file = File1.Path

If Len(path_file) = 3 Then
Form1.Image1.Picture = LoadPicture(Dir1.Path & File1.FileName)
Else
Form1.Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
End If

Form1.Top = 0
Form1.Left = folder_view.Width + 200


'Form1.Height = Form1.Image1.Height
'Form1.Width = Form1.Image1.Width

'If Form1.Image1.Picture.Height < Form1.Image1.Height Then
 '   Form1.Image1.Stretch = False
'End If

'If Form1.Image1.Picture.Width < Form1.Image1.Width Then
 '   Form1.Image1.Stretch = False
'End If

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
MDIForm1.StatusBar1.Panels(3).Text = File1.Path & "\" & File1.FileName


End Sub




