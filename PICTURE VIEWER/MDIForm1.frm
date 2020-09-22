VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00808080&
   Caption         =   "Picture Viewer ( Version 1.0 )"
   ClientHeight    =   6255
   ClientLeft      =   2505
   ClientTop       =   870
   ClientWidth     =   9075
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton Command9 
         Height          =   615
         Left            =   7440
         Picture         =   "MDIForm1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Set Image as wallpaper"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         Picture         =   "MDIForm1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Show Thumbnails"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   6480
         Picture         =   "MDIForm1.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Slide Show"
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Select File Pattern "
         Top             =   0
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "MDIForm1.frx":0FD0
         Left            =   9480
         List            =   "MDIForm1.frx":0FE6
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Height          =   615
         Left            =   4560
         Picture         =   "MDIForm1.frx":1020
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Copy to"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Height          =   615
         Left            =   3600
         Picture         =   "MDIForm1.frx":1462
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Move to"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Rename"
         Height          =   615
         Left            =   2760
         Picture         =   "MDIForm1.frx":18A4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Rename Selected File"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Height          =   615
         Left            =   1800
         Picture         =   "MDIForm1.frx":1CE6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Delete Selected File (Not Recovered)"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   840
         Picture         =   "MDIForm1.frx":2128
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "File Information"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton up 
         Height          =   615
         Left            =   0
         Picture         =   "MDIForm1.frx":256A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "CD.. (Go to the Parent Directory )"
         Top             =   0
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5160
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "05-Apr-02"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "7:13 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7805
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Filemenu 
      Caption         =   "&File"
      Begin VB.Menu Openfile 
         Caption         =   "&Open"
      End
      Begin VB.Menu opendir 
         Caption         =   "Open &Directory"
      End
      Begin VB.Menu sse 
         Caption         =   "-"
      End
      Begin VB.Menu renamemnu 
         Caption         =   "&Rename"
      End
      Begin VB.Menu Delmenu 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu ssep 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu optionmenu 
      Caption         =   "&Option"
      Begin VB.Menu tset 
         Caption         =   "&Thumbnail Settings"
      End
      Begin VB.Menu theme 
         Caption         =   "&Themes"
         Begin VB.Menu Green 
            Caption         =   "&Green"
         End
         Begin VB.Menu Red 
            Caption         =   "&Red"
         End
         Begin VB.Menu blue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu sep 
            Caption         =   "-"
         End
         Begin VB.Menu choice 
            Caption         =   "&Choice"
         End
      End
   End
   Begin VB.Menu sdmenu 
      Caption         =   "Slideshow"
      Begin VB.Menu sdsettings 
         Caption         =   "&Slide-Show Settings"
      End
   End
   Begin VB.Menu Abolutmenu 
      Caption         =   "About"
      Begin VB.Menu apv 
         Caption         =   "About Picture Viewer"
      End
   End
   Begin VB.Menu pmenu 
      Caption         =   "Pmenu"
      Visible         =   0   'False
      Begin VB.Menu setwallpaper 
         Caption         =   "Set As Wall-Paper"
      End
      Begin VB.Menu popsep1 
         Caption         =   "-"
      End
      Begin VB.Menu cpymenu 
         Caption         =   "Copy to"
      End
      Begin VB.Menu moveto 
         Caption         =   "Move To"
      End
      Begin VB.Menu remenu 
         Caption         =   "Rename"
      End
      Begin VB.Menu Delmenu2 
         Caption         =   "Delete"
      End
      Begin VB.Menu popsep2 
         Caption         =   "-"
      End
      Begin VB.Menu infomenu 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fo As New FileSystemObject

Private Sub apv_Click()
frmAbout.Show
End Sub

Private Sub blue_Click()
blue.Checked = True
Green.Checked = False
Red.Checked = False
choice.Checked = False

folder_view.Dir1.BackColor = &HFF8080
folder_view.File1.BackColor = &HFF8080
End Sub

Private Sub choice_Click()
Green.Checked = False
Red.Checked = False
blue.Checked = False
choice.Checked = True

cd1.ShowColor
folder_view.Dir1.BackColor = cd1.Color
folder_view.File1.BackColor = cd1.Color

End Sub




Private Sub Combo1_Click()

If Combo1.Text = "*.bmp" Then
    folder_view.File1.Pattern = "*.bmp"
ElseIf Combo1.Text = "*.gif" Then
    folder_view.File1.Pattern = "*.gif"
ElseIf Combo1.Text = "*.jpg" Then
    folder_view.File1.Pattern = "*.jpg"
ElseIf Combo1.Text = "*.wmf" Then
    folder_view.File1.Pattern = "*.wmf"
ElseIf Combo1.Text = "*.tga" Then
    folder_view.File1.Pattern = "*.tga"
ElseIf Combo1.Text = "All Picture Files" Then
    folder_view.File1.Pattern = "*.bmp;*.gif;*.jpg;*.tga;*.wmf"
End If

StatusBar1.Panels(4).Text = folder_view.File1.ListCount & "Object(s)"

On Error Resume Next
    Unload Form12
Exit Sub

End Sub



Private Sub Command1_Click()
Form9.Show
End Sub

Private Sub Command2_Click()
If Len(folder_view.File1.FileName) <> 0 Then
    Form4.Show
Else
    MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If


End Sub

Private Sub Command3_Click()
p = MsgBox("Are you Sure you want to Delete File", vbYesNo, "Delete ...?")

If p = 6 Then

If Len(folder_view.File1.FileName) <> 0 Then

On Error GoTo c

If Len(folder_view.Dir1.Path) <> 3 Then

fo.DeleteFile folder_view.Dir1.Path & "\" & folder_view.File1.FileName
folder_view.File1.Refresh
Else

fo.DeleteFile folder_view.Dir1.Path & folder_view.File1.FileName
folder_view.File1.Refresh

End If

Exit Sub
c:
    MsgBox "File Attribute is Read Only You can not Delete it..."

Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
End If

On Error Resume Next
    Unload Form12
Exit Sub

End Sub

Private Sub Command4_Click()
If Len(folder_view.File1.FileName) <> 0 Then
Form5.Show
Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
On Error Resume Next
    Unload Form12
Exit Sub
End Sub

Private Sub Command5_Click()

If Len(folder_view.File1.FileName) <> 0 Then
Form7.Show
Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
On Error Resume Next
    Unload Form12
Exit Sub
End Sub

Private Sub Command6_Click()
If Len(folder_view.File1.FileName) <> 0 Then
Form6.Show
Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
On Error Resume Next
    Unload Form12
Exit Sub
End Sub

Private Sub Command7_Click()
Form10.Show
End Sub

Private Sub Command8_Click()
If folder_view.File1.ListCount <> 0 Then
    Form12.Show
Else
    MsgBox "No Picture Found in Current Folder ...", vbInformation, "No Image Found..."
End If

End Sub

Private Sub Command9_Click()
Dim wallp As String
Dim e As String
Dim f As New FileSystemObject
wallp = Clear

If Len(folder_view.Dir1.Path) <> 3 Then
    wallp = folder_view.File1.Path & "\" & folder_view.File1.FileName
Else
    wallp = folder_view.File1.Path & folder_view.File1.FileName
End If
           
e = f.GetExtensionName(wallp)
If e = "gif" Or e = "GIF" Then
    MsgBox "Invalid File...Select BMP OR JPEG File", vbOKOnly Or vbInformation
    GoTo Y:
    
End If

X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, wallp, 2)
Y:

End Sub

Private Sub cpymenu_Click()
If Len(folder_view.File1.FileName) <> 0 Then
Form6.Show
Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
On Error Resume Next
    Unload Form12
Exit Sub
End Sub

Private Sub delmenu_Click()
p = MsgBox("Are you Sure you want to Delete File", vbYesNo, "Delete ...?")

If p = 6 Then
If Len(folder_view.File1.FileName) <> 0 Then

On Error GoTo c

If Len(folder_view.Dir1.Path) <> 3 Then

fo.DeleteFile folder_view.Dir1.Path & "\" & folder_view.File1.FileName
folder_view.File1.Refresh
Else

fo.DeleteFile folder_view.Dir1.Path & folder_view.File1.FileName
folder_view.File1.Refresh

End If

Exit Sub
c:
    MsgBox "File Attribute is Read Only You can not Delete it..."

Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
End If
On Error Resume Next
    Unload Form12
Exit Sub
End Sub

Private Sub Delmenu2_Click()
p = MsgBox("Are you Sure you want to Delete File", vbYesNo, "Delete ...?")

If p = 6 Then

If Len(folder_view.File1.FileName) <> 0 Then

On Error GoTo c

If Len(folder_view.Dir1.Path) <> 3 Then

fo.DeleteFile folder_view.Dir1.Path & "\" & folder_view.File1.FileName
folder_view.File1.Refresh
Else

fo.DeleteFile folder_view.Dir1.Path & folder_view.File1.FileName
folder_view.File1.Refresh

End If

Exit Sub
c:
    MsgBox "File Attribute is Read Only You can not Delete it..."

Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
End If

On Error Resume Next
    Unload Form12
Exit Sub
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Green_Click()
MDIForm1.Green.Checked = True
MDIForm1.Red.Checked = False
MDIForm1.blue.Checked = False
MDIForm1.choice.Checked = False

folder_view.Dir1.BackColor = &H808000
folder_view.File1.BackColor = &H808000
End Sub





Private Sub infomenu_Click()
If Len(folder_view.File1.FileName) <> 0 Then
    Form4.Show
Else
    MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If

End Sub

Private Sub MDIForm_Load()
MDIForm1.Green.Checked = True
Load Form13
Form13.Visible = False
Combo1.Text = "All Picture Files"
folder_view.Top = 0
folder_view.Left = 0
folder_view.Width = 2340
folder_view.Height = MDIForm1.Height - 2000
Form1.Show
Form1.Top = 0
Form1.Left = folder_view.Width + 200
Form1.Height = MDIForm1.Height - 2000
Form1.Width = MDIForm1.Width - folder_view.Width - 400
Form1.VScroll1.Left = Form1.Width - 350
Form1.VScroll1.Height = Form1.Height - 350
Form1.HScroll1.Top = Form1.Height - 350
Form1.HScroll1.Width = Form1.Width - 350
folder_view.Drive1.Drive = "c:\"
folder_view.Dir1.Path = "c:\"
folder_view.File1.Height = MDIForm1.Height - 5450

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub


Private Sub moveto_Click()
If Len(folder_view.File1.FileName) <> 0 Then
Form7.Show
Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
On Error Resume Next
    Unload Form12
Exit Sub
End Sub

Private Sub opendir_Click()
Form3.Show
End Sub

Private Sub Openfile_Click()
Form2.Show
End Sub

Private Sub pppmenu_Click()
    Form1.OLE1.DoVerb (0)
End Sub

Private Sub Red_Click()
Red.Checked = True
Green.Checked = False
blue.Checked = False
choice.Checked = False
folder_view.Dir1.BackColor = &H404080
folder_view.File1.BackColor = &H404080
End Sub



Private Sub remenu_Click()
If Len(folder_view.File1.FileName) <> 0 Then
Form5.Show
Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
On Error Resume Next
    Unload Form12
Exit Sub
End Sub

Private Sub renamemnu_Click()
If Len(folder_view.File1.FileName) <> 0 Then
Form5.Show
Else
MsgBox "You have not selected any file...", vbOKOnly, "Select any file..."
End If
On Error Resume Next
    Unload Form12
Exit Sub
End Sub



Private Sub sdsettings_Click()
Form10.Show
End Sub






Private Sub setwallpaper_Click()
Dim wallp As String
Dim e As String
Dim f As New FileSystemObject
wallp = Clear

If Len(folder_view.Dir1.Path) <> 3 Then
    wallp = folder_view.File1.Path & "\" & folder_view.File1.FileName
Else
    wallp = folder_view.File1.Path & folder_view.File1.FileName
End If
           
e = f.GetExtensionName(wallp)
If e = "gif" Or e = "GIF" Then
    MsgBox "Invalid File...Select BMP OR JPEG File", vbOKOnly Or vbInformation
    GoTo Y:
    
End If

X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, wallp, 2)
Y:
End Sub

Private Sub tset_Click()
Unload Form12
Form13.Show
End Sub

Private Sub up_Click()
    path_file = folder_view.Dir1.Path
If Len(path_file) = 3 Then
    MsgBox "You are in the root Drive you can not go back...", , "Can not go Back"
Else
    For i = 1 To Len(path_file)
        c = Mid(path_file, i, 1)
        If c = "\" Then
            pos = i
        End If
    Next
    
    folder_view.Dir1.Path = Mid(path_file, 1, pos)
End If

End Sub

