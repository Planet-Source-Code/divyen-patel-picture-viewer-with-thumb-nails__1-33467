VERSION 5.00
Begin VB.Form Form12 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Thumbnail View"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11700
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form12.frx":030A
   ScaleHeight     =   8280
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   10200
      Picture         =   "Form12.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   10200
      Picture         =   "Form12.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton Command1 
         Caption         =   "Return To Large View"
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Click On the Image for large Preview ..."
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "One Screen Up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8880
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "One Screen Down"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8880
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Click On Return To large View to goto main window ...."
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Main Window will not enabled untill you Close this window ..."
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   -120
      TabIndex        =   6
      Top             =   1440
      Width           =   11775
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         DragMode        =   1  'Automatic
         Height          =   1455
         Index           =   0
         Left            =   240
         MouseIcon       =   "Form12.frx":0FD0
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         ToolTipText     =   "Click Here For Large View"
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim c As Integer
Dim te As Integer
Dim m As Integer
Dim fh As Integer
Dim fw As Integer



Private Sub Command1_Click()
'Unload Form12
Form12.Hide
MDIForm1.Enabled = True
MDIForm1.Show
End Sub

Private Sub Command2_Click()
'For i = 0 To c - 1
'   Image1(i).Top = Image1(i).Top + 100
'Next

If Frame2.Top < 1440 Then
    Frame2.Top = Frame2.Top + 1455
End If



End Sub

Private Sub Command3_Click()
'For i = 0 To c - 1
'    Image1(i).Top = Image1(i).Top - 100
'Next
If Frame2.Top + Frame2.Height > 8285 Then
    Frame2.Top = Frame2.Top - 1455
End If


End Sub

Private Sub Form_Load()
m = 8
Form14.ProgressBar1.Max = 100
Form14.ProgressBar1.Value = 0
Form14.Label4.Caption = folder_view.File1.ListCount
Form14.Visible = True

If Form13.Slider2.Value = 5 Then
    m = m - 1
End If

If Form13.Slider2.Value = 7 Then
    m = m - 3
End If

If Form13.Slider2.Value >= 8 And Form13.Slider2.Value <= 9 Then
    m = m - 4
End If

If Form13.Slider2.Value = 10 Then
    m = m - 5
End If



Image1(0).Height = Form13.Image1.Height
Image1(0).Width = Form13.Image1.Width

fh = Image1(0).Height
fw = Image1(0).Width

MDIForm1.Enabled = False
folder_view.File1.Refresh
c = folder_view.File1.ListCount
If Len(folder_view.Dir1.Path) <> 3 Then

    Image1(0).Picture = LoadPicture(folder_view.File1.Path & "\" & folder_view.File1.List(o))
    Image1(0).Visible = True
        
        Text1(0).Left = Image1(0).Left
        Text1(0).Top = Image1(0).Top + Image1(0).Height
        Text1(0).Width = Image1(0).Width
        Text1(0).Text = folder_view.File1.List(0)
        
        Text1(0).Visible = True
    te = 1
    
For i = 1 To c - 1
    DoEvents
    If te <> m Then
        Load Image1(i)
        Image1(i).Left = Image1(i - 1).Left + Image1(i - 1).Width + 100
        Image1(i).Top = Image1(i - 1).Top
        Load Text1(i)
        Text1(i).Left = Image1(i).Left
        Text1(i).Top = Image1(i).Top + Image1(i).Height
        Text1(i).Width = Image1(i).Width
        Text1(i).Text = folder_view.File1.List(i)
        Text1(i).Visible = True
        
        Image1(i).Picture = LoadPicture(folder_view.File1.Path & "\" & folder_view.File1.List(i))
        If Image1(i).Picture.Height < Image1(i).Height And Image1(i).Picture.Width < Image1(i).Width Then
            Image1(i).Stretch = False
            Image1(i).Width = fw
            Image1(i).Height = fh
        End If
        
        Image1(i).Visible = True
        
    Else
        Load Image1(i)
        Image1(i).Top = Image1(i - 1).Top + Image1(i - 1).Height + 350
        Image1(i).Left = Image1(0).Left
        Image1(i).Picture = LoadPicture(folder_view.File1.Path & "\" & folder_view.File1.List(i))
        Load Text1(i)
        Text1(i).Left = Image1(i).Left
        Text1(i).Top = Image1(i).Top + Image1(i).Height
        Text1(i).Width = Image1(i).Width
        Text1(i).Text = folder_view.File1.List(i)
        Text1(i).Visible = True
        If Image1(i).Picture.Height < Image1(i).Height And Image1(i).Picture.Width < Image1(i).Width Then
            Image1(i).Stretch = False
            Image1(i).Width = fw
            Image1(i).Height = fh
        End If
        Image1(i).Visible = True
        Image1(i).Visible = True
        
        te = 0
    End If
    
    te = te + 1
    Form14.Label2.Caption = i
    Form14.ProgressBar1.Value = Int((i / (c - 1)) * 100)
    Form14.Label5.Caption = Int((i / (c - 1)) * 100) & " %"
    If i = c - 1 Then
        Unload Form14
    End If
    
Next



Frame2.Height = Image1(c - 1).Top + Image1(c - 1).Height + Text1(c - 1).Height





'Frame2.Width = Image1(c - 1).Left + Image1(c - 1).Width

'Form12.Height = Image1(c - 1).Top + Image1(c - 1).Height
'Form12.Width = Image1(c - 1).Left + Image1(c - 1).Width

Else
Image1(0).Picture = LoadPicture(folder_view.File1.Path & folder_view.File1.List(o))
    Image1(0).Visible = True
    te = 1
Text1(0).Left = Image1(0).Left
        Text1(0).Top = Image1(0).Top + Image1(0).Height
        Text1(0).Width = Image1(0).Width
        Text1(0).Text = folder_view.File1.List(0)
        
        Text1(0).Visible = True
For i = 1 To c - 1
 DoEvents
    If te <> m Then
        Load Image1(i)
        Image1(i).Left = Image1(i - 1).Left + Image1(i - 1).Width + 100
        Image1(i).Top = Image1(i - 1).Top
        Image1(i).Picture = LoadPicture(folder_view.File1.Path & folder_view.File1.List(i))
        Load Text1(i)
        Text1(i).Left = Image1(i).Left
        Text1(i).Top = Image1(i).Top + Image1(i).Height
        Text1(i).Width = Image1(i).Width
        Text1(i).Text = folder_view.File1.List(i)
        Text1(i).Visible = True
        If Image1(i).Picture.Height < Image1(i).Height And Image1(i).Picture.Width < Image1(i).Width Then
            Image1(i).Stretch = False
            Image1(i).Width = fw
            Image1(i).Height = fh
        End If
        
        Image1(i).Visible = True
    Else
        Load Image1(i)
        Image1(i).Top = Image1(i - 1).Top + Image1(i - 1).Height + 350
        Image1(i).Left = Image1(0).Left
        Image1(i).Picture = LoadPicture(folder_view.File1.Path & folder_view.File1.List(i))
        Load Text1(i)
        Text1(i).Left = Image1(i).Left
        Text1(i).Top = Image1(i).Top + Image1(i).Height
        Text1(i).Width = Image1(i).Width
        Text1(i).Text = folder_view.File1.List(i)
        Text1(i).Visible = True
        If Image1(i).Picture.Height < Image1(i).Height And Image1(i).Picture.Width < Image1(i).Width Then
            Image1(i).Stretch = False
            Image1(i).Width = fw
            Image1(i).Height = fh
        End If
        Image1(i).Visible = True
        te = 0
    End If
    
    te = te + 1
    Form14.Label2.Caption = i
    Form14.ProgressBar1.Value = Int((i / (c - 1)) * 100)
    Form14.Label5.Caption = Int((i / (c - 1)) * 100) & " %"
    
        
    


Next
Unload Form14

Frame2.Height = Image1(c - 1).Top + Image1(c - 1).Height + Text1(c - 1).Height

'Frame2.Width = Image1(c - 1).Left + Image1(c - 1).Width

'Form12.Height = Image1(c - 1).Top + Image1(c - 1).Height
'Form12.Width = Image1(c - 1).Left + Image1(c - 1).Width


End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.Enabled = True
    MDIForm1.Show
End Sub


Private Sub Image1_Click(Index As Integer)
    MDIForm1.Enabled = True
    
    
    Form1.Image1.Picture = LoadPicture(folder_view.Dir1.Path & "\" & folder_view.File1.List(Index))
    Form1.HScroll1.Max = 0
Form1.VScroll1.Max = 0

If Form1.Image1.Width > Form1.HScroll1.Width Then
    Form1.HScroll1.Max = Form1.Image1.Width - Form1.HScroll1.Width
    folder_view.File1.ListIndex = Index
End If

If Form1.Image1.Height > Form1.VScroll1.Height Then
    Form1.VScroll1.Max = Form1.Image1.Height - Form1.VScroll1.Height
    folder_view.File1.ListIndex = Index
End If

Form1.HScroll1.LargeChange = 1000
Form1.VScroll1.LargeChange = 1000

Form1.HScroll1.SmallChange = 500
Form1.VScroll1.SmallChange = 500
'    Unload fomr12
Form12.Hide
MDIForm1.Show
End Sub

Private Sub Image1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    MDIForm1.Enabled = True
    If Len(folder_view.Dir1.Path) <> 3 Then
        Form1.Image1.Picture = LoadPicture(folder_view.Dir1.Path & "\" & folder_view.File1.List(Index))
        folder_view.File1.ListIndex = Index
    Else
        Form1.Image1.Picture = LoadPicture(folder_view.Dir1.Path & folder_view.File1.List(Index))
        folder_view.File1.ListIndex = Index
    End If
    
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
    'Unload Form12
    Form12.Hide
MDIForm1.Show
End Sub


