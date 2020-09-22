VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Info"
   ClientHeight    =   5955
   ClientLeft      =   2490
   ClientTop       =   1950
   ClientWidth     =   7470
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7470
   Begin VB.Frame Frame1 
      Caption         =   "File Property"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   7215
      Begin VB.Label prop1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   8
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "Picture Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "File Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "File Size (In bytes )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   "Last Modified Date and Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label prop1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label prop1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label prop1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   1
         Top             =   1800
         Width           =   3615
      End
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   1680
      Picture         =   "Form4.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3840
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fo As New FileSystemObject
Dim mb As Variant
Dim ext As String
Private Sub Form_Load()

Form4.Left = 2425
Form4.Top = 420


Form4.BackColor = RGB(212, 212, 212)
Form4.Frame1.BackColor = RGB(212, 212, 212)

folder_view.Refresh
Form4.Refresh

folder_view.Refresh
prop1(0).Caption = folder_view.File1.FileName

If Len(folder_view.Dir1.Path) <> 3 Then
    mb = FileLen(folder_view.Dir1.Path & "\" & folder_view.File1.FileName)
    Form4.Image2.Picture = LoadPicture(folder_view.Dir1.Path & "\" & folder_view.File1.FileName)
    
Else
    mb = FileLen(folder_view.Dir1.Path & folder_view.File1.FileName)
    Form4.Image2.Picture = LoadPicture(folder_view.Dir1.Path & folder_view.File1.FileName)
End If

prop1(1).Caption = mb & " Bytes "

If Len(folder_view.Dir1.Path) <> 3 Then
    prop1(2).Caption = FileDateTime(folder_view.Dir1.Path & "\" & folder_view.File1.FileName)
    ext = fo.GetExtensionName(folder_view.Dir1.Path & "\" & folder_view.File1.FileName)
Else
    prop1(2).Caption = FileDateTime(folder_view.Dir1.Path & folder_view.File1.FileName)
    ext = fo.GetExtensionName(folder_view.Dir1.Path & folder_view.File1.FileName)
End If

'ext = Mid(prop1(0).Caption, Len(prop1(0).Caption) - 3, Len(prop1(0).Caption))



If ext = "bmp" Or ext = "BMP" Then
    prop1(3).Caption = "Bitmap File"
ElseIf ext = "GIF" Or ext = "gif" Then
    prop1(3).Caption = "Gif Compressed File"
ElseIf ext = "jpg" Or ext = "JPG" Then
    prop1(3).Caption = " jpeg File"
End If




End Sub



