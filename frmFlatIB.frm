VERSION 5.00
Begin VB.Form frmFlatIB 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1890
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6525
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   945
      TabIndex        =   0
      Top             =   975
      Width           =   5310
   End
   Begin VB.PictureBox picHide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      Picture         =   "frmFlatIB.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   465
      ScaleHeight     =   315
      ScaleWidth      =   6060
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   6060
      Begin VB.PictureBox picX 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5760
         Picture         =   "frmFlatIB.frx":058A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   30
         Width           =   240
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   15
         Width           =   1290
      End
   End
   Begin AxMessageBox.axButton axButton2 
      Height          =   420
      Left            =   3450
      TabIndex        =   4
      Top             =   1380
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      ButtonType      =   7
      Caption         =   "axButton2"
      Enabled         =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFlatIB.frx":0B14
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AxMessageBox.axButton axButton1 
      Height          =   420
      Left            =   4905
      TabIndex        =   3
      Top             =   1380
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      ButtonType      =   7
      Caption         =   "axButton1"
      Enabled         =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFlatIB.frx":0B30
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   500
      Left            =   210
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   500
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmFlatIB.frx":0B4C
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   945
      TabIndex        =   2
      Top             =   465
      Width           =   5310
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFlatIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub axButton1_Click()
asClicked = 1
strResp = Text1.Text
Unload Me
End Sub

Private Sub axButton2_Click()
asClicked = 2
Unload Me
End Sub

Private Sub Form_Load()
picIcon.BorderStyle = 0
lblMessage.BorderStyle = 0
End Sub

Private Sub Form_Resize()
If lblTitle.Width > 5500 Then
  Me.Width = lblTitle.Left + lblTitle.Width + 1000
End If

lblMessage.Move 945, 465, Me.Width - (945 + 300)
lblMessage.AutoSize = True

If lblMessage.Height > 425 Then
  Me.Height = lblMessage.Top + lblMessage.Height + Text1.Height + 800
  Text1.Move 945, lblMessage.Top + lblMessage.Height + 120
End If

Text1.Width = lblMessage.Width
picBar.Width = Me.ScaleWidth - picBar.Left
picX.Left = picBar.Width - 300
picX.Picture = picHide.Picture
axButton1.Move Me.Width - 1650, Me.Height - 550
axButton2.Move Me.Width - 3105, Me.Height - 550

End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormBar Me.hwnd
End Sub

Private Sub picBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormBar Me.hwnd
End Sub

Private Sub picX_Click()
asClicked = 2
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  axButton1_Click
ElseIf KeyAscii = vbKeyEscape Then
  axButton2_Click
End If
End Sub
