VERSION 5.00
Begin VB.Form frmFlatIB4 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3105
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6450
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6450
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picX 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6045
      Picture         =   "frmFlatIB4.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1425
      TabIndex        =   0
      Top             =   1590
      Width           =   4020
   End
   Begin VB.PictureBox picHide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2520
      Picture         =   "frmFlatIB4.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2700
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3150
      Left            =   -15
      ScaleHeight     =   3120
      ScaleWidth      =   1170
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1200
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   345
         ScaleHeight     =   540
         ScaleWidth      =   465
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   300
         Width           =   500
      End
   End
   Begin AxMessageBox.axButton axButton2 
      Height          =   420
      Left            =   3045
      TabIndex        =   3
      Top             =   2475
      Width           =   1440
      _ExtentX        =   2540
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
      MICON           =   "frmFlatIB4.frx":0B14
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
      Left            =   4650
      TabIndex        =   2
      Top             =   2475
      Width           =   1440
      _ExtentX        =   2540
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
      MICON           =   "frmFlatIB4.frx":0B30
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1410
      TabIndex        =   8
      Top             =   285
      Width           =   1200
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmFlatIB4.frx":0B4C
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
      Height          =   615
      Left            =   1455
      TabIndex        =   1
      Top             =   900
      Width           =   4620
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFlatIB4"
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
If lblTitle.Width > 4500 Then
  Me.Width = lblTitle.Left + lblTitle.Width + 1000
End If

lblMessage.Move 1450, 900, (Me.Width - lblMessage.Left) - 500
lblMessage.AutoSize = True

If lblMessage.Height > 615 Then
  Me.Height = lblMessage.Top + lblMessage.Height + Text1.Height + 800
  Text1.Move 165, lblMessage.Top + lblMessage.Height + 100, (Me.Width - Text1.Left) - 600
End If

Text1.Width = (Me.Width - Text1.Left) - 600 'lblMessage.Width - 500
'picBar.Width = Me.Width + 30
'picX.Left = picBar.Width - 380
'picX.Picture = picHide.Picture
'axButton1.Move Me.Width - 1650, Me.Height - 580
'axButton2.Move Me.Width - 3250, Me.Height - 580
picBar.Width = 1200 'Me.Width + 30
picX.Left = Me.Width - 400 ' picBar.Width - 380
picX.Picture = picHide.Picture
axButton1.Move Me.Width - 1750, Me.Height - 700
axButton2.Move Me.Width - 3350, Me.Height - 700


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

