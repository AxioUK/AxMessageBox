VERSION 5.00
Begin VB.Form frmClassic 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5445
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
   LinkTopic       =   "frmMessage"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmClassic.frx":0000
   ScaleHeight     =   1575
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      Picture         =   "frmClassic.frx":058A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   525
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   15
      ScaleHeight     =   300
      ScaleWidth      =   5385
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -15
      Width           =   5415
      Begin VB.PictureBox picX 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5130
         Picture         =   "frmClassic.frx":0E54
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   45
         Width           =   240
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
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   150
         TabIndex        =   5
         Top             =   15
         Width           =   1890
      End
   End
   Begin AxMessageBox.axButton axButton2 
      Height          =   420
      Left            =   4185
      TabIndex        =   2
      Top             =   525
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   741
      ButtonType      =   3
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmClassic.frx":13DE
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
      Left            =   4185
      TabIndex        =   1
      Top             =   1035
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   741
      ButtonType      =   3
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmClassic.frx":13FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mensaje.........................................."
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
      Height          =   225
      Left            =   765
      TabIndex        =   0
      Top             =   465
      Width           =   3255
   End
End
Attribute VB_Name = "frmClassic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub axButton1_Click()
asClicked = 1
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
Dim MinAncho As Integer

MinAncho = lblMessage.Left + lblMessage.Width + 1515

'If lblTitle.Width > 4900 And lblTitle.Width > MinAncho Then
If lblTitle.Width > MinAncho Then
  Me.Width = lblTitle.Left + lblTitle.Width + 500
Else
  Me.Width = MinAncho
End If

If lblMessage.Height > 1005 Then
  Me.Height = lblMessage.Top + lblMessage.Height + 290
End If

picBar.Width = Me.ScaleWidth
picX.Left = picBar.Width - 300
picX.Picture = Me.Picture
axButton1.Move Me.ScaleWidth - 1300, Me.Height - 660
axButton2.Move Me.ScaleWidth - 1300, Me.Height - 1170

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

