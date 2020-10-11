VERSION 5.00
Begin VB.Form frmIB3D 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5460
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   1170
      Width           =   3660
   End
   Begin VB.PictureBox picX 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5160
      Picture         =   "frmIB3D.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   30
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   120
      Picture         =   "frmIB3D.frx":058A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   510
   End
   Begin AxMessageBox.axButton axButton2 
      Height          =   420
      Left            =   4155
      TabIndex        =   2
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmIB3D.frx":0E54
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
      Left            =   4155
      TabIndex        =   1
      Top             =   1110
      Width           =   1200
      _ExtentX        =   2117
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmIB3D.frx":0E70
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
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   15
      Width           =   675
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje...                                        Mensaje...                                        Mensaje..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   585
      Index           =   1
      Left            =   780
      TabIndex        =   5
      Top             =   480
      Width           =   3285
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje...                                        Mensaje...                                        Mensaje..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Index           =   0
      Left            =   795
      TabIndex        =   4
      Top             =   495
      Width           =   3285
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   1
      X1              =   210
      X2              =   5460
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   210
      X2              =   5460
      Y1              =   330
      Y2              =   330
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
      Index           =   0
      Left            =   255
      TabIndex        =   8
      Top             =   30
      Width           =   675
   End
End
Attribute VB_Name = "frmIB3D"
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
End Sub

Private Sub Form_Resize()

If lblTitle(1).Width > 4850 Then
  Me.Width = lblTitle(1).Left + lblTitle(1).Width + 500
End If
'795,495 / 780,480
lblMessage(1).Move 780, 480, Me.Width - (780 + axButton1.Width + 300)
lblMessage(0).Move 795, 495, Me.Width - (795 + axButton1.Width + 300)
lblMessage(1).AutoSize = True
lblMessage(0).AutoSize = True

If lblMessage(0).Height > 590 Then
  Me.Height = lblMessage(1).Top + lblMessage(1).Height + 700
  Text1.Move 230, lblMessage(1).Top + lblMessage(1).Height + 150
End If

Text1.Width = lblMessage(0).Width + picIcon.Width
picX.Move Me.ScaleWidth - 300, 30
Line1(0).X2 = Me.ScaleWidth - 90
Line1(1).X2 = Me.ScaleWidth - 90
axButton1.Move Me.ScaleWidth - 1350, Me.ScaleHeight - 550
axButton2.Move Me.ScaleWidth - 1350, Me.ScaleHeight - 1100
End Sub

Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormBar Me.hwnd
End Sub

Private Sub picBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  MoveFormBar Me.hwnd
End If
End Sub

Private Sub picX_Click()
asClicked = 2
Unload Me
End Sub

