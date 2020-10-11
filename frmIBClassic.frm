VERSION 5.00
Begin VB.Form frmIBClassic 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5475
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
   Picture         =   "frmIBClassic.frx":0000
   ScaleHeight     =   1680
   ScaleWidth      =   5475
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   180
      TabIndex        =   0
      Top             =   1200
      Width           =   3900
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      Picture         =   "frmIBClassic.frx":058A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Width           =   525
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   15
      ScaleHeight     =   300
      ScaleWidth      =   5460
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -15
      Width           =   5490
      Begin VB.PictureBox picX 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5160
         Picture         =   "frmIBClassic.frx":0E54
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   15
         Width           =   1890
      End
   End
   Begin AxMessageBox.axButton axButton2 
      Height          =   420
      Left            =   4230
      TabIndex        =   3
      Top             =   600
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmIBClassic.frx":13DE
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
      Left            =   4230
      TabIndex        =   2
      Top             =   1110
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmIBClassic.frx":13FA
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
      Caption         =   $"frmIBClassic.frx":1416
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
      Left            =   810
      TabIndex        =   1
      Top             =   435
      Width           =   3255
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmIBClassic"
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
If lblTitle.Width > 4900 Then
  Me.Width = lblTitle.Left + lblTitle.Width + 500
End If
If lblMessage.Height < 450 Then
  lblMessage.Height = 650
  Form_Resize
End If

lblMessage.Move 810, 435, Me.Width - (810 + axButton1.Width + 300)
lblMessage.AutoSize = True

If lblMessage.Height > 615 Then
  Me.Height = lblMessage.Top + lblMessage.Height + Text1.Height + 390
  Text1.Move 180, lblMessage.Top + lblMessage.Height + 200
End If

Text1.Width = lblMessage.Width + picIcon.Width
picBar.Width = Me.ScaleWidth
picX.Left = picBar.Width - 300
picX.Picture = Me.Picture
axButton1.Move Me.ScaleWidth - 1280, Me.Height - 600
axButton2.Move Me.ScaleWidth - 1280, Me.Height - 1070

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

