VERSION 5.00
Begin VB.Form frmFlat3 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1785
   ClientLeft      =   15
   ClientTop       =   15
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
   ScaleHeight     =   1785
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picHide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   30
      Picture         =   "frmFlatMsg3.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1545
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   -15
      ScaleHeight     =   540
      ScaleWidth      =   5445
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -15
      Width           =   5475
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   135
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   15
         Width           =   500
      End
      Begin VB.PictureBox picX 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5145
         Picture         =   "frmFlatMsg3.frx":058A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   60
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
         Left            =   825
         TabIndex        =   6
         Top             =   135
         Width           =   675
      End
   End
   Begin AxMessageBox.axButton axButton2 
      Height          =   420
      Left            =   4005
      TabIndex        =   2
      Top             =   690
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFlatMsg3.frx":0B14
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
      Left            =   4005
      TabIndex        =   1
      Top             =   1245
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFlatMsg3.frx":0B30
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
      Caption         =   "Mensaje................................................"
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
      Left            =   210
      TabIndex        =   0
      Top             =   660
      Width           =   3615
   End
End
Attribute VB_Name = "frmFlat3"
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

MinAncho = lblMessage.Left + lblMessage.Width + 1650
'lblTitle.Width > 4500 And
If lblTitle.Width > MinAncho Then
  Me.Width = lblTitle.Left + lblTitle.Width + 1000
Else
  Me.Width = MinAncho
End If

If lblMessage.Height > 1000 Then
  Me.Height = lblMessage.Top + lblMessage.Height + 290
End If

picBar.Width = Me.Width
picX.Left = picBar.Width - 380
picX.Picture = picHide.Picture
axButton1.Move Me.Width - 1500, Me.Height - 550
axButton2.Move Me.Width - 1500, Me.Height - 1050

'-----------------
'With lblMessage
'  Me.Height = .Top + .Height + 290
'End With
'axButton1.Top = Me.Height - 660
'axButton2.Top = Me.Height - 1170
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

