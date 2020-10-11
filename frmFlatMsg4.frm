VERSION 5.00
Begin VB.Form frmFlat4 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7125
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
   ScaleHeight     =   3375
   ScaleWidth      =   7125
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picX 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6750
      Picture         =   "frmFlatMsg4.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   240
   End
   Begin VB.PictureBox picHide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6570
      Picture         =   "frmFlatMsg4.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   -15
      ScaleHeight     =   3375
      ScaleWidth      =   1170
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1200
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   345
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   315
         Width           =   500
      End
   End
   Begin AxMessageBox.axButton axButton2 
      Height          =   420
      Left            =   3780
      TabIndex        =   2
      Top             =   2730
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
      MICON           =   "frmFlatMsg4.frx":0B14
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
      Left            =   5400
      TabIndex        =   1
      Top             =   2730
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
      MICON           =   "frmFlatMsg4.frx":0B30
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
      Caption         =   "Titulo 00000000000000000000000000"
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
      Left            =   1500
      TabIndex        =   6
      Top             =   360
      Width           =   5040
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mensaje.........................................................................."
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
      Left            =   1530
      TabIndex        =   0
      Top             =   900
      Width           =   5175
   End
End
Attribute VB_Name = "frmFlat4"
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
Dim MsgAncho As Integer
Dim TtlAncho As Integer
Dim CtrlAncho As Integer

MsgAncho = lblMessage.Left + lblMessage.Width '+ 450
TtlAncho = lblTitle.Left + lblTitle.Width '+ 450
CtrlAncho = 6850

If MsgAncho > TtlAncho And MsgAncho > CtrlAncho Then
  Me.Width = lblMessage.Left + lblMessage.Width + 450
ElseIf TtlAncho > MsgAncho And TtlAncho > CtrlAncho Then
  Me.Width = lblTitle.Left + lblTitle.Width + 1000
Else
  Me.Width = CtrlAncho
End If

Me.Height = lblMessage.Top + lblMessage.Height + 700

picBar.Width = 1200 'Me.Width + 30
picX.Left = Me.Width - 400 ' picBar.Width - 380
picX.Picture = picHide.Picture
axButton1.Move Me.Width - 1750, Me.Height - 600
axButton2.Move Me.Width - 3350, Me.Height - 600

'-----------
'Me.Height = lblMessage.Top + lblMessage.Height + 650
'axButton1.Top = Me.Height - 590
'axButton2.Top = Me.Height - 590
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

