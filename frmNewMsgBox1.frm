VERSION 5.00
Begin VB.Form frmNS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
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
   ForeColor       =   &H00000000&
   LinkTopic       =   "frmMessage"
   ScaleHeight     =   2055
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   120
      ScaleHeight     =   525
      ScaleWidth      =   525
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   75
      Width           =   525
   End
   Begin VB.PictureBox picX 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E2FC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5775
      Picture         =   "frmNewMsgBox1.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   105
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   150
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   900
      Width           =   525
   End
   Begin AxMessageBox.axButton axButton2 
      Height          =   375
      Left            =   3210
      TabIndex        =   2
      Top             =   1590
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      ButtonType      =   5
      Caption         =   "axButton2"
      Enabled         =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmNewMsgBox1.frx":058A
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
      Height          =   375
      Left            =   4695
      TabIndex        =   1
      Top             =   1590
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      ButtonType      =   5
      Caption         =   "axButton1"
      Enabled         =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmNewMsgBox1.frx":05A6
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
      Caption         =   "Message Title Extense String"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   810
      TabIndex        =   5
      Top             =   120
      Width           =   4605
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mensaje........................................................................."
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
      Left            =   840
      TabIndex        =   0
      Top             =   855
      Width           =   5115
   End
End
Attribute VB_Name = "frmNS"
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
Me.BorderStyle = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormBar Me.hwnd
End Sub

Private Sub Form_Resize()
Dim MsgAncho As Integer
Dim TtlAncho As Integer
Dim CtrlAncho As Integer

MsgAncho = lblMessage.Left + lblMessage.Width + 300
TtlAncho = lblTitle.Left + lblTitle.Width + 300
CtrlAncho = 4300

If MsgAncho > TtlAncho And MsgAncho > CtrlAncho Then
  Me.Width = lblMessage.Left + lblMessage.Width + 700
ElseIf TtlAncho > MsgAncho And TtlAncho > CtrlAncho Then
  Me.Width = lblTitle.Left + lblTitle.Width + 1000
Else
  Me.Width = CtrlAncho
End If


Me.Height = lblMessage.Top + lblMessage.Height + 670

picX.Left = Me.Width - 380
axButton1.Move Me.Width - 1450, Me.Height - 540
axButton2.Move Me.Width - 2940, Me.Height - 540
'-----------------------
Call CreateForm(picX.BackColor)
End Sub

Public Sub CreateForm(sColor As Long)

Call CreateTitle(Me, 9, sColor)
Call RoundCorner(Me, 9)
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormBar Me.hwnd
End Sub

Private Sub picV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormBar Me.hwnd
End Sub

Private Sub picX_Click()
asClicked = 2
Unload Me
End Sub

