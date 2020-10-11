VERSION 5.00
Begin VB.Form frmIcons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmIcons"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
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
   Icon            =   "frmIcons.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicOK 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   735
      Picture         =   "frmIcons.frx":000C
      ScaleHeight     =   585
      ScaleWidth      =   570
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1545
      Width           =   570
   End
   Begin VB.PictureBox PicError 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   75
      Picture         =   "frmIcons.frx":11FA
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1545
      Width           =   540
   End
   Begin VB.PictureBox PicInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1395
      Picture         =   "frmIcons.frx":216C
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1545
      Width           =   495
   End
   Begin VB.PictureBox PicQuestion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2055
      Picture         =   "frmIcons.frx":2EF6
      ScaleHeight     =   495
      ScaleWidth      =   570
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1545
      Width           =   570
   End
   Begin VB.PictureBox PicAlert 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2715
      Picture         =   "frmIcons.frx":3E2C
      ScaleHeight     =   510
      ScaleWidth      =   525
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1545
      Width           =   525
   End
   Begin VB.PictureBox pHelp 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   6840
      Picture         =   "frmIcons.frx":4CC6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pSupport2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   6090
      Picture         =   "frmIcons.frx":5990
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pStar 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   5340
      Picture         =   "frmIcons.frx":625A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pSupport1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   4590
      Picture         =   "frmIcons.frx":6B24
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pUser 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3855
      Picture         =   "frmIcons.frx":80A6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pKeys 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3105
      Picture         =   "frmIcons.frx":8970
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pBomb 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2355
      Picture         =   "frmIcons.frx":923A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pRay 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1605
      Picture         =   "frmIcons.frx":9B04
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pBug 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   870
      Picture         =   "frmIcons.frx":A3CE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pBlocks 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "frmIcons.frx":AC98
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   855
      Width           =   540
   End
   Begin VB.PictureBox pSmile 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   6840
      Picture         =   "frmIcons.frx":B562
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pNote 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   6088
      Picture         =   "frmIcons.frx":BE2C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pStop2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   5342
      Picture         =   "frmIcons.frx":C6F6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pLock 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   4596
      Picture         =   "frmIcons.frx":D3C0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pTips 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3850
      Picture         =   "frmIcons.frx":DC8A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pStop1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3104
      Picture         =   "frmIcons.frx":E554
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pAlert 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2358
      Picture         =   "frmIcons.frx":EE1E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pInfo 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1612
      Picture         =   "frmIcons.frx":F6E8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pCancel 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   866
      Picture         =   "frmIcons.frx":FFB2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox pOK 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "frmIcons.frx":1087C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
