VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test AxMessageBox DLL"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "FlatStyle4"
      Height          =   315
      Left            =   4905
      TabIndex        =   24
      Top             =   2205
      Width           =   1020
   End
   Begin VB.Frame Frame2 
      Caption         =   "AxInputBox"
      Height          =   1875
      Left            =   45
      TabIndex        =   15
      Top             =   2865
      Width           =   7875
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   3150
         TabIndex        =   22
         Text            =   "Esta es una Pregunta de Prueba"
         Top             =   495
         Width           =   3285
      End
      Begin VB.CommandButton Command13 
         Caption         =   "InputBox"
         Height          =   600
         Left            =   6675
         TabIndex        =   19
         Top             =   465
         Width           =   990
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   165
         TabIndex        =   18
         Top             =   495
         Width           =   1650
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   3150
         TabIndex        =   17
         Top             =   1185
         Width           =   3285
      End
      Begin VB.ListBox List2 
         Height          =   1035
         Left            =   1905
         TabIndex        =   16
         Top             =   495
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Texto a Mostrar en el InputBox"
         Height          =   195
         Left            =   3135
         TabIndex        =   23
         Top             =   255
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Texto Retornado"
         Height          =   195
         Left            =   3135
         TabIndex        =   21
         Top             =   945
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estilo                                Icono"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   285
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AxMsgBox"
      Height          =   2775
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7875
      Begin VB.TextBox txtTitle 
         Height          =   345
         Left            =   150
         TabIndex        =   12
         Top             =   450
         Width           =   4485
      End
      Begin VB.TextBox txtMsg 
         Height          =   1575
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1050
         Width           =   4500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "New Style OK"
         Height          =   375
         Left            =   6000
         TabIndex        =   10
         Top             =   405
         Width           =   1635
      End
      Begin VB.CommandButton Command8 
         Caption         =   "New Style Alert"
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   817
         Width           =   1635
      End
      Begin VB.CommandButton Command9 
         Caption         =   "New Style Error"
         Height          =   375
         Left            =   6000
         TabIndex        =   8
         Top             =   1229
         Width           =   1635
      End
      Begin VB.CommandButton Command10 
         Caption         =   "New Style Info"
         Height          =   375
         Left            =   6000
         TabIndex        =   7
         Top             =   1641
         Width           =   1635
      End
      Begin VB.CommandButton Command11 
         Caption         =   "New Style Question"
         Height          =   375
         Left            =   6000
         TabIndex        =   6
         Top             =   2055
         Width           =   1635
      End
      Begin VB.CommandButton Command1 
         Caption         =   "FlatStyle1"
         Height          =   315
         Left            =   4860
         TabIndex        =   5
         Top             =   1110
         Width           =   1020
      End
      Begin VB.CommandButton Command4 
         Caption         =   "FlatStyle2"
         Height          =   315
         Left            =   4860
         TabIndex        =   4
         Top             =   1455
         Width           =   1020
      End
      Begin VB.CommandButton Command5 
         Caption         =   "FlatStyle3"
         Height          =   315
         Left            =   4860
         TabIndex        =   3
         Top             =   1815
         Width           =   1020
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Classic"
         Height          =   315
         Left            =   4860
         TabIndex        =   2
         Top             =   405
         Width           =   1020
      End
      Begin VB.CommandButton Command7 
         Caption         =   "3D"
         Height          =   315
         Left            =   4860
         TabIndex        =   1
         Top             =   750
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         Height          =   195
         Left            =   225
         TabIndex        =   14
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   840
         Width           =   600
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xBox As New AxMsgBox

Dim sCadena As String, sTitle As String
Dim C As Integer


Private Sub Command1_Click()
C = xBox.AxMsgBox(FlatStyle1, txtMsg, txtTitle, icAlert, bAcceptCancel)

MsgBox "Presionado Boton " & C
End Sub

Private Sub Command13_Click()
Text2.Text = xBox.AxInputBox(List1.ListIndex, Text1.Text, txtTitle, List2.ListIndex, bAcceptCancel)
End Sub

Private Sub Command2_Click()
C = xBox.AxMsgBox(FlatStyle4, txtMsg, txtTitle, icok, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Command4_Click()
C = xBox.AxMsgBox(FlatStyle2, txtMsg, txtTitle, icbomb, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Command5_Click()
C = xBox.AxMsgBox(FlatStyle3, txtMsg, txtTitle, icok, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Command6_Click()
C = xBox.AxMsgBox(Classic, txtMsg, txtTitle, icbug, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Command7_Click()
C = xBox.AxMsgBox(Style3D, txtMsg, txtTitle, icNote, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Command8_Click()
C = xBox.AxMsgBox(NSAlert, txtMsg, txtTitle, icsupport2, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Command9_Click()
C = xBox.AxMsgBox(NSError, txtMsg, txtTitle, ickeys, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Command10_Click()
C = xBox.AxMsgBox(NSInfo, txtMsg, txtTitle, icuser, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Command11_Click()
C = xBox.AxMsgBox(NSQuestion, txtMsg, txtTitle, icsupport1, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Command3_Click()
C = xBox.AxMsgBox(NSOK, txtMsg, txtTitle, icray, bAcceptCancel)
MsgBox "Presionado Boton " & C

End Sub

Private Sub Form_Load()
sTitle = "Etiam iaculis dui non commodo elementum."
sCadena = "Lorem ipsum dolor sit amet, consectetur adipiscing elit." & vbCrLf & _
          "Fusce vel scelerisque mauris. Donec eu odio et sem vulputate egestas." & vbCrLf & _
          "Vivamus tincidunt elit ipsum, vitae ultricies turpis semper sit amet." & vbCrLf & _
          "Donec congue velit non neque sodales, sit amet sagittis felis pellentesque." & vbCrLf & _
          "Ut nec mi felis. Duis a dignissim dolor, non venenatis erat." & vbCrLf & _
          "Curabitur lobortis, odio ac pellentesque egestas, dui nisi fringilla nibh, sed tincidunt ligula felis vitae elit." & vbCrLf & _
          "Sed tristique a mi non congue."
          
txtTitle.Text = sTitle
txtMsg.Text = sCadena

With List1
    .AddItem "FlatStyle1"
    .AddItem "FlatStyle2"
    .AddItem "FlatStyle3"
    .AddItem "FlatStyle4"
    .AddItem "Classic"
    .AddItem "Style3D"
    .AddItem "NSError"
    .AddItem "NSOK"
    .AddItem "NSInfo"
    .AddItem "NSQuestion"
    .AddItem "NSAlert"
End With

With List2
    .AddItem "icNoIcon"
    .AddItem "icok"
    .AddItem "icCancel"
    .AddItem "icInfo"
    .AddItem "icHelp"
    .AddItem "icAlert"
    .AddItem "icStop1"
    .AddItem "icStop2"
    .AddItem "icTips"
    .AddItem "icLock"
    .AddItem "icNote"
    .AddItem "icSmile"
    .AddItem "icBlocks"
    .AddItem "icbug"
    .AddItem "icray"
    .AddItem "icbomb"
    .AddItem "ickeys"
    .AddItem "icuser"
    .AddItem "icsupport1"
    .AddItem "icsupport2"
    .AddItem "icStar"
End With

List1.ListIndex = 0
List2.ListIndex = 4
End Sub

