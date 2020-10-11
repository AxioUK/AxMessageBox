VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AxMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ClickedButton
  cButton1
  cButton2
End Enum

Public Enum MsgStyle
    [FlatStyle1] = 0
    [FlatStyle2] = 1
    [FlatStyle3] = 2
    [FlatStyle4] = 3
    [Classic] = 4
    [Style3D] = 5
    [NSError] = 6
    [NSOK] = 7
    [NSInfo] = 8
    [NSQuestion] = 9
    [NSAlert] = 10
End Enum

Public Enum StyleButtons
    bYesNo
    bAcceptCancel
    bOKOnly
    bAcceptOnly
End Enum

Public Enum sIcon
    [icNoIcon] = 0
    [icok] = 1
    [icCancel] = 2
    [icInfo] = 3
    [icHelp] = 4
    [icAlert] = 5
    [icStop1] = 6
    [icStop2] = 7
    [icTips] = 8
    [icLock] = 9
    [icNote] = 10
    [icSmile] = 11
    [icBlocks] = 12
    [icbug] = 13
    [icray] = 14
    [icbomb] = 15
    [ickeys] = 16
    [icuser] = 17
    [icsupport1] = 18
    [icsupport2] = 19
    [icStar] = 20
End Enum

Public Function AxInputBox(iboxStyle As MsgStyle, strQuestion As String, strTitle As String, ibIcon As sIcon, iButtons As StyleButtons, Optional strButton1 As String, Optional strButton2 As String) As String
Dim Frm As Form

' Estilos BASIC
Select Case iboxStyle
  Case Style3D
      Set Frm = frmIB3D
      GoTo Continuar

  Case FlatStyle1
      Set Frm = frmFlatIB
      GoTo Continuar
    
  Case FlatStyle2
      Set Frm = frmFlatIB2
      GoTo Continuar

  Case FlatStyle3
      Set Frm = frmFlatIB3
      GoTo Continuar
  
  Case FlatStyle4
      Set Frm = frmFlatIB4
      GoTo Continuar

  Case Classic
      Set Frm = frmIBClassic
      GoTo Continuar

End Select

' Estilos NS
With frmNewIBox
  Select Case iboxStyle
    Case NSError
      Set Frm = frmNewIBox
       .picV.Picture = frmIcons.PicError.Picture
       .picX.BackColor = &HE2E2FC
       .lblTitle.ForeColor = &H293FCF
      
    Case NSOK
      Set Frm = frmNewIBox
       .picV.Picture = frmIcons.PicOK.Picture
       .picX.BackColor = &HCCF0D9
       .lblTitle.ForeColor = &H8000&
      
    Case NSInfo
      Set Frm = frmNewIBox
       .picV.Picture = frmIcons.PicInfo.Picture
       .picX.BackColor = &HEDE1BC
       .lblTitle.ForeColor = &HFF8080
      
    Case NSQuestion
      Set Frm = frmNewIBox
       .picV.Picture = frmIcons.PicQuestion.Picture
       .picX.BackColor = &H9DF2F2
       .lblTitle.ForeColor = &H479AED
      
    Case NSAlert
      Set Frm = frmNewIBox
       .picV.Picture = frmIcons.PicAlert.Picture
       .picX.BackColor = &HC0E0FF
       .lblTitle.ForeColor = &H80FF&
      
  End Select

  .lblTitle.Caption = strTitle
End With

' Continuar desde Estilos Basic
Continuar:
Select Case iButtons
  Case bYesNo
    Frm.axButton1.Caption = "Si"
    Frm.axButton2.Caption = "No"
  Case bAcceptCancel
    Frm.axButton1.Caption = "Aceptar"
    Frm.axButton2.Caption = "Cancelar"
  Case bOKOnly
    Frm.axButton1.Caption = "OK"
    Frm.axButton2.Visible = False
  Case bAcceptOnly
    Frm.axButton1.Caption = "Aceptar"
    Frm.axButton2.Visible = False
End Select

Select Case ibIcon
    Case icNoIcon
      Frm.picIcon.Picture = frmIcons.Picture
    Case icok
      Frm.picIcon.Picture = frmIcons.pOK.Picture
    Case icCancel
      Frm.picIcon.Picture = frmIcons.pCancel.Picture
    Case icInfo
      Frm.picIcon.Picture = frmIcons.pInfo.Picture
    Case icAlert
      Frm.picIcon.Picture = frmIcons.pAlert.Picture
    Case icStop1
      Frm.picIcon.Picture = frmIcons.pStop1.Picture
    Case icStop2
      Frm.picIcon.Picture = frmIcons.pStop2.Picture
    Case icTips
      Frm.picIcon.Picture = frmIcons.pTips.Picture
    Case icLock
      Frm.picIcon.Picture = frmIcons.pLock.Picture
    Case icNote
      Frm.picIcon.Picture = frmIcons.pNote.Picture
    Case icSmile
      Frm.picIcon.Picture = frmIcons.pSmile.Picture
    Case icBlocks
      Frm.picIcon.Picture = frmIcons.pBlocks.Picture
    Case icbug
      Frm.picIcon.Picture = frmIcons.pBug.Picture
    Case icray
      Frm.picIcon.Picture = frmIcons.pRay.Picture
    Case icbomb
      Frm.picIcon.Picture = frmIcons.pBomb.Picture
    Case ickeys
      Frm.picIcon.Picture = frmIcons.pKeys.Picture
    Case icuser
      Frm.picIcon.Picture = frmIcons.pUser.Picture
    Case icsupport1
      Frm.picIcon.Picture = frmIcons.pSupport1.Picture
    Case icsupport2
      Frm.picIcon.Picture = frmIcons.pSupport2.Picture
    Case icStar
      Frm.picIcon.Picture = frmIcons.pStar.Picture
    Case icHelp
      Frm.picIcon.Picture = frmIcons.pHelp.Picture
End Select
    ' Seteo Mensaje ...
    If Frm.Name = frmIB3D.Name Then
      frmIB3D.lblTitle(0).Caption = strTitle
      frmIB3D.lblTitle(1).Caption = strTitle
      frmIB3D.lblMessage(0).Caption = strQuestion
      frmIB3D.lblMessage(1).Caption = strQuestion
    Else
      Frm.lblTitle.Caption = strTitle
      Frm.lblMessage.Caption = strQuestion
    End If
    ' Seteo Caption Botones
    If strButton1 <> vbNullString Then
      Frm.axButton1.Caption = strButton1
      Frm.axButton1.Visible = True
    End If
    If strButton2 <> vbNullString Then
      Frm.axButton2.Caption = strButton2
    End If
    ' Muestra MessageBox
    Frm.Show vbModal
    
    ' Retorna Botón Presionado
    If asClicked = 1 Then
      AxInputBox = strResp
    ElseIf asClicked = 2 Then
      AxInputBox = vbNullString
    End If
End Function

Public Function AxMsgBox(boxStyle As MsgStyle, strMessage As String, strTitle As String, msgIcon As sIcon, iButtons As StyleButtons, Optional strButton1 As String, Optional strButton2 As String) As ClickedButton
Dim Frm As Form
Dim lColor As Long

' Estilos BASIC
Select Case boxStyle
  Case Style3D
    With frm3D
      Set Frm = frm3D
      GoTo Continuar
    End With

  Case FlatStyle1
    With frmFlat
      Set Frm = frmFlat
      GoTo Continuar
    End With
    
  Case FlatStyle2
    With frmFlat2
      Set Frm = frmFlat2
      GoTo Continuar
    End With

  Case FlatStyle3
    With frmFlat3
      Set Frm = frmFlat3
      GoTo Continuar
    End With

  Case FlatStyle4
    With frmFlat4
      Set Frm = frmFlat4
      GoTo Continuar
    End With

  Case Classic
    With frmClassic
      Set Frm = frmClassic
      GoTo Continuar
    End With

End Select

' Estilos NS
With frmNS
  Select Case boxStyle
    Case NSError
      Set Frm = frmNS
      .picV.Picture = frmIcons.PicError.Picture
      .picX.BackColor = &HE8E8FF
      .lblTitle.ForeColor = &H293FCF
      lColor = &HE8E8FF
      
    Case NSOK
      Set Frm = frmNS
      .picV.Picture = frmIcons.PicOK.Picture
      .picX.BackColor = &HBFF2DF
      .lblTitle.ForeColor = &H8000&
      lColor = &HBFF2DF
      
    Case NSInfo
      Set Frm = frmNS
      .picV.Picture = frmIcons.PicInfo.Picture
      .picX.BackColor = &HF8E5BD
      .lblTitle.ForeColor = &HFF8080
      lColor = &HF8E5BD
      
    Case NSQuestion
      Set Frm = frmNS
      .picV.Picture = frmIcons.PicQuestion.Picture
      .picX.BackColor = &HB3EFFE
      .lblTitle.ForeColor = &H479AED
      lColor = &HB3EFFE
      
    Case NSAlert
      Set Frm = frmNS
      .picV.Picture = frmIcons.PicAlert.Picture
      .picX.BackColor = &HB3DEFF
      .lblTitle.ForeColor = &H80FF&
      lColor = &HB3DEFF
      
  End Select

  .lblTitle.Caption = strTitle
End With
     
' Continuar desde Estilos Basic
Continuar:
Select Case iButtons
  Case bYesNo
    Frm.axButton1.Caption = "Si"
    Frm.axButton2.Caption = "No"
  Case bAcceptCancel
    Frm.axButton1.Caption = "Aceptar"
    Frm.axButton2.Caption = "Cancelar"
  Case bOKOnly
    Frm.axButton1.Caption = "OK"
    Frm.axButton2.Visible = False
  Case bAcceptOnly
    Frm.axButton1.Caption = "Aceptar"
    Frm.axButton2.Visible = False
End Select

Select Case msgIcon
    Case icNoIcon
      Frm.picIcon.Picture = frmIcons.Picture
    Case icok
      Frm.picIcon.Picture = frmIcons.pOK.Picture
    Case icCancel
    Frm.picIcon.Picture = frmIcons.pCancel.Picture
    Case icInfo
    Frm.picIcon.Picture = frmIcons.pInfo.Picture
    Case icAlert
    Frm.picIcon.Picture = frmIcons.pAlert.Picture
    Case icStop1
    Frm.picIcon.Picture = frmIcons.pStop1.Picture
    Case icStop2
    Frm.picIcon.Picture = frmIcons.pStop2.Picture
    Case icTips
    Frm.picIcon.Picture = frmIcons.pTips.Picture
    Case icLock
    Frm.picIcon.Picture = frmIcons.pLock.Picture
    Case icNote
    Frm.picIcon.Picture = frmIcons.pNote.Picture
    Case icSmile
    Frm.picIcon.Picture = frmIcons.pSmile.Picture
    Case icBlocks
    Frm.picIcon.Picture = frmIcons.pBlocks.Picture
    Case icbug
    Frm.picIcon.Picture = frmIcons.pBug.Picture
    Case icray
    Frm.picIcon.Picture = frmIcons.pRay.Picture
    Case icbomb
    Frm.picIcon.Picture = frmIcons.pBomb.Picture
    Case ickeys
    Frm.picIcon.Picture = frmIcons.pKeys.Picture
    Case icuser
    Frm.picIcon.Picture = frmIcons.pUser.Picture
    Case icsupport1
    Frm.picIcon.Picture = frmIcons.pSupport1.Picture
    Case icsupport2
    Frm.picIcon.Picture = frmIcons.pSupport2.Picture
    Case icStar
    Frm.picIcon.Picture = frmIcons.pStar.Picture
    Case icHelp
    Frm.picIcon.Picture = frmIcons.pHelp.Picture
End Select
    ' Seteo Mensaje ...
    If Frm.Name = frm3D.Name Then
      frm3D.lblTitle(0).Caption = strTitle
      frm3D.lblTitle(1).Caption = strTitle
      frm3D.lblMessage(0).Caption = strMessage
      frm3D.lblMessage(1).Caption = strMessage
    Else
      Frm.lblTitle.Caption = strTitle
      Frm.lblMessage.Caption = strMessage
    End If
    ' Seteo Caption Botones
    If strButton1 <> vbNullString Then
      Frm.axButton1.Caption = strButton1
      Frm.axButton1.Visible = True
    End If
    If strButton2 <> vbNullString Then
      Frm.axButton2.Caption = strButton2
    End If
    
    ' Muestra MessageBox
    Frm.Show vbModal
    
    ' Retorna Botón Presionado
    If asClicked = 1 Then
      AxMsgBox = cButton1
    ElseIf asClicked = 2 Then
      AxMsgBox = cButton2
    End If

End Function


