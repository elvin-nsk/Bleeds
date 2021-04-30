VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   Caption         =   "Äîáàâèòü ïðèïóñêè"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   OleObjectBlob   =   "MainView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public IsOK As Boolean
Public IsRastr As Boolean
Public BleedsMin As Double
Public BleedsMax As Double

'===============================================================================

Private Sub UserForm_Activate()
  VisibilityCheck
End Sub

Private Sub tbBleeds_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNum KeyAscii
End Sub
Private Sub tbBleeds_AfterUpdate()
  CheckRange tbBleeds, BleedsMin, BleedsMax
End Sub

Private Sub tbTrim_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyInt KeyAscii
End Sub
Private Sub tbTrim_AfterUpdate()
  CheckRange tbTrim, 1, 10
End Sub

Private Sub btnCancel_Click()
  FormCancel
End Sub

Private Sub btnOK_Click()
  FormÎÊ
End Sub

'===============================================================================

Private Sub VisibilityCheck()
  If IsRastr Then RastrShow Else RastrHide
End Sub

Private Sub RastrShow()
  cbTrim.Enabled = True
  tbTrim.Enabled = True
  lblPix.Enabled = True
  cbFlatten.Enabled = True
End Sub

Private Sub RastrHide()
  cbTrim.Enabled = False
  tbTrim.Enabled = False
  lblPix.Enabled = False
  cbFlatten.Enabled = False
End Sub

Private Sub FormÎÊ()
  Me.Hide
  IsOK = True
End Sub

Private Sub FormCancel()
  Me.Hide
End Sub

'===============================================================================

Private Sub OnlyInt(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub OnlyNum(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub CheckRange(TextBox As MSForms.TextBox, ByVal Min As Double, Optional ByVal Max As Double = 2147483647)
  With TextBox
    If CDbl(.Value) > Max Then .Value = CStr(Max)
    If CDbl(.Value) < Min Then .Value = CStr(Min)
  End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Cancel = True
    FormCancel
  End If
End Sub
