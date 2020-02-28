VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Main 
   Caption         =   "Добавить припуски"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3615
   OleObjectBlob   =   "frm_Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=======================================================================================
' события
'=======================================================================================

Private Sub UserForm_Initialize()
  cfg.Load
  tbBleeds.Value = CStr(cfg.Bleeds)
  cbRound.Value = cfg.RoundSize
  Select Case cfg.RoundDec
    Case 0
      obRound0.Value = "True"
    Case 1
      obRound1.Value = "True"
    Case 2
      obRound2.Value = "True"
  End Select
  cbTrim.Value = cfg.BitmapTrim
  tbTrim.Value = cfg.BitmapTrimSize
  cbFlatten.Value = cfg.BitmapFlatten
End Sub

Private Sub UserForm_Terminate()
  cfg.Bleeds = tbBleeds.Value
  cfg.RoundSize = cbRound.Value
  Select Case True
    Case obRound0.Value = True
      cfg.RoundDec = 0
    Case obRound1.Value = True
      cfg.RoundDec = 1
    Case obRound2.Value = True
      cfg.RoundDec = 2
  End Select
  cfg.BitmapTrim = cbTrim.Value
  cfg.BitmapTrimSize = tbTrim.Value
  cfg.BitmapFlatten = cbFlatten.Value
  cfg.Save
End Sub

Private Sub tbBleeds_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  onlyNum KeyAscii
End Sub
Private Sub tbBleeds_AfterUpdate()
  checkRange tbBleeds, 0.5, 10000
End Sub

Private Sub tbTrim_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  onlyInt KeyAscii
End Sub
Private Sub tbTrim_AfterUpdate()
  checkRange tbTrim, 1, 10
End Sub

Private Sub btnOK_Click()
  FormОК
End Sub

Private Sub btnCancel_Click()
  FormCancel
End Sub

'=======================================================================================
' приватные функции
'=======================================================================================

Private Sub FormОК()
  Me.Hide
  DoBleeds
End Sub

Private Sub FormCancel()
  Me.Hide
End Sub

'=======================================================================================
' библиотека вспомогательных функций
'=======================================================================================

Private Sub onlyInt(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub onlyNum(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub checkRange(TextBox As MSForms.TextBox, ByVal Min#, ByVal Max#)
  If TextBox.Value > Max Then TextBox.Value = Max
  If TextBox.Value < Min Then TextBox.Value = Min
End Sub
