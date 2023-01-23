VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   Caption         =   "Добавить припуски"
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

Public BleedsMin As Double
Public BleedsMax As Double

Private Main As MainLogic
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

'===============================================================================

Private Sub UserForm_Activate()
    Set Main = MainLogic.Create(Me)
    Set App = Application
    ExecutabilityCheck
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

Private Sub btnExecute_Click()
    FormExecute
End Sub

Private Sub App_SelectionChange()
    ExecutabilityCheck
    VisibilityCheck
End Sub

'===============================================================================

Private Sub ExecutabilityCheck()
    If ActiveDocument Is Nothing Then
        btnExecute.Enabled = False
        btnExecute.Caption = "Нет документа"
        Exit Sub
    End If
    If ActiveSelectionRange.Count > 1 Then
        btnExecute.Enabled = False
        btnExecute.Caption = "Несколько объектов"
        Exit Sub
    End If
    If ActiveSelectionRange.Count < 1 Then
        btnExecute.Enabled = False
        btnExecute.Caption = "Не выбран объект"
        Exit Sub
    End If
    btnExecute.Enabled = True
    btnExecute.Caption = "Выполнить"
    btnExecute.SetFocus
End Sub

Private Sub VisibilityCheck()
    If ActiveDocument Is Nothing Then
        RastrHide
        Exit Sub
    End If
    If IsShapeType(ActiveSelectionRange.FirstShape, cdrBitmapShape) Then
        RastrShow
    Else
        RastrHide
    End If
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

Private Sub FormExecute()
    Main.Execute Me
End Sub

Private Sub FormCancel()
    Main.Dispose Me
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
