VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Config
' хранение, загрузка и сохранение настроек
'===============================================================================

Option Explicit

'===============================================================================

'имя файла настроек
Private Const FILENAME$ = "elvin_Bleeds.ini"

'===============================================================================

Public Bleeds#
Public RoundSize As Boolean
Public RoundDec&
Public BitmapTrim As Boolean
Public BitmapTrimSize&
Public BitmapFlatten As Boolean

'===============================================================================

Private Sub Assign(Action As iniAction)
  
  Dim iniFile As New cls_iniFile
  
  iniFilePrepare iniFile, Action
  
  With iniFile
    
    .KeyVar "Bleeds", Bleeds, 3
    .KeyVar "RoundSize", RoundSize, True
    .KeyVar "RoundDec", RoundDec, 0
    .KeyVar "BitmapTrim", BitmapTrim, True
    .KeyVar "BitmapTrimSize", BitmapTrimSize, 2
    .KeyVar "BitmapFlatten", BitmapFlatten, False
  
  End With
  
End Sub

'===============================================================================

Public Sub Load()
  Assign iniRead
End Sub
Public Sub Save()
  Assign iniWrite
End Sub

'===============================================================================

Sub iniFilePrepare(iniFile As cls_iniFile, Action As iniAction)

  Dim iniFilePath$

  With iniFile
    Select Case True
    Case lib_elvin.FileExist(Application.GMSManager.UserGMSPath & FILENAME)
      .File = Application.GMSManager.UserGMSPath & FILENAME
      .Action = Action
    Case lib_elvin.FileExist(Application.GMSManager.GMSPath & FILENAME)
      .File = Application.GMSManager.GMSPath & FILENAME
      .Action = Action
    Case Else
      .File = Application.GMSManager.UserGMSPath & FILENAME
      If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
    End Select
  End With

End Sub
