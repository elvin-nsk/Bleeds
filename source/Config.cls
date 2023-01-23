VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Config v2022.12.22
' хранение, загрузка и сохранение настроек
'===============================================================================

'@PredeclaredId
Option Explicit

#Const ReadOnly = False
#Const DefaultContent = False

'===============================================================================

Private Const iniFileName As String = "elvin_" & APP_NAME & ".ini"

'===============================================================================

Public Bleeds As Double
Public RoundSize As Boolean
Public RoundDec As Long
Public BitmapTrim As Boolean
Public BitmapTrimSize As Long
Public BitmapFlatten As Boolean

'-------------------------------------------------------------------------------

Private Sub Assign(ByVal Action As iniAction)

    With iniFile.Create
    
        iniFilePrepare .Self, Action
        
        '.Section = "DEFAULT"
        .KeyVar "Bleeds", Bleeds, 3
        .KeyVar "RoundSize", RoundSize, True
        .KeyVar "RoundDec", RoundDec, 0
        .KeyVar "BitmapTrim", BitmapTrim, True
        .KeyVar "BitmapTrimSize", BitmapTrimSize, 2
        .KeyVar "BitmapFlatten", BitmapFlatten, False
    
    End With
    
End Sub

'-------------------------------------------------------------------------------

'Опционально. Создаёт ini-файл заданного содержания (напр. с комментами).
#If DefaultContent Then
Private Sub MakeDefaultFile(ByVal File As String)

    Dim Content As String
    
    Content = Content & "; кодировка Windows-1251" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "[DEFAULT]" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; булка" & vbCrLf
    Content = Content & "Bulka=Рулька" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; палка" & vbCrLf
    Content = Content & "Palka=1" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; крышка" & vbCrLf
    Content = Content & "Kryshka=2" & vbCrLf
    
    WriteFile Content, File
    
End Sub
#End If

'===============================================================================
'           * * * Контент, не предназначенный для изменения * * *
'===============================================================================
' # Constructors

Public Function Bind() As Config
    Set Bind = New Config
    Bind.ForceLoad
End Function

#If Not ReadOnly Then
Private Sub Class_Terminate()
    ForceSave
End Sub
#End If

'===============================================================================
' # Public

Public Sub ForceLoad()
    Assign iniRead
End Sub

#If Not ReadOnly Then
Public Sub ForceSave()
    Assign iniWrite
End Sub
#End If

Public Property Get Self() As Config
    Set Self = Me
End Property

'===============================================================================

Private Sub iniFilePrepare(ByVal iniFile As iniFile, ByVal Action As iniAction)

    Dim iniFilePath As String

    With iniFile
        Select Case True
        Case FileExists(GMSManager.UserGMSPath & iniFileName)
            .File = Application.GMSManager.UserGMSPath & iniFileName
            .Action = Action
        Case FileExists(GMSManager.GMSPath & iniFileName)
            .File = Application.GMSManager.GMSPath & iniFileName
            .Action = Action
        Case Else
            .File = GMSManager.UserGMSPath & iniFileName
            #If DefaultContent Then
            MakeDefaultFile .File
            If Action = iniRead Then
                .Action = iniRead
            Else
                .Action = iniWrite
            End If
            #Else
            If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
            #End If
        End Select
    End With

End Sub

Private Function FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Function
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Private Sub WriteFile( _
               ByVal Content As String, _
               ByVal File As String _
           )
    Dim FileNum As Long
    FileNum = VBA.FreeFile
    Open File For Output Access Write As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub
