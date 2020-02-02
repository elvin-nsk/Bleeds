VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_iniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=======================================================================================
' Класс:             cls_iniFile
' Версия:            2020.02.02
' Автор:             elvin-nsk (me@elvin.nsk.ru)
' Использован код:   Scott Lyerly (scott_lyerly@tjx.com, or scott.c.lyerly@gmail.com)
'                    Daniel Pineault (http://www.cardaconsultants.com)
' Назначение:        чтение/запись ini-файлов
' Использование:     1) Пакетный способ
'                    задаём переменные File, Action, Section
'                    используем универсальную функцию
'                    чтения в переменную / записи из переменной в файл
'                    / сброса в значение по умолчанию:
'                    KeyVar(key$, Value As Variant, Optional ByVal Default As Variant)
'
' Зависимости:       самодостаточный
'=======================================================================================

Option Explicit

'=======================================================================================
' объявления для работы api-функций от Scott Lyerly
'=======================================================================================

'Declaration for Reading and Wrting to an INI file.

'API Functions for Reading and Writing to INI File

#If VBA7 Then
'Declare for reading INI files (64 bit)
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" _
  Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                    ByVal lpKeyName As Any, _
                                    ByVal lpDefault As String, _
                                    ByVal lpReturnedString As String, _
                                    ByVal nSize As Long, _
                                    ByVal lpFileName As String) As Long
                                      
'Declare for writing INI files (64 bit)
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" _
  Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                      ByVal lpKeyName As Any, _
                                      ByVal lpString As Any, _
                                      ByVal lpFileName As String) As Long
#Else
'Declare for reading INI files (32 bit)
Private Declare Function GetPrivateProfileString Lib "kernel32" _
  Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                    ByVal lpKeyName As Any, _
                                    ByVal lpDefault As String, _
                                    ByVal lpReturnedString As String, _
                                    ByVal nSize As Long, _
                                    ByVal lpFileName As String) As Long
                                      
'Declare for writing INI files (32 bit)
Private Declare Function WritePrivateProfileString Lib "kernel32" _
  Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                      ByVal lpKeyName As Any, _
                                      ByVal lpString As Any, _
                                      ByVal lpFileName As String) As Long
#End If

'=======================================================================================
' публичные нумераторы
'=======================================================================================

Public Enum iniAction
  iniRead = 1
  iniWrite = 2
  iniReset = 3
End Enum

Public Enum iniHandler
  useAPI = 1
  useVBA = 2
End Enum

Public Enum iniError
  OK = 0
  Unknown = 1
  NoFile = 2
  NoSection = 3
  NoKey = 4
  InvalidValue = 5
End Enum


'=======================================================================================
' приватные переменные
'=======================================================================================

Private myHandler As iniHandler
Private myFile$
Private myAction As iniAction
Private mySection$
Private myError As iniError


'=======================================================================================
' публичная часть
'=======================================================================================

Property Let File(f$)
  myFile = f
End Property
Property Get File() As String
  File = myFile
End Property

Property Let Action(a As iniAction)
  myAction = a
End Property
Property Get Action() As iniAction
  Action = myAction
End Property

Property Let Section(s$)
  mySection = s
End Property
Property Get Section() As String
  Section = mySection
End Property

Property Get Error() As iniError
  Err = myError
End Property

Public Function KeyVar(ByVal Key$, ByRef Value As Variant, Optional ByVal DefaultValue As Variant) As Boolean
  
  Dim v$
  
  KeyVar = False
  
  Select Case myAction
    Case iniRead
      KeyVar = Entry(iniRead, mySection, Key, v)
      If KeyVar = False And VarType(Value) = VarType(DefaultValue) Then
        Value = DefaultValue
      ElseIf KeyVar = True And v <> "" Then
        Select Case VarType(Value)
          Case vbInteger
            Value = CInt(v)
          Case vbLong
            Value = CLng(v)
          Case vbSingle
            Value = CSng(v)
          Case vbDouble
            Value = CDbl(v)
          Case vbCurrency
            Value = CCur(v)
          Case vbDate
            Value = CDate(v)
          Case vbBoolean
            Value = CBool(v)
          Case vbByte
            Value = CByte(v)
          Case vbString
            Value = v
          Case vbVariant
            Value = CVar(v)
          Case Else
            KeyVar = False
        End Select
      End If
    Case iniWrite
      If VarType(Value) = vbInteger Or _
                          vbLong Or _
                          vbSingle Or _
                          vbDouble Or _
                          vbCurrency Or _
                          vbDate Or _
                          vbBoolean Or _
                          vbByte Or _
                          vbString Or _
                          vbVariant Then
        v = CStr(Value)
        KeyVar = Entry(iniWrite, mySection, Key, v)
      End If
    Case iniReset
      Value = DefaultValue
      KeyVar = True
  End Select
  
End Function

Public Function Entry(ByVal Action As iniAction, _
                      ByVal Section$, _
                      ByVal Key$, _
                      ByRef Value$, _
                      Optional ByVal Default$ = "" _
                      ) As Boolean
    
  Dim v$
  
  Entry = False
  Select Case Action
    Case iniRead
      v = sManageSectionEntry(Action, Section, Key, myFile)
      If v <> "" Then
        Value = v
        Entry = True
      Else
        If Default <> "" Then Value = Default
      End If
    Case iniWrite
      If sManageSectionEntry(Action, Section, Key, myFile, Value) = Value Then
        Entry = True
      End If
    Case iniReset
      Value = Default
      Entry = True
  End Select

End Function


'=======================================================================================
' приватная часть
'=======================================================================================

Private Sub Class_Initialize()
  myHandler = useAPI
  myFile = ""
  myAction = iniReset
  mySection = "DEFAULT"
End Sub

'---------------------------------------------------------------------------------------
' автор - Scott Lyerly
' секция работы через Windows API
'---------------------------------------------------------------------------------------

'*******************************************************************************
' Description:  This reads an INI file section/key combination and
'               returns the read value as a string.
'
' Author:       Scott Lyerly
' Contact:      scott_lyerly@tjx.com, or scott.c.lyerly@gmail.com
'
' Notes:        Requires "Private Declare Function GetPrivateProfileString" and
'               "WritePrivateProfileString" to be added in the declarations
'               at the top of the module.
'
' Name:                 Date:           Init:   Modification:
' sManageSectionEntry   26-Nov-2013     SCL     Original development
'
' Arguments:    inAction    The action to take in teh funciton, reading or writing to
'                           to the INI file. Uses the enumeration iniAction in the
'                           declarations section.
'               sSection    The seciton of the INI file to search
'               sKey        The key of the INI from which to retrieve a value
'               sIniFile    The name and directory location of the INI file
'               sValue      The value to be written to the INI file (if writing - optional)
'
' Returns:      string      The return string is one of three things:
'                           1) The value being sought from the INI file.
'                           2) The value being written to the INI file (should match
'                              the sValue parameter).
'                           3) The word "Error". This can be changed to whatever makes
'                              the most sense to the programmer using it.
'*******************************************************************************
Private Function sManageSectionEntry(inAction As iniAction, _
                                     sSection As String, _
                                     sKey As String, _
                                     sIniFile As String, _
                                     Optional sValue As String) As String

  On Error GoTo Err_ManageSectionEntry

  ' Variable declarations.
  Dim sRetBuf         As String
  Dim iLenBuf         As Integer
  Dim sFileName       As String
  Dim sReturnValue    As String
  Dim lRetVal         As Long
    
  ' Based on the inAction parameter, take action.
  If inAction = iniRead Then  ' If reading from the INI file.

    ' Set the return buffer to by 256 spaces. This should be enough to
    ' hold the value being returned from the INI file, but if not,
    ' increase the value.
    sRetBuf = Space(1024)

    ' Get the size of the return buffer.
    iLenBuf = Len(sRetBuf)

    ' Read the INI Section/Key value into the return variable.
    sReturnValue = GetPrivateProfileString(sSection, _
                                           sKey, _
                                           "", _
                                           sRetBuf, _
                                           iLenBuf, _
                                           sIniFile)

    ' Trim the excess garbage that comes through with the variable.
    sReturnValue = Trim(Left(sRetBuf, sReturnValue))

    ' If we get a value returned, pass it back as the argument.
    ' Else pass "False".
    If Len(sReturnValue) > 0 Then
      sManageSectionEntry = sReturnValue
    Else
      sManageSectionEntry = ""
    End If
  ElseIf inAction = iniWrite Then ' If writing to the INI file.

    ' Check to see if a value was passed in the sValue parameter.
    If Len(sValue) = 0 Then
      sManageSectionEntry = ""

    Else
            
      ' Write to the INI file and capture the value returned
      ' in the API function.
      lRetVal = WritePrivateProfileString(sSection, _
                                          sKey, _
                                          sValue, _
                                          sIniFile)

      ' Check to see if we had an error wrting to the INI file.
      If lRetVal = 0 Then sManageSectionEntry = ""

    End If
  End If
    
Exit_Clean:
  Exit Function
    
Err_ManageSectionEntry:
  MsgBox Err.Number & ": " & Err.Description
  Resume Exit_Clean

End Function
