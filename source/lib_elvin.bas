Attribute VB_Name = "lib_elvin"
'=======================================================================================
' Модуль:            lib_elvin
' Версия:            2020.02.01
' Автор:             elvin-nsk (me@elvin.nsk.ru)
' Использован код:   dizzy (из макроса CtC)
'                    и др.
' Описание:          библиотека функций для макросов от elvin-nsk
' Использование:
' Зависимости:       самодостаточный
'=======================================================================================

Option Explicit

'=======================================================================================
' # приватные переменные модуля
'=======================================================================================

Private Type type_LayerProps
  Visible As Boolean
  Printable As Boolean
  Editable As Boolean
End Type

'=======================================================================================
' функции общего назначения
'=======================================================================================

'---------------------------------------------------------------------------------------
' Функции          : BoostStart, BoostFinish
' Версия           : 2019.09.06
' Авторы           : dizzy, elvin-nsk
' Назначение       : доработанные оптимизаторы от CtC
' Зависимости      : самодостаточные
'
' Параметры:
' ~~~~~~~~~~
'
'
' Использование:
' ~~~~~~~~~~~~~~
'
'---------------------------------------------------------------------------------------
Sub BoostStart(Optional ByVal UnDo$ = "", Optional ByVal Optimize = True)
  If UnDo <> "" And Not (ActiveDocument Is Nothing) Then ActiveDocument.BeginCommandGroup UnDo
  If Optimize Then Optimization = True
  EventsEnabled = False
  If Not ActiveDocument Is Nothing Then
    With ActiveDocument
      .SaveSettings
      '.PreserveSelection = False отключено, вызывает глюки с intersect, на производительность при включенной оптимизации почти не влияет
      .Unit = cdrMillimeter
      .ReferencePoint = cdrCenter
    End With
  End If
End Sub
Sub BoostFinish(Optional ByVal EndUndoGroup = True)
  EventsEnabled = True
  Optimization = False
  If Not ActiveDocument Is Nothing Then
    With ActiveDocument
      .RestoreSettings
      If EndUndoGroup Then .EndCommandGroup
    End With
    ActiveWindow.Refresh
  End If
  Application.Refresh
End Sub

'=======================================================================================
' функции для работы с объектами корела
'=======================================================================================

'все объекты на всех страницах, включая мастер-страницу - на один слой
'все страницы прибиваются, все объекты на слоях guides прибиваются
Function FlattenPagesToLayer(ByVal LayerName$) As Layer

  Dim DL As Layer: Set DL = ActiveDocument.MasterPage.DesktopLayer
  Dim DLstate As Boolean: DLstate = DL.Editable
  Dim p As Page
  Dim L As Layer
  
  DL.Editable = False
  
  For Each p In ActiveDocument.Pages
    For Each L In p.Layers
      If L.IsSpecialLayer Then
        L.Shapes.All.Delete
      Else
        L.Activate
        L.Editable = True
        With L.Shapes.All
          .MoveToLayer DL
          .OrderToBack
        End With
        L.Delete
      End If
    Next
    If p.Index <> 1 Then p.Delete
  Next
  
  Set FlattenPagesToLayer = ActiveDocument.Pages.First.CreateLayer(LayerName)
  FlattenPagesToLayer.MoveBelow ActiveDocument.Pages.First.GuidesLayer
  
  For Each L In ActiveDocument.MasterPage.Layers
    If Not L.IsSpecialLayer Or L.IsDesktopLayer Then
      L.Activate
      L.Editable = True
      With L.Shapes.All
        .MoveToLayer FlattenPagesToLayer
        .OrderToBack
      End With
      If Not L.IsSpecialLayer Then L.Delete
    Else
      L.Shapes.All.Delete
    End If
  Next
  
  FlattenPagesToLayer.Activate
  DL.Editable = DLstate

End Function

'дублировать активную страницу со всеми слоями и объектами
Function DuplicateActivePage(ByVal NumberOfPages&, Optional ExcludeLayerName$ = "") As Page
  Dim Range As ShapeRange
  Dim tShape As Shape, sDuplicate As Shape
  Dim tProps As type_LayerProps
  Dim i&
  For i = 1 To NumberOfPages
    Set Range = FindShapesActivePageLayers
    Set DuplicateActivePage = ActiveDocument.InsertPages(1, False, ActivePage.Index)
    DuplicateActivePage.SizeHeight = ActivePage.SizeHeight
    DuplicateActivePage.SizeWidth = ActivePage.SizeWidth
    For Each tShape In Range.ReverseRange
      If tShape.Layer.Name <> ExcludeLayerName Then
        layerPropsPreserve tShape.Layer, tProps
        layerPropsReset tShape.Layer
        Set sDuplicate = tShape.Duplicate
        sDuplicate.MoveToLayer FindLayerDuplicate(DuplicateActivePage, tShape.Layer)
        layerPropsRestore tShape.Layer, tProps
      End If
    Next tShape
  Next i
End Function

'ПРОВЕРИТЬ КАК СЛЕДУЕТ
Function IsOverlap(FirstShape As Shape, SecondShape As Shape) As Boolean
  Dim tShape As Shape
  Dim tProps As type_LayerProps
  'запоминаем какой слой был активным
  Dim tLayer As Layer: Set tLayer = ActiveLayer
  'запоминаем состояние первого слоя
  FirstShape.Layer.Activate
  layerPropsPreserve FirstShape.Layer, tProps
  layerPropsReset FirstShape.Layer
  Set tShape = FirstShape.Intersect(SecondShape)
  If tShape Is Nothing Then
    IsOverlap = False
  Else
    tShape.Delete
    IsOverlap = True
  End If
  'возвращаем всё на место
  layerPropsRestore FirstShape.Layer, tProps
  tLayer.Activate
End Function

'перекрашивает объект в чёрный или белый в серой шкале,
'в зависимости от исходного цвета
'ДОРАБОТАТЬ
Function ContrastShape(SrcShape As Shape) As Shape
  With SrcShape.Fill
    Select Case .Type
      Case cdrUniformFill
        .UniformColor.ConvertToGray
        If .UniformColor.Gray < 128 Then .UniformColor.GrayAssign 0 Else .UniformColor.GrayAssign 255
      Case cdrFountainFill
        'todo
    End Select
  End With
  With SrcShape.Outline
    If .Type <> cdrNoOutline Then
      .Color.ConvertToGray
      If .Color.Gray < 128 Then .Color.GrayAssign 0 Else .Color.GrayAssign 255
    End If
  End With
  Set ContrastShape = SrcShape
End Function

'обрезать битмап по CropEnvelopeShape, но по-умному, сначала кропнув на EXPANDBY пикселей побольше
Function TrimBitmap(BitmapShape As Shape, CropEnvelopeShape As Shape, Optional ByVal LeaveCropEnvelope As Boolean = True) As Shape

  Const EXPANDBY& = 2 'px
  
  Dim Crop As Shape
  Dim PxW#, PxH#
  Dim SaveUnit As cdrUnit

  If BitmapShape.Type <> cdrBitmapShape Then Exit Function
  
  'save
  SaveUnit = ActiveDocument.Unit
  
  ActiveDocument.Unit = cdrInch
  PxW = 1 / BitmapShape.Bitmap.ResolutionX
  PxH = 1 / BitmapShape.Bitmap.ResolutionY
  BitmapShape.Bitmap.ResetCropEnvelope
  Set Crop = BitmapShape.Layer.CreateRectangle(CropEnvelopeShape.LeftX - PxW * EXPANDBY, _
                                                CropEnvelopeShape.TopY + PxH * EXPANDBY, _
                                                CropEnvelopeShape.RightX + PxW * EXPANDBY, _
                                                CropEnvelopeShape.BottomY - PxH * EXPANDBY)
  Set TrimBitmap = Crop.Intersect(BitmapShape, False, False)
  If TrimBitmap Is Nothing Then
    Crop.Delete
    GoTo ExitFunction
  End If
  TrimBitmap.Bitmap.Crop
  Set TrimBitmap = CropEnvelopeShape.Intersect(TrimBitmap, LeaveCropEnvelope, False)
  
ExitFunction:
  'restore
  ActiveDocument.Unit = SaveUnit
  
End Function

'отрезать кусок от SourceShape по контуру Knife, возвращает отрезанный кусок
Function Dissect(ByRef SourceShape As Shape, ByRef Knife As Shape) As Shape
  Set Dissect = Knife.Intersect(SourceShape, True, True)
  Set SourceShape = Knife.Trim(SourceShape, True, False)
End Function

'инструмент Crop Tool
Function CropTool(ShapeOrRangeOrPage As Object, ByVal x1#, ByVal y1#, ByVal x2#, ByVal y2#, Optional ByVal Angle = 0) As ShapeRange
  If TypeOf ShapeOrRangeOrPage Is Shape Or _
     TypeOf ShapeOrRangeOrPage Is ShapeRange Or _
     TypeOf ShapeOrRangeOrPage Is Page Then
    Set CropTool = ShapeOrRangeOrPage.CustomCommand("Crop", "CropRectArea", x1, y1, x2, y2, Angle)
  End If
End Function

'инструмент Boundary
Function CreateBoundary(ShapeOrRange As Object) As Shape
  Dim tShape As Shape, Range As ShapeRange
  'просто объект не ест, надо конкретный тип
  If TypeOf ShapeOrRange Is Shape Then
    Set tShape = ShapeOrRange
    Set CreateBoundary = tShape.CustomCommand("Boundary", "CreateBoundary")
  ElseIf TypeOf ShapeOrRange Is ShapeRange Then
    Set Range = ShapeOrRange
    Set CreateBoundary = Range.CustomCommand("Boundary", "CreateBoundary")
  End If
End Function

'тестирует на nothing и на мусорный шейп в коллекции
Function IsNothing(TestShape As Shape) As Boolean
  Dim n$
  On Error GoTo ExitTrue
  If TestShape Is Nothing Then GoTo ExitTrue
  n = TestShape.Name
ExitFalse:
  IsNothing = False
  Exit Function
ExitTrue:
  IsNothing = True
End Function

'возвращает бОльшую сторону шейпа/рэйнджа/страницы
Function GreaterDim(ShapeOrRangeOrPage As Object) As Double
  If (Not TypeOf ShapeOrRangeOrPage Is Shape) And (Not TypeOf ShapeOrRangeOrPage Is ShapeRange) And (Not TypeOf ShapeOrRangeOrPage Is Page) Then Exit Function
  If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then GreaterDim = ShapeOrRangeOrPage.SizeWidth Else GreaterDim = ShapeOrRangeOrPage.SizeHeight
End Function

'возвращает среднее сторон шейпа/рэйнджа/страницы
Function AverageDim(ShapeOrRangeOrPage As Object) As Double
  If (Not TypeOf ShapeOrRangeOrPage Is Shape) And (Not TypeOf ShapeOrRangeOrPage Is ShapeRange) And (Not TypeOf ShapeOrRangeOrPage Is Page) Then Exit Function
  AverageDim = (ShapeOrRangeOrPage.SizeWidth + ShapeOrRangeOrPage.SizeHeight) / 2
End Function

'наращивает Rect со всех сторон на Space
Function SpaceBox(ShapeOrRange As Object, Space#) As Rect
  If (Not TypeOf ShapeOrRange Is Shape) And (Not TypeOf ShapeOrRange Is ShapeRange) Then Exit Function
  'Dim s As Shape 'debug
  Set SpaceBox = ShapeOrRange.BoundingBox
  SpaceBox.Inflate Space, Space, Space, Space
End Function

'является ли шейп/рэйндж/страница альбомным
Function IsLandscape(ShapeOrRangeOrPage As Object) As Boolean
  If (Not TypeOf ShapeOrRangeOrPage Is Shape) And (Not TypeOf ShapeOrRangeOrPage Is ShapeRange) And (Not TypeOf ShapeOrRangeOrPage Is Page) Then Exit Function
  If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then IsLandscape = True Else IsLandscape = False
End Function

'являются ли кривые дубликатами, находящимися друг над другом в одном месте (underlying dubs)
Function IsSameCurves(Curve1 As Curve, Curve2 As Curve) As Boolean
  Dim tNode As Node
  Dim tJitter#: tJitter = ConvertUnits(0.001, cdrMillimeter, ActiveDocument.Unit) 'допуск = 0.001 мм
  IsSameCurves = False
  If Curve1.Nodes.Count <> Curve2.Nodes.Count Then Exit Function
  If Abs(Curve1.Length - Curve2.Length) > tJitter Then Exit Function
  For Each tNode In Curve1.Nodes
    If Curve2.FindNodeAtPoint(tNode.PositionX, tNode.PositionY, tJitter * 2) Is Nothing Then Exit Function
  Next
  IsSameCurves = True
End Function

'=======================================================================================
' функции поиска
'=======================================================================================

Function FindShapesByName(SourceRange As ShapeRange, ByVal Name$) As ShapeRange
  Set FindShapesByName = SourceRange.Shapes.FindShapes(Name)
End Function

Function FindShapesByNamePart(SourceRange As ShapeRange, ByVal NamePart$) As ShapeRange
  Set FindShapesByNamePart = SourceRange.Shapes.FindShapes(Query:="@Name.Contains('" & NamePart & "')")
End Function

'возвращает все шейпы на всех слоях текущей страницы, по умолчанию без мастер-слоёв и без гайдов
Function FindShapesActivePageLayers(Optional GuidesLayers As Boolean = False, _
                                    Optional MasterLayers As Boolean = False _
                                    ) As ShapeRange
  Dim tLayer As Layer
  Set FindShapesActivePageLayers = New ShapeRange
  For Each tLayer In ActivePage.Layers
    If Not (tLayer.IsGuidesLayer And (GuidesLayers = False)) Then _
      FindShapesActivePageLayers.AddRange tLayer.Shapes.All
  Next
  If MasterLayers Then
    For Each tLayer In ActiveDocument.MasterPage.Layers
      If Not (tLayer.IsGuidesLayer And (GuidesLayers = False)) Then _
        FindShapesActivePageLayers.AddRange tLayer.Shapes.All
  Next
  End If
End Function

'возвращает коллекцию слоёв с текущей страницы, имена которых включают NamePart
Function FindLayersActivePageByNamePart(ByVal NamePart$, Optional ByVal SearchMasters = True) As Collection
  Dim tLayer As Layer
  Dim tLayers As Layers
  If SearchMasters Then Set tLayers = ActivePage.AllLayers Else Set tLayers = ActivePage.Layers
  Set FindLayersActivePageByNamePart = New Collection
  For Each tLayer In tLayers
    If InStr(tLayer.Name, NamePart) > 0 Then FindLayersActivePageByNamePart.Add tLayer
  Next
End Function

'найти дубликат слоя по ряду параметров (достовернее, чем поиск по имени)
Function FindLayerDuplicate(PageToSearch As Page, SrcLayer As Layer) As Layer
  For Each FindLayerDuplicate In PageToSearch.AllLayers
    With FindLayerDuplicate
      If (.Name = SrcLayer.Name) And _
         (.IsDesktopLayer = SrcLayer.IsDesktopLayer) And _
         (.Master = SrcLayer.Master) And _
         (.Color.IsSame(SrcLayer.Color)) Then _
         Exit Function
    End With
  Next
  Set FindLayerDuplicate = Nothing
End Function

'функция отсюда: https://stackoverflow.com/questions/38267950/check-if-a-value-is-in-an-array-or-not-with-excel-vba
Function IsStrInArr(ByVal stringToBeFound$, Arr As Variant) As Boolean
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) = stringToBeFound Then
            IsStrInArr = True
            Exit Function
        End If
    Next i
    IsStrInArr = False
End Function

'является ли число чётным :) Что такое Even и Odd запоминать лень...
Function IsChet(ByVal X) As Boolean
  If X Mod 2 = 0 Then IsChet = True Else IsChet = False
End Function

'Generates a guid, works on both mac and windows
'отсюда: https://github.com/Martin-Carlsson/Business-Intelligence-Goodies/blob/master/Excel/GenerateGiud/GenerateGiud.bas
Function CreateGUID() As String
  CreateGUID = randomHex(3) + "-" + _
    randomHex(2) + "-" + _
    randomHex(2) + "-" + _
    randomHex(2) + "-" + _
    randomHex(6)
End Function
'From: https://www.mrexcel.com/forum/excel-questions/301472-need-help-generate-hexadecimal-codes-randomly.html#post1479527
Private Function randomHex(lngCharLength As Long) As String
  Dim i As Long
  Randomize
  For i = 1 To lngCharLength
    randomHex = randomHex & Right$("0" & Hex(Rnd() * 256), 2)
  Next
End Function

'=======================================================================================
' функции работы с файлами
'=======================================================================================

'заменяет расширение файлу на заданное
Function SetFileExt(ByVal SourceFile$, ByVal NewExt$) As String
  If Right(SourceFile, 1) <> "\" And Len(SourceFile) > 0 Then
    SetFileExt = GetFileNameNoExt(SourceFile$) & "." & NewExt
  End If
End Function

'возвращает имя файла без расширения
Function GetFileNameNoExt(FILENAME) As String
  If Right(FILENAME, 1) <> "\" And Len(FILENAME) > 0 Then
    GetFileNameNoExt = Left(FILENAME, _
      Switch _
        (InStr(FILENAME, ".") = 0, _
          Len(FILENAME), _
        InStr(FILENAME, ".") > 0, _
          InStrRev(FILENAME, ".") - 1))
  End If
End Function

'создаёт папку, если не было
Function MakeDir(Path$) As String
  If Dir(Path, vbDirectory) = "" Then MkDir (Path)
  MakeDir = Path
End Function

'---------------------------------------------------------------------------------------
' Procedure : FileExist
' DateTime  : 2007-Mar-06 13:51
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Test for the existance of a file; Returns True/False
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - name of the file to be tested for including full path
'---------------------------------------------------------------------------------------
Function FileExist(ByVal sFile As String) As Boolean
On Error GoTo Err_Handler
 
    If Len(Dir(sFile)) > 0 Then
        FileExist = True
    End If
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: FileExist" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFileName
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Return the filename from a path\filename input
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - string of a path and filename (ie: "c:\temp\test.xls")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2008-Feb-06                 Initial Release
'---------------------------------------------------------------------------------------
Function GetFileName(sFile As String)
On Error GoTo Err_Handler
 
    GetFileName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: GetFileName" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFilePath
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Return the path from a path\filename input
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - string of a path and filename (ie: "c:\temp\test.xls")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2008-Feb-06                 Initial Release
'---------------------------------------------------------------------------------------
Function GetFilePath(sFile As String)
On Error GoTo Err_Handler
 
    GetFilePath = Left(sFile, InStrRev(sFile, "\"))
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: GetFilePath" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

'=======================================================================================
' # приватные функции модуля
'=======================================================================================

Private Sub layerPropsPreserve(L As Layer, ByRef Props As type_LayerProps)
  With Props
    .Visible = L.Visible
    .Printable = L.Printable
    .Editable = L.Editable
  End With
End Sub
Private Sub layerPropsRestore(L As Layer, ByRef Props As type_LayerProps)
  With Props
    L.Visible = .Visible
    L.Printable = .Printable
    L.Editable = .Editable
  End With
End Sub
Private Sub layerPropsReset(L As Layer)
  With L
    .Visible = True
    .Printable = True
    .Editable = True
  End With
End Sub
