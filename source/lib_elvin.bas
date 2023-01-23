Attribute VB_Name = "lib_elvin"
'===============================================================================
'   Модуль          : lib_elvin
'   Версия          : 2023.01.23
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Использован код : dizzy (из макроса CtC), Alex Vakulenko
'                     и др.
'   Описание        : библиотека функций для макросов от elvin-nsk
'   Использование   :
'   Зависимости     : самодостаточный
'===============================================================================

Option Explicit

'===============================================================================
' # приватные переменные модуля

Private Type typeLayerProps
    Visible As Boolean
    Printable As Boolean
    Editable As Boolean
End Type

Private StartTime As Double

'===============================================================================
' # публичные переменные

Public Const CustomError = vbObjectError Or 32

Public Type typeMatrix
    d11 As Double
    d12 As Double
    d21 As Double
    d22 As Double
    tx As Double
    ty As Double
End Type

'===============================================================================
' # функции общего назначения

'-------------------------------------------------------------------------------
' Функции           : BoostStart, BoostFinish
' Версия            : 2022.05.31
' Авторы            : dizzy, elvin-nsk
' Назначение        : доработанные оптимизаторы от CtC
' Зависимости       : самодостаточные
'
' Параметры:
' ~~~~~~~~~~
'
'
' Использование:
' ~~~~~~~~~~~~~~
'
'-------------------------------------------------------------------------------
Public Sub BoostStart( _
               Optional ByVal UndoGroupName As String = "", _
               Optional ByVal Optimize As Boolean = True _
           )
    If Not UndoGroupName = "" And Not ActiveDocument Is Nothing Then _
        ActiveDocument.BeginCommandGroup UndoGroupName
    If Optimize And Not Optimization Then Optimization = True
    If EventsEnabled Then EventsEnabled = False
    If Not ActiveDocument Is Nothing Then
        With ActiveDocument
            .SaveSettings
            .PreserveSelection = False
            .Unit = cdrMillimeter
            .WorldScale = 1
            .ReferencePoint = cdrCenter
        End With
    End If
End Sub
Public Sub BoostFinish(Optional ByVal EndUndoGroup As Boolean = True)
    If Not EventsEnabled Then EventsEnabled = True
    If Optimization Then Optimization = False
    If Not ActiveDocument Is Nothing Then
        With ActiveDocument
            .RestoreSettings
            .PreserveSelection = True
            If EndUndoGroup Then .EndCommandGroup
        End With
        ActiveWindow.Refresh
    End If
    Application.Windows.Refresh
End Sub

Public Sub Throw(Optional ByVal Message As String = "Неизвестная ошибка")
    VBA.Err.Raise CustomError, , Message
End Sub

'===============================================================================
' # функции поиска и получения информации об объектах корела

'возвращает среднее сторон шейпа/рэйнджа/страницы
Public Property Get AverageDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="AverageDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Property
    End If
    AverageDim = (ShapeOrRangeOrPage.SizeWidth + ShapeOrRangeOrPage.SizeHeight) _
               / 2
End Property

'находит все шейпы, включая шейпы в поверклипах, с рекурсией,
'опционально исключая шейпы-поверклипы и шейпы-группы
Public Property Get FindAllShapes( _
                    ByVal Shapes As ShapeRange, _
                    Optional ExcludeGroupShapes As Boolean = False, _
                    Optional ExcludePowerClipShapes As Boolean = False _
                ) As ShapeRange
    Dim Shape As Shape
    Set FindAllShapes = CreateShapeRange
    FindAllShapes.AddRange Shapes.Shapes.FindShapes
    For Each Shape In FindPowerClips(Shapes)
        FindAllShapes.AddRange FindAllShapes(Shape.PowerClip.Shapes.All)
    Next Shape
    If ExcludeGroupShapes Then
        FindAllShapes.RemoveRange _
            FindAllShapes.Shapes.FindShapes(Type:=cdrGroupShape)
    End If
    If ExcludePowerClipShapes Then
        FindAllShapes.RemoveRange _
            FindPowerClips(FindAllShapes)
    End If
End Property

'находит все шейпы с данным именем, включая шейпы в поверклипах, с рекурсией
Public Property Get FindShapesByName( _
                    ByVal Shapes As ShapeRange, _
                    ByVal Name As String _
                ) As ShapeRange
    Set FindShapesByName = FindAllShapes(Shapes).Shapes.FindShapes(Name)
End Property

'находит все шейпы, часть имени которых совпадает с NamePart,
'включая шейпы в поверклипах, с рекурсией
Public Property Get FindShapesByNamePart( _
                    ByVal Shapes As ShapeRange, _
                    ByVal NamePart As String _
                ) As ShapeRange
    Set FindShapesByNamePart = FindAllShapes(Shapes).Shapes.FindShapes( _
                                   Query:="@Name.Contains('" & NamePart & "')" _
                               )
End Property

'находит поверклипы, без рекурсии
Public Property Get FindPowerClips(ByVal Shapes As ShapeRange) As ShapeRange
    Set FindPowerClips = CreateShapeRange
    'On Error Resume Next
    'FindPowerClips.AddRange Shapes.Shapes.FindShapes(Query:="!@com.PowerClip.IsNull")
    Dim Shape As Shape
    For Each Shape In Shapes
        If Not lib_elvin.IsNothing(Shape) Then _
            If Not Shape.PowerClip Is Nothing Then FindPowerClips.Add Shape
    Next Shape
End Property

'находит содержимое поверклипов, без рекурсии
Public Property Get FindShapesInPowerClips( _
                        ByVal Shapes As ShapeRange _
                    ) As ShapeRange
    Dim Shape As Shape
    Set FindShapesInPowerClips = CreateShapeRange
    For Each Shape In FindPowerClips(Shapes)
        FindShapesInPowerClips.AddRange Shape.PowerClip.Shapes.All
    Next Shape
End Property

'отсюда: https://community.coreldraw.com/talk/coreldraw_graphics_suite_x4/f/coreldraw-graphics-suite-x4/57576/macro-list-fonts-within-a-text-file
Public Sub FindFontsInRange( _
                ByVal TextRange As TextRange, _
                ByVal ioFonts As Collection _
            )
    Dim FontName As String
    Dim Before As TextRange, After As TextRange
    FontName = TextRange.Font
    If FontName = "" Then
        ' There are more than one font in the range
        ' Divide the range in two and look into each half separately
        ' to see if any of them has the same font. Repeat recursively
        Set Before = TextRange.Duplicate
        Before.End = (Before.Start + Before.End) \ 2
        Set After = TextRange.Duplicate
        After.Start = Before.End
        FindFontsInRange Before, ioFonts
        FindFontsInRange After, ioFonts
    Else
        AddFontToCollection FontName, ioFonts
    End If
End Sub
'+++
Private Sub AddFontToCollection( _
                ByVal FontName As String, _
                ByVal ioFonts As Collection _
            )
    Dim Font As Variant
    Dim Found As Boolean
    Found = False
    For Each Font In ioFonts
        If Font = FontName Then
            Found = True
            Exit For
        End If
    Next Font
    If Not Found Then ioFonts.Add FontName
End Sub

Public Property Get DiffWithinTolerance( _
                     ByVal Number1 As Variant, _
                     ByVal Number2 As Variant, _
                     ByVal Tolerance As Variant _
                 ) As Boolean
    DiffWithinTolerance = VBA.Abs(Number1 - Number2) < Tolerance
End Property

'возвращает все шейпы на всех слоях текущей страницы, по умолчанию - без мастер-слоёв и без гайдов
Public Property Get FindShapesActivePageLayers( _
                    Optional ByVal GuidesLayers As Boolean, _
                    Optional ByVal MasterLayers As Boolean _
                ) As ShapeRange
    Dim tLayer As Layer
    Set FindShapesActivePageLayers = CreateShapeRange
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
End Property

Public Property Get FindShapesWithText( _
                     ByVal Source As ShapeRange, _
                     ByVal Text As String _
                 ) As ShapeRange
    Dim TextShapes As ShapeRange
    Set TextShapes = Source.Shapes.FindShapes(Type:=cdrTextShape)
    Set FindShapesWithText = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In TextShapes
        If VBA.InStr(1, Shape.Text.Story.Text, Text, vbTextCompare) > 0 Then _
            FindShapesWithText.Add Shape
    Next Shape
End Property

'возвращает коллекцию слоёв с текущей страницы, имена которых включают NamePart
Public Property Get FindLayersActivePageByNamePart( _
                    ByVal NamePart As String, _
                    Optional ByVal SearchMasters = True _
                ) As Collection
    Dim tLayer As Layer
    Dim tLayers As Layers
    If SearchMasters Then
        Set tLayers = ActivePage.AllLayers
    Else
        Set tLayers = ActivePage.Layers
    End If
    Set FindLayersActivePageByNamePart = New Collection
    For Each tLayer In tLayers
        If InStr(tLayer.Name, NamePart) > 0 Then _
            FindLayersActivePageByNamePart.Add tLayer
    Next
End Property

'найти дубликат слоя по ряду параметров (достовернее, чем поиск по имени)
Public Property Get FindLayerDuplicate( _
                    ByVal PageToSearch As Page, _
                    ByVal SrcLayer As Layer _
                ) As Layer
    For Each FindLayerDuplicate In PageToSearch.AllLayers
        With FindLayerDuplicate
            If (.Name = SrcLayer.Name) And _
                 (.IsDesktopLayer = SrcLayer.IsDesktopLayer) And _
                 (.Master = SrcLayer.Master) And _
                 (.Color.IsSame(SrcLayer.Color)) Then _
                 Exit Property
        End With
    Next
    Set FindLayerDuplicate = Nothing
End Property

Public Property Get GetAverageColor(ByVal Colors As Collection) As Color
    If Colors.Count = 0 Then
        Throw "No colors in colors collection"
        Exit Property
    End If
    If Colors.Count = 1 Then
        Set GetAverageColor = Colors(1).GetCopy
        Exit Property
    End If
    Dim Index As Long
    For Index = 1 To Colors.Count
        Set GetAverageColor = _
            GetMixedColor( _
                GetAverageColor, _
                Colors(Index), _
                100 - (100 / Index) _
            )
    Next Index
End Property

Public Property Get GetAverageColorFromShapes( _
                    ByVal Shapes As ShapeRange, _
                    Optional ByVal Fills As Boolean = True, _
                    Optional ByVal Outlines As Boolean = True _
                ) As Color
    On Error GoTo NoColor
    Set GetAverageColorFromShapes = GetAverageColor( _
        GetBoundColors( _
            Shapes:=Shapes, _
            Fills:=Fills, _
            Outlines:=Outlines _
        ) _
    )
NoColor:
End Property

Public Property Get GetBoundColors( _
                    ByVal Shapes As ShapeRange, _
                    Optional ByVal Fills As Boolean = True, _
                    Optional ByVal Outlines As Boolean = True _
                ) As Collection
    Set GetBoundColors = New Collection
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.CreateSelection
        If Fills Then _
            AppendCollection GetBoundColors, GetBoundColorsFromFill(Shape)
        If Outlines Then _
            If ShapeHasOutline(Shape) Then _
                GetBoundColors.Add Shape.Outline.Color
    Next Shape
End Property

Public Property Get GetBoundColorsFromFill( _
                    ByVal Shape As Shape _
                ) As Collection
    Set GetBoundColorsFromFill = New Collection
    With Shape.Fill
        If Shape.Fill.Type = cdrUniformFill Then
            GetBoundColorsFromFill.Add Shape.Fill.UniformColor
        ElseIf Shape.Fill.Type = cdrFountainFill Then
            AppendCollection GetBoundColorsFromFill, _
                             GetBoundColorsFromFountain(Shape)
        ElseIf Shape.Fill.Type = cdrPatternFill Then
            If Shape.Fill.Pattern.Type = cdrTwoColorPattern Then
                AppendCollection GetBoundColorsFromFill, _
                                 GetBoundColorsFromTwoColorPattern(Shape)
            End If
        End If
    End With
End Property

Public Property Get GetBoundColorsFromFountain( _
                    ByVal Shape As Shape _
                ) As Collection
    Set GetBoundColorsFromFountain = New Collection
    Dim FColor As FountainColor
    For Each FColor In Shape.Fill.Fountain.Colors
        GetBoundColorsFromFountain.Add FColor.Color
    Next FColor
End Property

Public Property Get GetBoundColorsFromTwoColorPattern( _
                     ByVal Shape As Shape _
                 ) As Collection
    Set GetBoundColorsFromTwoColorPattern = New Collection
    GetBoundColorsFromTwoColorPattern.Add Shape.Fill.Pattern.FrontColor
    GetBoundColorsFromTwoColorPattern.Add Shape.Fill.Pattern.BackColor
End Property

Public Property Get GetBottomOrderShape(ByVal Shapes As ShapeRange) As Shape
    If Shapes.Count = 0 Then Exit Property
    Set GetBottomOrderShape = Shapes(1)
    If Shapes.Count = 1 Then Exit Property
    Dim Index As Long
    For Index = 2 To Shapes.Count
        If Shapes(Index).ZOrder > GetBottomOrderShape.ZOrder Then
            Set GetBottomOrderShape = Shapes(Index)
        End If
    Next Index
End Property

Public Property Get GetColorLightness(ByVal Color As Color) As Long
    Dim GrayScale As Color
    Set GrayScale = Color.GetCopy
    GrayScale.ConvertToGray
    GetColorLightness = GrayScale.Gray
End Property

Public Property Get GetHeightKeepProportions( _
                    ByVal Rect As Rect, _
                    ByVal Width As Double _
                ) As Double
    Dim WidthToHeight As Double
    WidthToHeight = Rect.Width / Rect.Height
    GetHeightKeepProportions = Width / WidthToHeight
End Property

Public Property Get GetMixedColor( _
                    ByVal MaybeColor1 As Color, _
                    ByVal MaybeColor2 As Color, _
                    Optional ByVal MixRatio As Long = 50 _
                ) As Color
    If MaybeColor1 Is Nothing And MaybeColor2 Is Nothing Then Exit Property
    If MaybeColor1 Is Nothing Then
        Set GetMixedColor = MaybeColor2.GetCopy
        Exit Property
    ElseIf MaybeColor2 Is Nothing Then
        Set GetMixedColor = MaybeColor1.GetCopy
        Exit Property
    End If
    Set GetMixedColor = MaybeColor1.GetCopy
    GetMixedColor.BlendWith MaybeColor2, MixRatio
End Property

Public Property Get GetTopOrderShape(ByVal Shapes As ShapeRange) As Shape
    If Shapes.Count = 0 Then Exit Property
    Set GetTopOrderShape = Shapes(1)
    If Shapes.Count = 1 Then Exit Property
    Dim Index As Long
    For Index = 2 To Shapes.Count
        If Shapes(Index).ZOrder < GetTopOrderShape.ZOrder Then
            Set GetTopOrderShape = Shapes(Index)
        End If
    Next Index
End Property

Public Property Get GetWidthKeepProportions( _
                    ByVal Rect As Rect, _
                    ByVal Height As Double _
                ) As Double
    Dim WidthToHeight As Double
    WidthToHeight = Rect.Width / Rect.Height
    GetWidthKeepProportions = Height * WidthToHeight
End Property

'возвращает бОльшую сторону шейпа/рэйнджа/страницы
Public Property Get GreaterDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape And Not TypeOf ShapeOrRangeOrPage Is ShapeRange And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="GreaterDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Property
    End If
    If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then
        GreaterDim = ShapeOrRangeOrPage.SizeWidth
    Else
        GreaterDim = ShapeOrRangeOrPage.SizeHeight
    End If
End Property

'является ли шейп/рэйндж/страница альбомным
Public Property Get IsLandscape(ByVal ShapeOrRangeOrPage As Object) As Boolean
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="IsLandscape", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Property
    End If
    If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then
        IsLandscape = True
    Else
        IsLandscape = False
    End If
End Property

'тестирует на пустой кореловский объект
'для пустого объекта коллекции,
'т. к. для Nothing ошибка может быть уже на этапе вызова
Public Property Get IsNothing(ByVal Object As Object) As Boolean
    Dim t As Variant
    If Object Is Nothing Then GoTo ExitTrue
    If TypeOf Object Is Document Then
        On Error GoTo ExitTrue
        t = Object.Name
    ElseIf TypeOf Object Is Page Then
        On Error GoTo ExitTrue
        t = Object.Name
    ElseIf TypeOf Object Is Layer Then
        On Error GoTo ExitTrue
        t = Object.Name
    ElseIf TypeOf Object Is Shape Then
        On Error GoTo ExitTrue
        t = Object.Name
    ElseIf TypeOf Object Is Curve Then
        On Error GoTo ExitTrue
        t = Object.Length
    ElseIf TypeOf Object Is SubPath Then
        On Error GoTo ExitTrue
        t = Object.Closed
    ElseIf TypeOf Object Is Segment Then
        On Error GoTo ExitTrue
        t = Object.AbsoluteIndex
    ElseIf TypeOf Object Is Node Then
        On Error GoTo ExitTrue
        t = Object.AbsoluteIndex
    End If
    Exit Property
ExitTrue:
    IsNothing = True
End Property

'todo: ПРОВЕРИТЬ КАК СЛЕДУЕТ
Public Property Get IsOverlap( _
                    ByVal FirstShape As Shape, _
                    ByVal SecondShape As Shape _
                ) As Boolean
    
    Dim tIS As Shape
    Dim tShape1 As Shape, tShape2 As Shape
    Dim tBound1 As Shape, tBound2 As Shape
    Dim tProps As typeLayerProps
    
    If FirstShape.Type = cdrConnectorShape _
    Or SecondShape.Type = cdrConnectorShape Then _
        Exit Property
    
    'запоминаем какой слой был активным
    Dim tLayer As Layer: Set tLayer = ActiveLayer
    'запоминаем состояние первого слоя
    FirstShape.Layer.Activate
    LayerPropsPreserveAndReset FirstShape.Layer, tProps
    
    If IsIntersectReady(FirstShape) Then
        Set tShape1 = FirstShape
    Else
        Set tShape1 = CreateBoundary(FirstShape)
        Set tBound1 = tShape1
    End If
    
    If IsIntersectReady(SecondShape) Then
        Set tShape2 = SecondShape
    Else
        Set tShape2 = CreateBoundary(SecondShape)
        Set tBound2 = tShape2
    End If
    
    Set tIS = tShape1.Intersect(tShape2)
    If tIS Is Nothing Then
        IsOverlap = False
    Else
        tIS.Delete
        IsOverlap = True
    End If
    
    On Error Resume Next
        tBound1.Delete
        tBound2.Delete
    On Error GoTo 0
    
    'возвращаем всё на место
    LayerPropsRestore FirstShape.Layer, tProps
    tLayer.Activate

End Property

'IsOverlap здорового человека - меряет по габаритам,
'но зато стабильно работает и в большинстве случаев его достаточно
Public Property Get IsOverlapBox( _
                    ByVal FirstShape As Shape, _
                    ByVal SecondShape As Shape _
                ) As Boolean
    Dim tShape As Shape
    Dim tProps As typeLayerProps
    'запоминаем какой слой был активным
    Dim tLayer As Layer: Set tLayer = ActiveLayer
    'запоминаем состояние первого слоя
    FirstShape.Layer.Activate
    LayerPropsPreserveAndReset FirstShape.Layer, tProps
    Dim tRect As Rect
    Set tRect = FirstShape.BoundingBox.Intersect(SecondShape.BoundingBox)
    If tRect.Width = 0 And tRect.Height = 0 Then
        IsOverlapBox = False
    Else
        IsOverlapBox = True
    End If
    'возвращаем всё на место
    LayerPropsRestore FirstShape.Layer, tProps
    tLayer.Activate
End Property

Public Property Get IsSameColor( _
                    ByVal MaybeColor1 As Variant, _
                    ByVal MaybeColor2 As Variant _
                ) As Boolean
    If VBA.IsEmpty(MaybeColor1) Or VBA.IsEmpty(MaybeColor2) Then Exit Property
    IsSameColor = MaybeColor1.IsSame(MaybeColor2)
End Property

'являются ли кривые дубликатами, находящимися друг над другом в одном месте
'(underlying dubs)
Public Property Get IsSameCurves( _
                    ByVal Curve1 As Curve, _
                    ByVal Curve2 As Curve _
                ) As Boolean
    Dim tNode As Node
    Dim Tolerance As Double
    'допуск = 0.001 мм
    Tolerance = ConvertUnits(0.001, cdrMillimeter, ActiveDocument.Unit)
    IsSameCurves = False
    If Not Curve1.Nodes.Count = Curve2.Nodes.Count Then Exit Property
    If Abs(Curve1.Length - Curve2.Length) > Tolerance Then Exit Property
    For Each tNode In Curve1.Nodes
        If Curve2.FindNodeAtPoint( _
               tNode.PositionX, _
               tNode.PositionY, _
               Tolerance * 2 _
           ) Is Nothing Then Exit Property
    Next
    IsSameCurves = True
End Property

Public Property Get IsShape(ByVal MaybeShape As Variant) As Boolean
    If VBA.IsObject(MaybeShape) Then
        If Not MaybeShape Is Nothing Then
            IsShape = TypeOf MaybeShape Is Shape
        End If
    End If
End Property

Public Property Get IsShapeRange(ByVal MaybeShapeRange As Variant) As Boolean
    If VBA.IsObject(MaybeShapeRange) Then
        If Not MaybeShapeRange Is Nothing Then
            IsShapeRange = TypeOf MaybeShapeRange Is ShapeRange
        End If
    End If
End Property

Public Property Get IsShapeType( _
                    ByVal MaybeShape As Variant, _
                    ByVal ShapeType As cdrShapeType _
                ) As Boolean
    If Not IsShape(MaybeShape) Then Exit Property
    IsShapeType = (MaybeShape.Type = ShapeType)
End Property

'возвращает меньшую сторону шейпа/рэйнджа/страницы
Public Property Get LesserDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="LesserDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Property
    End If
    If ShapeOrRangeOrPage.SizeWidth < ShapeOrRangeOrPage.SizeHeight Then
        LesserDim = ShapeOrRangeOrPage.SizeWidth
    Else
        LesserDim = ShapeOrRangeOrPage.SizeHeight
    End If
End Property

Public Property Get ShapeHasCurve(ByVal Shape As Shape) As Boolean
    On Error GoTo Fail
    ShapeHasCurve = Not (Shape.Curve Is Nothing)
Fail:
End Property

Public Property Get ShapeHasOutline(ByVal Shape As Shape) As Boolean
    On Error GoTo Fail
    ShapeHasOutline = Not (Shape.Outline.Type = cdrNoOutline)
Fail:
End Property

Public Property Get ShapeHasUniformFill(ByVal Shape As Shape) As Boolean
    On Error GoTo Fail
    ShapeHasUniformFill = (Shape.Fill.Type = cdrUniformFill)
Fail:
End Property

Public Property Get ShapeIsInGroup(ByVal Shape As Shape) As Boolean
    On Error GoTo Fail
    ShapeIsInGroup = Not (Shape.ParentGroup Is Nothing)
Fail:
End Property

'возвращает коллекцию слоёв, на которых лежат шейпы из ренджа
Public Property Get ShapeRangeLayers( _
                        ByVal ShapeRange As ShapeRange _
                    ) As Collection
    
    Dim tShape As Shape
    Dim tLayer As Layer
    Dim inCol As Boolean
    
    If ShapeRange.Count = 0 Then Exit Property
    Set ShapeRangeLayers = New Collection
    If ShapeRange.Count = 1 Then
        ShapeRangeLayers.Add ShapeRange(1).Layer
        Exit Property
    End If
    
    For Each tShape In ShapeRange
        inCol = False
        For Each tLayer In ShapeRangeLayers
            If tLayer Is tShape.Layer Then
                inCol = True
                Exit For
            End If
        Next tLayer
        If inCol = False Then ShapeRangeLayers.Add tShape.Layer
    Next tShape

End Property

'возвращает Rect, равный габаритам объекта плюс Space со всех сторон
Public Property Get SpaceBox( _
                    ByVal MaybeShapeOrRange As Variant, _
                    ByVal Space As Double _
                ) As Rect
    If VBA.IsEmpty(MaybeShapeOrRange) Then Exit Property
    ThrowIfNotShapeOrRange MaybeShapeOrRange
    Set SpaceBox = MaybeShapeOrRange.BoundingBox.GetCopy
    SpaceBox.Inflate Space, Space, Space, Space
End Property

'возвращает Outline если или Empty
Public Property Get TryGetOutline(ByVal Shape As Shape) As Variant
    If ShapeHasOutline(Shape) Then Set TryGetOutline = Shape.Outline
End Property

'===============================================================================
' # функции манипуляций с объектами корела

Public Function BreakApart(ByVal Shape As Shape) As ShapeRange
    If Shape.Curve.SubPaths.Count < 2 Then
        Set BreakApart = CreateShapeRange
        BreakApart.Add Shape
        Exit Function
    End If
    Set BreakApart = Shape.BreakApartEx
    If BreakApart.Count > 1 Then Exit Function
    Set BreakApart = CreateShapeRange
    Dim RemainingShape As Shape
    Dim ExtractedShape As Shape
    'RemainingShape и ExtractedShape в Extract
    'на самом деле наоборот, чем в спеках
    Set RemainingShape = Shape.Curve.SubPaths.First.Extract(ExtractedShape)
    BreakApart.Add ExtractedShape
    BreakApart.AddRange BreakApart(RemainingShape)
End Function

'перекрашивает объект в чёрный или белый в серой шкале,
'в зависимости от исходного цвета
'ДОРАБОТАТЬ
Public Function ContrastShape(ByVal Shape As Shape) As Shape
    With Shape.Fill
        Select Case .Type
            Case cdrUniformFill
                .UniformColor.ConvertToGray
                If .UniformColor.Gray < 128 Then
                    .UniformColor.GrayAssign 0
                Else
                    .UniformColor.GrayAssign 255
                End If
            Case cdrFountainFill
                'todo
        End Select
    End With
    With Shape.Outline
        If Not .Type = cdrNoOutline Then
            .Color.ConvertToGray
            If .Color.Gray < 128 Then _
                .Color.GrayAssign 0 Else .Color.GrayAssign 255
        End If
    End With
    Set ContrastShape = Shape
End Function

'правильно копирует Shape или ShapeRange на другой слой
Public Function CopyToLayer( _
                    ByVal ShapeOrRange As Object, _
                    ByVal Layer As Layer _
                ) As Object

    If Not TypeOf ShapeOrRange Is Shape And Not TypeOf ShapeOrRange Is ShapeRange Then
        Err.Raise 13, Source:="CopyToLayer", _
                  Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    End If
    
    Set CopyToLayer = ShapeOrRange.Duplicate
    MoveToLayer CopyToLayer, Layer

End Function

'инструмент Boundary
Public Function CreateBoundary(ByVal ShapeOrRange As Object) As Shape
    On Error GoTo Catch
    Dim tShape As Shape, tRange As ShapeRange
    'просто объект не ест, надо конкретный тип
    If TypeOf ShapeOrRange Is Shape Then
        Set tShape = ShapeOrRange
        Set CreateBoundary = tShape.CustomCommand("Boundary", "CreateBoundary")
    ElseIf TypeOf ShapeOrRange Is ShapeRange Then
        Set tRange = ShapeOrRange
        Set CreateBoundary = tRange.CustomCommand("Boundary", "CreateBoundary")
    Else
        Err.Raise 13, Source:="CreateBoundary", _
            Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    End If
    Exit Function
Catch:
    Debug.Print Err.Number
End Function

'создаёт слой, если такой слой есть - возвращает этот слой
Public Function CreateOrFindLayer( _
                    ByVal Page As Page, _
                    ByVal Name As String _
                ) As Layer
    Set CreateOrFindLayer = Page.Layers.Find(Name)
    If CreateOrFindLayer Is Nothing Then
        Set CreateOrFindLayer = Page.CreateLayer(Name)
    End If
End Function

'инструмент Crop Tool
Public Function CropTool( _
                    ByVal ShapeOrRangeOrPage As Object, _
                    ByVal x1#, ByVal y1#, _
                    ByVal x2#, ByVal y2#, _
                    Optional ByVal Angle = 0 _
                ) As ShapeRange
    If TypeOf ShapeOrRangeOrPage Is Shape Or _
         TypeOf ShapeOrRangeOrPage Is ShapeRange Or _
         TypeOf ShapeOrRangeOrPage Is Page Then
        Set CropTool = ShapeOrRangeOrPage.CustomCommand("Crop", "CropRectArea", x1, y1, x2, y2, Angle)
    Else
        Err.Raise 13, Source:="CropTool", _
            Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Function
    End If
End Function

'отрезать кусок от Shape по контуру Knife, возвращает отрезанный кусок
Public Function Dissect(ByRef Shape As Shape, ByRef Knife As Shape) As Shape
    Set Dissect = Intersect(Knife, Shape, True, True)
    Set Shape = Knife.Trim(Shape, True, False)
End Function

'дублировать активную страницу со всеми слоями и объектами
Public Function DuplicateActivePage( _
                    ByVal NumberOfPages As Long, _
                    Optional ByVal ExcludeLayerName As String = "" _
                ) As Page
    Dim tRange As ShapeRange
    Dim tShape As Shape, sDuplicate As Shape
    Dim tProps As typeLayerProps
    Dim i&
    For i = 1 To NumberOfPages
        Set tRange = FindShapesActivePageLayers
        Set DuplicateActivePage = _
            ActiveDocument.InsertPages(1, False, ActivePage.Index)
        DuplicateActivePage.SizeHeight = ActivePage.SizeHeight
        DuplicateActivePage.SizeWidth = ActivePage.SizeWidth
        For Each tShape In tRange.ReverseRange
            If tShape.Layer.Name <> ExcludeLayerName Then
                LayerPropsPreserveAndReset tShape.Layer, tProps
                Set sDuplicate = tShape.Duplicate
                sDuplicate.MoveToLayer _
                    FindLayerDuplicate(DuplicateActivePage, tShape.Layer)
                LayerPropsRestore tShape.Layer, tProps
            End If
        Next tShape
    Next i
End Function

Public Sub FillInside( _
               ByVal MaybeShapeOrRange As Variant, _
               ByVal MaybeTargetRect As Variant _
           )
    If VBA.IsEmpty(MaybeShapeOrRange) _
    Or VBA.IsEmpty(MaybeTargetRect) Then Exit Sub
    ThrowIfNotShapeOrRange MaybeShapeOrRange
    
    If GetHeightKeepProportions( _
           MaybeShapeOrRange.BoundingBox, _
           MaybeTargetRect.Width _
       ) > MaybeTargetRect.Height Then
        MaybeShapeOrRange.SetSize MaybeTargetRect.Width
    Else
        MaybeShapeOrRange.SetSize , MaybeTargetRect.Height
    End If
    MaybeShapeOrRange.CenterX = MaybeTargetRect.CenterX
    MaybeShapeOrRange.CenterY = MaybeTargetRect.CenterY
End Sub

Public Sub FitInside( _
               ByVal ShapeToFit As Shape, _
               ByVal TargetRect As Rect _
           )
    If GetHeightKeepProportions(ShapeToFit.BoundingBox, TargetRect.Width) _
     > TargetRect.Height Then
        ShapeToFit.SetSize , TargetRect.Height
    Else
        ShapeToFit.SetSize TargetRect.Width
    End If
    ShapeToFit.CenterX = TargetRect.CenterX
    ShapeToFit.CenterY = TargetRect.CenterY
End Sub

'все объекты на всех страницах, включая мастер-страницу - на один слой
'все страницы прибиваются, все объекты на слоях guides прибиваются
Public Function FlattenPagesToLayer(ByVal LayerName As String) As Layer

    Dim DL As Layer: Set DL = ActiveDocument.MasterPage.DesktopLayer
    Dim DLstate As Boolean: DLstate = DL.Editable
    Dim P As Page
    Dim L As Layer
    
    DL.Editable = False
    
    For Each P In ActiveDocument.Pages
        For Each L In P.Layers
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
        If P.Index <> 1 Then P.Delete
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

'правильный интерсект
Public Function Intersect( _
                    ByVal SourceShape As Shape, _
                    ByVal TargetShape As Shape, _
                    Optional ByVal LeaveSource As Boolean = True, _
                    Optional ByVal LeaveTarget As Boolean = True _
                ) As Shape
                                     
    Dim tPropsSource As typeLayerProps
    Dim tPropsTarget As typeLayerProps
    
    If Not SourceShape.Layer Is TargetShape.Layer Then _
        LayerPropsPreserveAndReset SourceShape.Layer, tPropsSource
    LayerPropsPreserveAndReset TargetShape.Layer, tPropsTarget
    
    Set Intersect = SourceShape.Intersect(TargetShape)
    
    If Not SourceShape.Layer Is TargetShape.Layer Then _
        LayerPropsRestore SourceShape.Layer, tPropsSource
    LayerPropsRestore TargetShape.Layer, tPropsTarget
    
    If Intersect Is Nothing Then Exit Function
    
    Intersect.OrderFrontOf TargetShape
    If Not LeaveSource Then SourceShape.Delete
    If Not LeaveTarget Then TargetShape.Delete

End Function

'инструмент Join Curves
Public Function JoinCurves(ByVal ShapeOrShapes As Variant, ByVal Tolerance As Double)
    ShapeOrShapes.CustomCommand "ConvertTo", "JoinCurves", Tolerance
End Function

'не работает с поверклипом
Public Sub MatrixCopy(ByVal SourceShape As Shape, ByVal TargetShape As Shape)
    Dim tMatrix As typeMatrix
    With tMatrix
        SourceShape.GetMatrix .d11, .d12, .d21, .d22, .tx, .ty
        TargetShape.SetMatrix .d11, .d12, .d21, .d22, .tx, .ty
    End With
End Sub

'правильно перемещает Shape или ShapeRange на другой слой
Public Function MoveToLayer( _
                    ByVal MaybeShapeOrRange As Variant, _
                    ByVal MaybeLayer As Layer _
                )
    If VBA.IsEmpty(MaybeShapeOrRange) _
    Or VBA.IsEmpty(MaybeLayer) Then Exit Function
    ThrowIfNotShapeOrRange MaybeShapeOrRange
    
    Dim tSrcLayer() As Layer
    Dim tProps() As typeLayerProps
    Dim tLayersCol As Collection
    Dim i&
    
    If TypeOf MaybeShapeOrRange Is Shape Then
    
        Set tLayersCol = New Collection
        tLayersCol.Add MaybeShapeOrRange.Layer
        
    ElseIf TypeOf MaybeShapeOrRange Is ShapeRange Then
        
        If MaybeShapeOrRange.Count < 1 Then Exit Function
        Set tLayersCol = ShapeRangeLayers(MaybeShapeOrRange)
        
    Else
    
        Throw "Type mismatch: MaybeShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    
    End If
    
    ReDim tSrcLayer(1 To tLayersCol.Count)
    ReDim tProps(1 To tLayersCol.Count)
    For i = 1 To tLayersCol.Count
        Set tSrcLayer(i) = tLayersCol(i)
        LayerPropsPreserveAndReset tSrcLayer(i), tProps(i)
    Next i
    MaybeShapeOrRange.MoveToLayer MaybeLayer
    For i = 1 To tLayersCol.Count
        LayerPropsRestore tSrcLayer(i), tProps(i)
    Next i

End Function

'удаление сегмента
'автор: Alex Vakulenko http://www.oberonplace.com/vba/drawmacros/delsegment.htm
Public Sub SegmentDelete(ByVal Segment As Segment)
    If Not Segment.EndNode.IsEnding Then
        Segment.EndNode.BreakApart
        Set Segment = Segment.SubPath.LastSegment
    End If
    Segment.EndNode.Delete
End Sub

'присвоить цвет абриса ренджу
Public Sub SetOutlineColor( _
               ByVal Shapes As ShapeRange, _
               ByVal Color As Color _
           )
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.Outline.Color.CopyAssign Color
    Next Shape
End Sub

Public Sub Trim( _
               ByVal TrimmerShape As Shape, _
               ByRef TargetShape As Shape _
           )
    Set TargetShape = TrimmerShape.Trim(TargetShape)
End Sub

'обрезать битмап по CropEnvelopeShape, но по-умному,
'сначала кропнув на EXPANDBY пикселей побольше
Public Function TrimBitmap( _
                    ByVal BitmapShape As Shape, _
                    ByVal CropEnvelopeShape As Shape, _
                    Optional ByVal LeaveCropEnvelope As Boolean = True _
                ) As Shape

    Const EXPANDBY& = 2 'px
    
    Dim tCrop As Shape
    Dim tPxW#, tPxH#
    Dim tSaveUnit As cdrUnit

    If Not BitmapShape.Type = cdrBitmapShape Then Exit Function
    
    'save
    tSaveUnit = ActiveDocument.Unit
    
    ActiveDocument.Unit = cdrInch
    tPxW = 1 / BitmapShape.Bitmap.ResolutionX
    tPxH = 1 / BitmapShape.Bitmap.ResolutionY
    BitmapShape.Bitmap.ResetCropEnvelope
    Set tCrop = BitmapShape.Layer.CreateRectangle( _
                    CropEnvelopeShape.LeftX - tPxW * EXPANDBY, _
                    CropEnvelopeShape.TopY + tPxH * EXPANDBY, _
                    CropEnvelopeShape.RightX + tPxW * EXPANDBY, _
                    CropEnvelopeShape.BottomY - tPxH * EXPANDBY _
                )
    Set TrimBitmap = Intersect(tCrop, BitmapShape, False, False)
    If TrimBitmap Is Nothing Then
        tCrop.Delete
        GoTo Finally
    End If
    TrimBitmap.Bitmap.Crop
    Set TrimBitmap = _
        Intersect(CropEnvelopeShape, TrimBitmap, LeaveCropEnvelope, False)
    
Finally:
    'restore
    ActiveDocument.Unit = tSaveUnit
    
End Function

Public Function Weld(ByVal Shapes As ShapeRange) As Shape
    Set Weld = Shapes.FirstShape
    Do Until Shapes.Count = 1
        Shapes(1).CreateSelection
        Shapes(2).AddToSelection
        Shapes.Remove 1
        Shapes.Remove 1
        With ActiveSelectionRange
            Set Weld = .FirstShape.Weld(.LastShape)
        End With
        Shapes.Add Weld
    Loop
End Function

'===============================================================================
' # функции работы с файлами

Public Function AddProperEndingToPath(ByVal Path As String) As String
    If Not VBA.Right$(Path, 1) = "\" Then AddProperEndingToPath = Path & "\" _
    Else: AddProperEndingToPath = Path
End Function

'существует ли файл или папка (папка должна заканчиваться на "\")
Public Property Get FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Property
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Property

Public Property Get FindFileInGMSFolders(ByVal FileName As String) As String
    FindFileInGMSFolders = GMSManager.UserGMSPath & FileName
    If Not FileExists(FindFileInGMSFolders) Then _
        FindFileInGMSFolders = GMSManager.GMSPath & FileName
    If Not FileExists(FindFileInGMSFolders) Then _
        FindFileInGMSFolders = ""
End Property

'возвращает имя файла без расширения
Public Property Get GetFileNameNoExt(ByVal FileName As String) As String
    If VBA.Right(FileName, 1) <> "\" And VBA.Len(FileName) > 0 Then
        GetFileNameNoExt = Left(FileName, _
            Switch _
                (InStr(FileName, ".") = 0, _
                    Len(FileName), _
                InStr(FileName, ".") > 0, _
                    InStrRev(FileName, ".") - 1))
    End If
End Property

'полное имя временного файла
Public Property Get GetTempFile() As String
    GetTempFile = GetTempFolder & GetTempFileName
End Property

'имя временного файла
Public Property Get GetTempFileName() As String
    GetTempFileName = "elvin_" & CreateGUID & ".tmp"
End Property

'находит временную папку
Public Property Get GetTempFolder() As String
    GetTempFolder = AddProperEndingToPath(VBA.Environ$("TEMP"))
    If FileExists(GetTempFolder) Then Exit Property
    GetTempFolder = AddProperEndingToPath(VBA.Environ$("TMP"))
    If FileExists(GetTempFolder) Then Exit Property
    GetTempFolder = "c:\temp\"
    If FileExists(GetTempFolder) Then Exit Property
    GetTempFolder = "c:\windows\temp\"
    If FileExists(GetTempFolder) Then Exit Property
End Property

Public Property Get GetFileName(ByVal File As String) As String
    GetFileName = VBA.Right(File, VBA.Len(File) - VBA.InStrRev(File, "\"))
End Property

Public Property Get GetFilePath(ByVal File As String) As String
    GetFilePath = VBA.Left(File, VBA.InStrRev(File, "\"))
End Property

'создаёт папку, если не было
'возвращает Path обратно (для inline-использования)
Public Function MakeDir(ByVal Path As String) As String
    If VBA.Dir(Path, vbDirectory) = "" Then MkDir Path
    MakeDir = Path
End Function

'загружает файл в строку
Public Function ReadFile(ByVal File As String) As String
    Dim tFileNum As Long
    tFileNum = FreeFile
    Open File For Input As #tFileNum
    ReadFile = Input(LOF(tFileNum), tFileNum)
    Close #tFileNum
End Function

'загружает файл в строку через ADODB, можно задать кодировку
Public Function ReadFileAD( _
                    ByVal File As String, _
                    Optional ByVal CharSet As String = "utf-8" _
                ) As String
    Dim ADODB As Object
    Set ADODB = VBA.CreateObject("ADODB.Stream")
    ADODB.CharSet = CharSet
    ADODB.Open
    ADODB.LoadFromFile File
    ReadFileAD = ADODB.ReadText()
    ADODB.Close
End Function

'заменяет расширение файлу на заданное
Public Function SetFileExt( _
                    ByVal SourceFile As String, _
                    ByVal NewExt As String _
                ) As String
    If Right(SourceFile, 1) <> "\" And Len(SourceFile) > 0 Then
        SetFileExt = GetFileNameNoExt(SourceFile$) & "." & NewExt
    End If
End Function

'сохраняет строку Content в файл, перезаписывая, делая в процессе temp файл,
'и оставляя бэкап, если необходимо
Public Sub WriteFile( _
               ByVal Content As String, _
               ByVal File As String, _
               Optional ByVal KeepBak As Boolean = False _
           )

    Dim tFileNum As Long
    tFileNum = FreeFile
    Dim tBak As String
    tBak = SetFileExt(File, "bak")
    Dim tTemp As String
    
    If KeepBak Then
        If FileExists(File) Then FileCopy File, tBak
    Else
        If FileExists(File) Then
            tTemp = GetFilePath(File) & GetTempFileName
            FileCopy File, tTemp
        End If
    End If
        
    Open File For Output Access Write As #tFileNum
    Print #tFileNum, Content
    Close #tFileNum
    
    On Error Resume Next
        If Not KeepBak Then Kill tTemp
    On Error GoTo 0

End Sub

'сохраняет строку Content в файл через ADODB, можно задать кодировку
Public Sub WriteFileAD( _
               ByVal File As String, _
               ByVal Content As String, _
               Optional CharSet As String = "utf-8" _
           )
    Dim ADODB As Object
    Set ADODB = VBA.CreateObject("ADODB.Stream")
    ADODB.CharSet = CharSet
    ADODB.Open
    ADODB.WriteText Content
    ADODB.SaveToFile File
    ADODB.Close
End Sub

'===============================================================================
' # прочие функции

Public Sub AppendCollection( _
               ByVal Destination As Collection, _
               ByVal SourceToAdd As Collection _
           )
    Dim Item As Variant
    For Each Item In SourceToAdd
        Destination.Add Item
    Next Item
End Sub

Public Sub Assign(ByRef Destination As Variant, ByRef Value As Variant)
    If VBA.IsObject(Value) Then
        Set Destination = Value
    Else
        Destination = Value
    End If
End Sub

Public Function GetCollectionCopy(ByVal Source As Collection) As Collection
    Set GetCollectionCopy = New Collection
    Dim Item As Variant
    For Each Item In Source
        GetCollectionCopy.Add Item
    Next Item
End Function

Public Function GetCollectionFromDictionary( _
                    ByVal Dictionary As Scripting.IDictionary _
                ) As Collection
    Set GetCollectionFromDictionary = New Collection
    Dim Item As Variant
    For Each Item In Dictionary.Items
        GetCollectionFromDictionary.Add Item
    Next Item
End Function

Public Function GetDictionaryCopy( _
                    ByVal Source As Scripting.IDictionary _
                ) As Scripting.Dictionary
    Set GetDictionaryCopy = New Scripting.Dictionary
    Dim Key As Variant
    For Each Key In Source.Keys
        GetDictionaryCopy.Add Key, Source.Item(Key)
    Next Key
End Function

Public Property Get Count(ByRef Arr As Variant) As Long
    Count = UBound(Arr) - LBound(Arr) + 1
End Property

Public Property Get CountCopy(ByVal Arr As Variant) As Long
    CountCopy = UBound(Arr) - LBound(Arr) + 1
End Property

'https://www.codegrepper.com/code-examples/vb/excel+vba+generate+guid+uuid
Public Function CreateGUID( _
                    Optional ByVal Lowercase As Boolean, _
                    Optional ByVal Parens As Boolean _
                ) As String
    Dim k As Long, H As String
    CreateGUID = VBA.Space(36)
    For k = 1 To VBA.Len(CreateGUID)
        VBA.Randomize
        Select Case k
            Case 9, 14, 19, 24:         H = "-"
            Case 15:                    H = "4"
            Case 20:                    H = VBA.Hex(VBA.Rnd * 3 + 8)
            Case Else:                  H = VBA.Hex(VBA.Rnd * 15)
        End Select
        Mid(CreateGUID, k, 1) = H
    Next
    If Lowercase Then CreateGUID = VBA.LCase$(CreateGUID)
    If Parens Then CreateGUID = "{" & CreateGUID & "}"
End Function

Public Property Get FindMaxItemNum(ByVal Collection As Collection) As Long
    FindMaxItemNum = 1
    Dim i As Long
    For i = 1 To Collection.Count
        If VBA.IsNumeric(Collection(i)) Then
            If Collection(i) > Collection(FindMaxItemNum) Then _
                FindMaxItemNum = i
        End If
    Next i
End Property

Public Property Get FindMinItemNum(ByVal Collection As Collection) As Long
    FindMinItemNum = 1
    Dim i As Long
    For i = 1 To Collection.Count
        If VBA.IsNumeric(Collection(i)) Then
            If Collection(i) < Collection(FindMinItemNum) Then _
                FindMinItemNum = i
        End If
    Next i
End Property

'является ли число чётным :) Что такое Even и Odd запоминать лень...
Public Property Get IsChet(ByVal X As Variant) As Boolean
    If X Mod 2 = 0 Then IsChet = True Else IsChet = False
End Property

'делится ли Number на Divider нацело
Public Property Get IsDivider( _
                    ByVal Number As Long, _
                    ByVal Divider As Long _
                ) As Boolean
    If Number Mod Divider = 0 Then IsDivider = True Else IsDivider = False
End Property

Public Property Get IsLowerCase(ByVal Str As String) As Boolean
    If VBA.LCase(Str) = Str Then IsLowerCase = True
End Property

Public Property Get IsSame( _
                        ByRef Value1 As Variant, _
                        ByRef Value2 As Variant _
                    ) As Boolean
    If VBA.IsObject(Value1) And VBA.IsObject(Value2) Then
        IsSame = Value1 Is Value2
    ElseIf Not VBA.IsObject(Value1) And Not VBA.IsObject(Value2) Then
        IsSame = (Value1 = Value2)
    End If
End Property

Public Property Get IsUpperCase(ByVal Str As String) As Boolean
    If VBA.UCase(Str) = Str Then IsUpperCase = True
End Property

Public Sub RemoveElementFromCollection( _
               ByVal Collection As Collection, _
               ByVal Element As Variant _
           )
    If Collection.Count = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To Collection.Count
        If IsSame(Element, Collection(i)) Then
            Collection.Remove i
            Exit Sub
        End If
    Next i
End Sub

Public Property Get Max(ByRef EnumerableSet As Variant) As Variant
    Dim Item As Variant
    For Each Item In EnumerableSet
        If VBA.IsNumeric(Item) Then
            If Item > Max Then Max = Item
        End If
    Next Item
End Property

Public Property Get MaxOfTwo( _
                        ByVal Value1 As Variant, _
                        ByVal Value2 As Variant _
                    ) As Variant
    If Value1 > Value2 Then MaxOfTwo = Value1 Else MaxOfTwo = Value2
End Property

Public Function MeasureStart()
    StartTime = Timer
End Function
Public Function MeasureFinish(Optional ByVal Message As String = "")
    Debug.Print Message & VBA.CStr(Round(Timer - StartTime, 3)) & " секунд"
End Function

Public Property Get Min(ByRef EnumerableSet As Variant) As Variant
    Dim Item As Variant
    For Each Item In EnumerableSet
        If VBA.IsNumeric(Item) Then
            If Item < Min Then
                Min = Item
            End If
        End If
    Next Item
End Property

Public Property Get MinOfTwo( _
                        ByVal Value1 As Variant, _
                        ByVal Value2 As Variant _
                    ) As Variant
    If Value1 < Value2 Then MinOfTwo = Value1 Else MinOfTwo = Value2
End Property

Public Property Get Contains( _
                        ByRef EnumerableSet As Variant, _
                        ByRef Items As Variant _
                    ) As Boolean
    Dim Element As Variant
    Dim Item As Variant
    Dim ItemExists As Boolean
    For Each Item In Items
        For Each Element In EnumerableSet
            If IsSame(Item, Element) Then
                ItemExists = True
                Exit For
            End If
        Next Element
        If Not ItemExists Then Exit Property
        ItemExists = False
    Next Item
    Contains = True
End Property

Public Property Get Pack(ParamArray Items() As Variant) As Variant()
    Dim Length As Long
    Length = UBound(Items) - LBound(Items) + 1
    If Length = 0 Then
        Pack = Array()
        Exit Property
    End If
    ReDim Result(1 To Length) As Variant
    Dim Index As Long
    Dim Item As Variant
    For Each Item In Items
        Index = Index + 1
        Assign Result(Index), Item
    Next Item
    Pack = Result
End Property

'создаёт ShapeRange из Shape/Shapes/ShapeRange
Public Function PackShapes(ParamArray Shapes() As Variant) As ShapeRange
    Set PackShapes = CreateShapeRange
    Dim Item As Variant
    For Each Item In Shapes
        If TypeOf Item Is Shape Then
            PackShapes.Add Item
        ElseIf TypeOf Item Is ShapeRange Then
            PackShapes.AddRange Item
        ElseIf TypeOf Item Is Shapes Then
            PackShapes.AddRange Item.All
        Else
            Throw "Не является шейпом"
        End If
    Next Item
End Function

Private Sub Resize(ByRef Arr As Variant, ByVal Length As Long)
    ReDim Preserve Arr(LBound(Arr) To LBound(Arr) + Length - 1)
End Sub

'случайное целое от LowerBound до UpperBound
Public Property Get RndInt( _
                    ByVal LowerBound As Long, _
                    ByVal UpperBound As Long _
                ) As Long
    RndInt = VBA.Int((UpperBound - LowerBound + 1) * VBA.Rnd + LowerBound)
End Property

'===============================================================================
' # приватные функции модуля

Private Sub LayerPropsPreserve(ByVal L As Layer, ByRef Props As typeLayerProps)
    With Props
        .Visible = L.Visible
        .Printable = L.Printable
        .Editable = L.Editable
    End With
End Sub
Private Sub LayerPropsReset(ByVal L As Layer)
    With L
        If Not .Visible Then .Visible = True
        If Not .Printable Then .Printable = True
        If Not .Editable Then .Editable = True
    End With
End Sub
Private Sub LayerPropsRestore(ByVal L As Layer, ByRef Props As typeLayerProps)
    With Props
        If L.Visible <> .Visible Then L.Visible = .Visible
        If L.Printable <> .Printable Then L.Printable = .Printable
        If L.Editable <> .Editable Then L.Editable = .Editable
    End With
End Sub
Private Sub LayerPropsPreserveAndReset( _
                ByVal L As Layer, _
                ByRef Props As typeLayerProps _
            )
    LayerPropsPreserve L, Props
    LayerPropsReset L
End Sub

'для IsOverlap
Private Function IsIntersectReady(ByVal Shape As Shape) As Boolean
    With Shape
        If .Type = cdrCustomShape Or _
             .Type = cdrBlendGroupShape Or _
             .Type = cdrOLEObjectShape Or _
             .Type = cdrExtrudeGroupShape Or _
             .Type = cdrContourGroupShape Or _
             .Type = cdrBevelGroupShape Or _
             .Type = cdrConnectorShape Or _
             .Type = cdrMeshFillShape Or _
             .Type = cdrTextShape Then
            IsIntersectReady = False
        Else
            IsIntersectReady = True
        End If
    End With
End Function

Private Sub ThrowIfNotShapeOrRange( _
                ByVal MaybeShapeOrRange As Variant _
            )
    If VBA.IsObject(MaybeShapeOrRange) Then
        If Not MaybeShapeOrRange Is Nothing Then
            If TypeOf MaybeShapeOrRange Is Shape _
            Or TypeOf MaybeShapeOrRange Is ShapeRange Then _
                Exit Sub
        End If
    End If
    Throw "Тип должен быть Shape или ShapeRange"
End Sub

Private Sub ThrowIfNotCollectionOrArray(ByRef CollectionOrArray As Variant)
    If VBA.IsObject(CollectionOrArray) Then _
        If TypeOf CollectionOrArray Is Collection Then Exit Sub
    If VBA.IsArray(CollectionOrArray) Then Exit Sub
    VBA.Err.Raise 13, Source:="lib_elvin", _
                  Description:="Type mismatch: CollectionOrArray должен быть Collection или Array"
End Sub

'===============================================================================
' # юнит-тесты модуля

Private Sub TestContains()
    Debug.Assert Contains(Array(1, 2, 3), Array(3, 1, 2)) = True
    Debug.Assert Contains(Array(1, 2, 3), Array(3, 1, 4)) = False
    Debug.Print "TestContains is OK"
End Sub
