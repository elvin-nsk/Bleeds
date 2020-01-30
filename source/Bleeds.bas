Attribute VB_Name = "Bleeds"
'=======================================================================================
' Макрос           : Bleeds
' Версия           : 2020.01.29
' Автор            : elvin-nsk (me@elvin.nsk.ru)
'=======================================================================================

Option Explicit

Const RELEASE As Boolean = True

'=======================================================================================
' глобальные переменные
'=======================================================================================

Enum enum_Sides
  LeftSide = 0
  RightSide = 1
  TopSide = 3
  BottomSide = 4
End Enum

Enum enum_Corners
  TopLeftCorner = 0
  TopRightCorner = 1
  BottomLeftCorner = 2
  BottomRightCorner = 3
End Enum

'=======================================================================================
' публичные процедуры
'=======================================================================================

Sub Start()
  
  'припуск
  Const BLEEDSIZE# = 3 'мм
  'на сколько подрезать битмап
  Const CLEARSIDES& = 2 'пиксели
  'до какого знака округлять обрезной размер (0 - до целых мм)
  Const CLEARROUNDDECPLACES& = 0
  
  If RELEASE Then On Error GoTo ErrHandler
  
  Dim tBitmapShape As Shape, tBleeds As Shape, tFinal As Shape
  Dim tRange As New ShapeRange
  Dim BitmapW#, BitmapH#
  
  BoostStart "Поставить припуски", RELEASE
  
  If Not ActiveSelectionRange.FirstShape Is Nothing Then
    Set tBitmapShape = ActiveSelectionRange.FirstShape
  Else
    MsgBox "Не выбран объект"
    Exit Sub
  End If
  If tBitmapShape.Type <> cdrBitmapShape Then
    MsgBox "Выбранный объект не является растром"
    Exit Sub
  End If
  
  BitmapW = Round(tBitmapShape.SizeWidth, CLEARROUNDDECPLACES)
  BitmapH = Round(tBitmapShape.SizeHeight, CLEARROUNDDECPLACES)
  
  ShrinkBitmap tBitmapShape, CLEARSIDES
  
  tBitmapShape.SetSize BitmapW, BitmapH
  
  Set tBleeds = CreateBleeds(tBitmapShape, BLEEDSIZE)
  tBleeds.Name = "припуски"
  tRange.Add tBleeds
  tRange.Add tBitmapShape
  Set tFinal = tRange.Group
  If tBitmapShape.Name = "" Then
    tFinal.Name = "группа - растр с припусками"
  Else
    tFinal.Name = tBitmapShape.Name & " (группа с припусками)"
  End If
  tFinal.CreateSelection
  
ExitSub:
  BoostFinish
  Exit Sub

ErrHandler:
  MsgBox "Ошибка: " & Err.Description, vbCritical
  Resume ExitSub

End Sub

'=======================================================================================
' функции
'=======================================================================================

Sub ShrinkBitmap(BitmapShape As Shape, ByVal Pixels&)
  
  Dim tCrop As Shape
  Dim tPxW#, tPxH#, tSizeW#, tSizeH#, tAngleMult&
  Dim tSaveUnit As cdrUnit, tSavePoint As cdrReferencePoint
  
  If BitmapShape.Type <> cdrBitmapShape Then Exit Sub
  If Pixels < 1 Then Exit Sub
  
  'save
  tSaveUnit = ActiveDocument.Unit
  tSavePoint = ActiveDocument.ReferencePoint
  
  ActiveDocument.Unit = cdrInch
  ActiveDocument.ReferencePoint = cdrCenter
  With BitmapShape
    tSizeW = .SizeWidth
    tSizeH = .SizeHeight
    tAngleMult = .RotationAngle \ 90
    .ClearTransformations
    .RotationAngle = tAngleMult * 90
    .SetSize tSizeW, tSizeH
    tPxW = 1 / .Bitmap.ResolutionX
    tPxH = 1 / .Bitmap.ResolutionY
    Set tCrop = .Layer.CreateRectangle(BitmapShape.LeftX + tPxW * Pixels, _
                                                  .TopY - tPxH * Pixels, _
                                                  .RightX - tPxW * Pixels, _
                                                  .BottomY + tPxH * Pixels)
  End With
  Set BitmapShape = trimBitmap(BitmapShape, tCrop, False)
  
  'restore
  ActiveDocument.Unit = tSaveUnit
  ActiveDocument.ReferencePoint = tSavePoint

End Sub

Function CreateBleeds(BitmapShape As Shape, ByVal Bleed#) As Shape
  
  Dim tRange As New ShapeRange
  
  If BitmapShape.Type <> cdrBitmapShape Then Exit Function
  
  tRange.Add createSideBleed(BitmapShape, Bleed, LeftSide)
  tRange.Add createSideBleed(BitmapShape, Bleed, RightSide)
  tRange.Add createSideBleed(BitmapShape, Bleed, TopSide)
  tRange.Add createSideBleed(BitmapShape, Bleed, BottomSide)
  
  tRange.Add createCornerBleed(BitmapShape, Bleed, TopLeftCorner)
  tRange.Add createCornerBleed(BitmapShape, Bleed, TopRightCorner)
  tRange.Add createCornerBleed(BitmapShape, Bleed, BottomLeftCorner)
  tRange.Add createCornerBleed(BitmapShape, Bleed, BottomRightCorner)
  
  Set CreateBleeds = tRange.Group

End Function

'=======================================================================================
' вторичные функции
'=======================================================================================

Function createSideBleed(BitmapShape As Shape, ByVal Bleed#, ByVal Side As enum_Sides) As Shape

  Dim tCrop As Shape
  Dim tLeftAdd#, tRightAdd#, tTopAdd#, tBottomAdd#
  Dim tShiftX#, tShiftY#
  Dim tFlip As cdrFlipAxes
  
  Select Case Side
    Case LeftSide
      tRightAdd = -(BitmapShape.SizeWidth - Bleed)
      tFlip = cdrFlipHorizontal
      tShiftX = -Bleed
    Case RightSide
      tLeftAdd = BitmapShape.SizeWidth - Bleed
      tFlip = cdrFlipHorizontal
      tShiftX = Bleed
    Case TopSide
      tBottomAdd = BitmapShape.SizeHeight - Bleed
      tFlip = cdrFlipVertical
      tShiftY = Bleed
    Case BottomSide
      tTopAdd = -(BitmapShape.SizeHeight - Bleed)
      tFlip = cdrFlipVertical
      tShiftY = -Bleed
  End Select

  Set tCrop = BitmapShape.Layer.CreateRectangle(BitmapShape.LeftX + tLeftAdd, _
                                                BitmapShape.TopY + tTopAdd, _
                                                BitmapShape.RightX + tRightAdd, _
                                                BitmapShape.BottomY + tBottomAdd)
  Set createSideBleed = trimBitmap(BitmapShape.Duplicate, tCrop, False)
  createSideBleed.Flip tFlip
  createSideBleed.Move tShiftX, tShiftY
  createSideBleed.Name = "боковой припуск"

End Function

Function createCornerBleed(BitmapShape As Shape, ByVal Bleed#, ByVal Corner As enum_Corners) As Shape

  Dim tCrop As Shape
  Dim tLeftAdd#, tRightAdd#, tTopAdd#, tBottomAdd#
  Dim tShiftX#, tShiftY#
  
  Select Case Corner
    Case TopLeftCorner
      tRightAdd = -(BitmapShape.SizeWidth - Bleed)
      tBottomAdd = BitmapShape.SizeHeight - Bleed
      tShiftX = -Bleed
      tShiftY = Bleed
    Case TopRightCorner
      tLeftAdd = BitmapShape.SizeWidth - Bleed
      tBottomAdd = BitmapShape.SizeHeight - Bleed
      tShiftX = Bleed
      tShiftY = Bleed
    Case BottomLeftCorner
      tRightAdd = -(BitmapShape.SizeWidth - Bleed)
      tTopAdd = -(BitmapShape.SizeHeight - Bleed)
      tShiftX = -Bleed
      tShiftY = -Bleed
    Case BottomRightCorner
      tLeftAdd = BitmapShape.SizeWidth - Bleed
      tTopAdd = -(BitmapShape.SizeHeight - Bleed)
      tShiftX = Bleed
      tShiftY = -Bleed
  End Select

  Set tCrop = BitmapShape.Layer.CreateRectangle(BitmapShape.LeftX + tLeftAdd, _
                                                BitmapShape.TopY + tTopAdd, _
                                                BitmapShape.RightX + tRightAdd, _
                                                BitmapShape.BottomY + tBottomAdd)
  Set createCornerBleed = trimBitmap(BitmapShape.Duplicate, tCrop, False)
  createCornerBleed.Flip cdrFlipBoth
  createCornerBleed.Move tShiftX, tShiftY
  createCornerBleed.Name = "угловой припуск"
  
End Function

Function trimBitmap(BitmapShape As Shape, CropEnvelopeShape As Shape, Optional ByVal LeaveCropEnvelope As Boolean = True) As Shape

  Const EXPANDBY& = 2 'px
  
  Dim tCrop As Shape
  Dim tPxW#, tPxH#
  Dim tSaveUnit As cdrUnit

  If BitmapShape.Type <> cdrBitmapShape Then Exit Function
  
  'save
  tSaveUnit = ActiveDocument.Unit
  
  ActiveDocument.Unit = cdrInch
  tPxW = 1 / BitmapShape.Bitmap.ResolutionX
  tPxH = 1 / BitmapShape.Bitmap.ResolutionY
  BitmapShape.Bitmap.ResetCropEnvelope
  Set tCrop = BitmapShape.Layer.CreateRectangle(CropEnvelopeShape.LeftX - tPxW * EXPANDBY, _
                                                CropEnvelopeShape.TopY + tPxH * EXPANDBY, _
                                                CropEnvelopeShape.RightX + tPxW * EXPANDBY, _
                                                CropEnvelopeShape.BottomY - tPxH * EXPANDBY)
  Set trimBitmap = tCrop.Intersect(BitmapShape, False, False)
  If trimBitmap Is Nothing Then
    tCrop.Delete
    GoTo ExitFunction
  End If
  trimBitmap.Bitmap.Crop
  Set trimBitmap = CropEnvelopeShape.Intersect(trimBitmap, LeaveCropEnvelope, False)
  
ExitFunction:
  'restore
  ActiveDocument.Unit = tSaveUnit
  
End Function

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
