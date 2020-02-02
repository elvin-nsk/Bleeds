Attribute VB_Name = "Bleeds"
'=======================================================================================
' Макрос           : Bleeds
' Версия           : 2020.02.02
' Автор            : elvin-nsk (me@elvin.nsk.ru)
'=======================================================================================

Option Explicit

Const RELEASE As Boolean = True

'=======================================================================================
' переменные
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

Public cfg As cls_cfg

'=======================================================================================
' публичные процедуры
'=======================================================================================

Sub Start()
  If ActiveSelectionRange Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count > 1 Then
    MsgBox "Выбрано несколько объектов"
    Exit Sub
  End If
  If ActiveSelectionRange.Count < 1 Then
    MsgBox "Выберите объект"
    Exit Sub
  End If
  Set cfg = New cls_cfg
  frm_Main.Show
  Set cfg = Nothing
End Sub

'=======================================================================================
' функции
'=======================================================================================

Function DoBleeds()
  
  If RELEASE Then On Error GoTo ErrHandler
  
  Dim tSrcShape As Shape, tBleeds As Shape, tFinal As Shape
  Dim tRange As New ShapeRange
  Dim tW#, tH#, tName$
  
  lib_elvin.BoostStart "Добавить припуски", RELEASE
  
  Set tSrcShape = ActiveSelectionRange.FirstShape
  
  'если округляем размер
  If cfg.RoundSize Then
    tW = Round(tSrcShape.SizeWidth, cfg.RoundDec)
    tH = Round(tSrcShape.SizeHeight, cfg.RoundDec)
  Else
    tW = tSrcShape.SizeWidth
    tH = tSrcShape.SizeHeight
  End If
  
  'если подчищаем битмап
  If tSrcShape.Type = cdrBitmapShape Then
    If cfg.BitmapTrim Then
      ShrinkBitmap tSrcShape, cfg.BitmapTrimSize
    End If
  End If
  
  tSrcShape.SetSize tW, tH
  
  Set tBleeds = CreateBleeds(tSrcShape, cfg.Bleeds)
  
  'если растрируем всё обратно в битмап
  If tSrcShape.Type = cdrBitmapShape And cfg.BitmapFlatten Then
    tName = tSrcShape.Name
    Set tFinal = Flatten(tSrcShape, tBleeds)
    tFinal.Name = tName
  Else 'а нет - так группируем, обзываем
    tBleeds.Name = "припуски"
    tRange.Add tBleeds
    tRange.Add tSrcShape
    Set tFinal = tRange.Group
    If tSrcShape.Name = "" Then
      tFinal.Name = "группа - растр с припусками"
    Else
      tFinal.Name = tSrcShape.Name & " (группа с припусками)"
    End If
  End If
  
  tFinal.CreateSelection
  
ExitSub:
  lib_elvin.BoostFinish
  Exit Function

ErrHandler:
  MsgBox "Ошибка: " & Err.Description, vbCritical
  Resume ExitSub

End Function

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
  Set BitmapShape = TrimBitmap(BitmapShape, tCrop, False)
  
  'restore
  ActiveDocument.Unit = tSaveUnit
  ActiveDocument.ReferencePoint = tSavePoint

End Sub

Function CreateBleeds(BitmapShape As Shape, ByVal Bleed#) As Shape
  
  Dim tRange As New ShapeRange
  
  On Error Resume Next
  
  tRange.Add createSideBleed(BitmapShape, Bleed, LeftSide)
  tRange.Add createSideBleed(BitmapShape, Bleed, RightSide)
  tRange.Add createSideBleed(BitmapShape, Bleed, TopSide)
  tRange.Add createSideBleed(BitmapShape, Bleed, BottomSide)
  
  tRange.Add createCornerBleed(BitmapShape, Bleed, TopLeftCorner)
  tRange.Add createCornerBleed(BitmapShape, Bleed, TopRightCorner)
  tRange.Add createCornerBleed(BitmapShape, Bleed, BottomLeftCorner)
  tRange.Add createCornerBleed(BitmapShape, Bleed, BottomRightCorner)
  
  On Error GoTo 0
  
  Set CreateBleeds = tRange.Group

End Function

Function Flatten(SourceBitmap As Shape, BleedsGroup As Shape) As Shape
  Dim tRange As New ShapeRange
  Dim tW#, tH#
  tRange.Add SourceBitmap
  tRange.Add BleedsGroup
  tW = tRange.SizeWidth
  tH = tRange.SizeHeight
  tRange.SetPixelAlignedRendering True
  With SourceBitmap.Bitmap
  If .ResolutionX <> .ResolutionY Then
    tRange.SizeHeight = tRange.SizeHeight * .ResolutionY / .ResolutionX
  End If
    Set Flatten = tRange.ConvertToBitmapEx(.Mode, , .Transparent, .ResolutionX, cdrNoAntiAliasing, False)
  End With
  tRange.SetSize tW, tH
End Function

'=======================================================================================
' вторичные функции
'=======================================================================================

Function createSideBleed(BitmapShape As Shape, ByVal Bleed#, ByVal Side As enum_Sides) As Shape
  
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
  
  Set createSideBleed = lib_elvin.CropTool(BitmapShape.Duplicate, BitmapShape.LeftX + tLeftAdd, _
                                                                  BitmapShape.TopY + tTopAdd, _
                                                                  BitmapShape.RightX + tRightAdd, _
                                                                  BitmapShape.BottomY + tBottomAdd).FirstShape
  If createSideBleed Is Nothing Then Exit Function
  createSideBleed.Flip tFlip
  createSideBleed.Move tShiftX, tShiftY
  createSideBleed.Name = "боковой припуск"

End Function

Function createCornerBleed(BitmapShape As Shape, ByVal Bleed#, ByVal Corner As enum_Corners) As Shape
  
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
  Set createCornerBleed = lib_elvin.CropTool(BitmapShape.Duplicate, BitmapShape.LeftX + tLeftAdd, _
                                                                    BitmapShape.TopY + tTopAdd, _
                                                                    BitmapShape.RightX + tRightAdd, _
                                                                    BitmapShape.BottomY + tBottomAdd).FirstShape
  If createCornerBleed Is Nothing Then Exit Function
  createCornerBleed.Flip cdrFlipBoth
  createCornerBleed.Move tShiftX, tShiftY
  createCornerBleed.Name = "угловой припуск"
  
End Function
