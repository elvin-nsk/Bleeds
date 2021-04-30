Attribute VB_Name = "Bleeds"
'===============================================================================
' Макрос           : Bleeds
' Версия           : 2021.04.30
' Автор            : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = True

'===============================================================================

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

'===============================================================================

Sub Start()

  If RELEASE Then On Error GoTo Catch
  
  If ActiveSelectionRange Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count > 1 Then
    MsgBox "Выбрано несколько объектов"
    Exit Sub
  End If
  If ActiveSelectionRange.Count < 1 Then
    MsgBox "Выберите объект"
    Exit Sub
  End If
  
  Dim Cfg As Config
  Set Cfg = New Config
  Dim View As MainView
  Set View = New MainView
  
  With View
    
    If ActiveSelectionRange.FirstShape.Type = cdrBitmapShape Then .IsRastr = True
    .BleedsMin = 0.1
    .BleedsMax = 10000
    
    Cfg.Load
    .tbBleeds.Value = CStr(Cfg.Bleeds)
    .cbRound.Value = Cfg.RoundSize
    Select Case Cfg.RoundDec
      Case 0
        .obRound0.Value = True
      Case 1
        .obRound1.Value = True
      Case 2
        .obRound2.Value = True
    End Select
    .cbTrim.Value = Cfg.BitmapTrim
    .tbTrim.Value = Cfg.BitmapTrimSize
    .cbFlatten.Value = Cfg.BitmapFlatten
    
    .Show
    If Not .IsOK Then Exit Sub
    
    Cfg.Bleeds = CDbl(.tbBleeds.Value)
    Cfg.RoundSize = .cbRound.Value
    Select Case True
      Case .obRound0.Value = True
        Cfg.RoundDec = 0
      Case .obRound1.Value = True
        Cfg.RoundDec = 1
      Case .obRound2.Value = True
        Cfg.RoundDec = 2
    End Select
    Cfg.BitmapTrim = .cbTrim.Value
    Cfg.BitmapTrimSize = .tbTrim.Value
    Cfg.BitmapFlatten = .cbFlatten.Value
    Cfg.Save
  
  End With
  
  lib_elvin.BoostStart "Добавить припуски", RELEASE
    
  SetBleeds Cfg
  
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  MsgBox "Ошибка: " & Err.Description, vbCritical
  Resume Finally
End Sub

'===============================================================================

Function SetBleeds(Cfg As Config)
  
  Dim SrcShape As Shape, Bleeds As Shape, Final As Shape
  Dim Range As New ShapeRange
  Dim W#, H#, Name$
  
  Set SrcShape = ActiveSelectionRange.FirstShape
  
  'если округляем размер
  If Cfg.RoundSize Then
    W = Round(SrcShape.SizeWidth, Cfg.RoundDec)
    H = Round(SrcShape.SizeHeight, Cfg.RoundDec)
  Else
    W = SrcShape.SizeWidth
    H = SrcShape.SizeHeight
  End If
  
  'если подчищаем битмап
  If SrcShape.Type = cdrBitmapShape Then
    If Cfg.BitmapTrim Then
      ShrinkBitmap SrcShape, Cfg.BitmapTrimSize
    End If
  End If
  
  SrcShape.SetSize W, H
  
  Set Bleeds = CreateBleeds(SrcShape, Cfg.Bleeds)
  
  'если растрируем всё обратно в битмап
  If SrcShape.Type = cdrBitmapShape And Cfg.BitmapFlatten Then
    Name = SrcShape.Name
    Set Final = Flatten(SrcShape, Bleeds)
    Final.Name = Name
  Else 'а нет - так группируем, обзываем
    Bleeds.Name = "припуски"
    Range.Add Bleeds
    Range.Add SrcShape
    Set Final = Range.Group
    If SrcShape.Name = "" Then
      Final.Name = "группа - растр с припусками"
    Else
      Final.Name = SrcShape.Name & " (группа с припусками)"
    End If
  End If
  
  Final.CreateSelection

End Function

Sub ShrinkBitmap(BitmapShape As Shape, ByVal Pixels&)
  
  Dim Crop As Shape
  Dim PxW#, PxH#, SizeW#, SizeH#, AngleMult&
  Dim SaveUnit As cdrUnit, SavePoint As cdrReferencePoint
  
  If BitmapShape.Type <> cdrBitmapShape Then Exit Sub
  If Pixels < 1 Then Exit Sub
  
  'save
  SaveUnit = ActiveDocument.Unit
  SavePoint = ActiveDocument.ReferencePoint
  
  ActiveDocument.Unit = cdrInch
  ActiveDocument.ReferencePoint = cdrCenter
  With BitmapShape
    SizeW = .SizeWidth
    SizeH = .SizeHeight
    AngleMult = .RotationAngle \ 90
    .ClearTransformations
    .RotationAngle = AngleMult * 90
    .SetSize SizeW, SizeH
    PxW = 1 / .Bitmap.ResolutionX
    PxH = 1 / .Bitmap.ResolutionY
    Set Crop = .Layer.CreateRectangle(BitmapShape.LeftX + PxW * Pixels, _
                                                  .TopY - PxH * Pixels, _
                                                  .RightX - PxW * Pixels, _
                                                  .BottomY + PxH * Pixels)
  End With
  Set BitmapShape = TrimBitmap(BitmapShape, Crop, False)
  
  'restore
  ActiveDocument.Unit = SaveUnit
  ActiveDocument.ReferencePoint = SavePoint

End Sub

Function CreateBleeds(BitmapShape As Shape, ByVal Bleed#) As Shape
  
  Dim Range As New ShapeRange
  
  On Error Resume Next
  
  Range.Add CreateSideBleed(BitmapShape, Bleed, LeftSide)
  Range.Add CreateSideBleed(BitmapShape, Bleed, RightSide)
  Range.Add CreateSideBleed(BitmapShape, Bleed, TopSide)
  Range.Add CreateSideBleed(BitmapShape, Bleed, BottomSide)
  
  Range.Add CreateCornerBleed(BitmapShape, Bleed, TopLeftCorner)
  Range.Add CreateCornerBleed(BitmapShape, Bleed, TopRightCorner)
  Range.Add CreateCornerBleed(BitmapShape, Bleed, BottomLeftCorner)
  Range.Add CreateCornerBleed(BitmapShape, Bleed, BottomRightCorner)
  
  On Error GoTo 0
  
  Set CreateBleeds = Range.Group

End Function

Function Flatten(SourceBitmap As Shape, BleedsGroup As Shape) As Shape
  Dim Range As New ShapeRange
  Dim W#, H#
  Range.Add SourceBitmap
  Range.Add BleedsGroup
  W = Range.SizeWidth
  H = Range.SizeHeight
  Range.SetPixelAlignedRendering True
  With SourceBitmap.Bitmap
    If .ResolutionX <> .ResolutionY Then
      Range.SizeHeight = Range.SizeHeight * .ResolutionY / .ResolutionX
    End If
    Set Flatten = Range.ConvertToBitmapEx(.Mode, , .Transparent, .ResolutionX, cdrNoAntiAliasing, False)
  End With
  Flatten.SetSize W, H
End Function

Function CreateSideBleed(BitmapShape As Shape, ByVal Bleed#, ByVal Side As enum_Sides) As Shape
  
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
  
  Set CreateSideBleed = lib_elvin.CropTool(BitmapShape.Duplicate, BitmapShape.LeftX + tLeftAdd, _
                                                                  BitmapShape.TopY + tTopAdd, _
                                                                  BitmapShape.RightX + tRightAdd, _
                                                                  BitmapShape.BottomY + tBottomAdd).FirstShape
  If CreateSideBleed Is Nothing Then Exit Function
  CreateSideBleed.Flip tFlip
  CreateSideBleed.Move tShiftX, tShiftY
  CreateSideBleed.Name = "боковой припуск"

End Function

Function CreateCornerBleed(BitmapShape As Shape, ByVal Bleed#, ByVal Corner As enum_Corners) As Shape
  
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
  Set CreateCornerBleed = lib_elvin.CropTool(BitmapShape.Duplicate, BitmapShape.LeftX + tLeftAdd, _
                                                                    BitmapShape.TopY + tTopAdd, _
                                                                    BitmapShape.RightX + tRightAdd, _
                                                                    BitmapShape.BottomY + tBottomAdd).FirstShape
  If CreateCornerBleed Is Nothing Then Exit Function
  CreateCornerBleed.Flip cdrFlipBoth
  CreateCornerBleed.Move tShiftX, tShiftY
  CreateCornerBleed.Name = "угловой припуск"
  
End Function
