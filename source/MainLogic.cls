VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

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

Private Type typeThis
    Cfg As Config
End Type
Private This As typeThis

'===============================================================================
' # Constructors

Public Function Create( _
                    ByVal View As MainView _
                ) As MainLogic
    Set Create = New MainLogic
    Dim Cfg As Config
    Set Cfg = Config.Bind
    
    With View
        
        .BleedsMin = 0.1
        .BleedsMax = 10000
        .tbBleeds.Value = VBA.CStr(Cfg.Bleeds)
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
    
    End With
    
    Create.Inject Cfg
End Function

Friend Sub Inject( _
               ByVal Cfg As Config _
           )
    With This
        Set .Cfg = Cfg
    End With
End Sub

'===============================================================================
' # Logic

Public Function Execute(ByVal View As MainView)

    If Bleeds.RELEASE Then On Error GoTo Catch
    
    FillCfgFromView View
    
    lib_elvin.BoostStart "Добавить припуски", Bleeds.RELEASE
        
    SetBleeds This.Cfg
    
Finally:
    lib_elvin.BoostFinish
    Exit Function

Catch:
    MsgBox "Ошибка: " & Err.Description, vbCritical
    Resume Finally
End Function

'===============================================================================
' # Public

Public Function Dispose(ByVal View As MainView)
    FillCfgFromView View
End Function

Public Property Get Self() As MainLogic
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Sub FillCfgFromView(ByVal View As MainView)
    With View
        
        This.Cfg.Bleeds = VBA.CDbl(.tbBleeds.Value)
        This.Cfg.RoundSize = .cbRound.Value
        Select Case True
            Case .obRound0.Value = True
                This.Cfg.RoundDec = 0
            Case .obRound1.Value = True
                This.Cfg.RoundDec = 1
            Case .obRound2.Value = True
                This.Cfg.RoundDec = 2
        End Select
        This.Cfg.BitmapTrim = .cbTrim.Value
        This.Cfg.BitmapTrimSize = .tbTrim.Value
        This.Cfg.BitmapFlatten = .cbFlatten.Value
    
    End With
End Sub

Private Function SetBleeds(ByVal Cfg As Config)
    
    Dim SrcShape As Shape, Bleeds As Shape, Final As Shape
    Dim Range As New ShapeRange
    Dim Width As Double, Height As Double, Name As String
    
    Set SrcShape = ActiveSelectionRange.FirstShape
    
    'если округляем размер
    If Cfg.RoundSize Then
        Width = Round(SrcShape.SizeWidth, Cfg.RoundDec)
        Height = Round(SrcShape.SizeHeight, Cfg.RoundDec)
    Else
        Width = SrcShape.SizeWidth
        Height = SrcShape.SizeHeight
    End If
    
    'если подчищаем битмап
    If SrcShape.Type = cdrBitmapShape Then
        If Cfg.BitmapTrim Then
            ShrinkBitmap SrcShape, Cfg.BitmapTrimSize
        End If
    End If
    
    SrcShape.SetSize Width, Height
    
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

Private Sub ShrinkBitmap( _
                ByRef ioBitmapShape As Shape, _
                ByVal Pixels As Long _
            )
    
    Dim Crop As Shape
    Dim PxW As Double, PxH As Double
    Dim SizeW As Double, SizeH As Double, AngleMult As Long
    Dim SaveUnit As cdrUnit, SavePoint As cdrReferencePoint
    
    If ioBitmapShape.Type <> cdrBitmapShape Then Exit Sub
    If Pixels < 1 Then Exit Sub
    
    'save
    SaveUnit = ActiveDocument.Unit
    SavePoint = ActiveDocument.ReferencePoint
    
    ActiveDocument.Unit = cdrInch
    ActiveDocument.ReferencePoint = cdrCenter
    With ioBitmapShape
        SizeW = .SizeWidth
        SizeH = .SizeHeight
        AngleMult = .RotationAngle \ 90
        .ClearTransformations
        .RotationAngle = AngleMult * 90
        .SetSize SizeW, SizeH
        PxW = 1 / .Bitmap.ResolutionX
        PxH = 1 / .Bitmap.ResolutionY
        Set Crop = _
            .Layer.CreateRectangle( _
                ioBitmapShape.LeftX + PxW * Pixels, _
                .TopY - PxH * Pixels, _
                .RightX - PxW * Pixels, _
                .BottomY + PxH * Pixels _
            )
    End With
    Set ioBitmapShape = TrimBitmap(ioBitmapShape, Crop, False)
    
    'restore
    ActiveDocument.Unit = SaveUnit
    ActiveDocument.ReferencePoint = SavePoint

End Sub

Private Function CreateBleeds( _
                     ByVal BitmapShape As Shape, _
                     ByVal Bleed As Double _
                 ) As Shape
    
    Dim Shapes As New ShapeRange
    
    On Error Resume Next
    
    With Shapes
        .Add CreateSideBleed(BitmapShape, Bleed, LeftSide)
        .Add CreateSideBleed(BitmapShape, Bleed, RightSide)
        .Add CreateSideBleed(BitmapShape, Bleed, TopSide)
        .Add CreateSideBleed(BitmapShape, Bleed, BottomSide)
        
        .Add CreateCornerBleed(BitmapShape, Bleed, TopLeftCorner)
        .Add CreateCornerBleed(BitmapShape, Bleed, TopRightCorner)
        .Add CreateCornerBleed(BitmapShape, Bleed, BottomLeftCorner)
        .Add CreateCornerBleed(BitmapShape, Bleed, BottomRightCorner)
    End With
    
    On Error GoTo 0
    
    Set CreateBleeds = Shapes.Group

End Function

Private Function Flatten( _
                     ByVal SourceBitmap As Shape, _
                     ByVal BleedsGroup As Shape _
                 ) As Shape
    Dim Range As New ShapeRange
    Dim Width#, Height#
    Range.Add SourceBitmap
    Range.Add BleedsGroup
    Width = Range.SizeWidth
    Height = Range.SizeHeight
    Range.SetPixelAlignedRendering True
    With SourceBitmap.Bitmap
        If .ResolutionX <> .ResolutionY Then
            Range.SizeHeight = _
                Range.SizeHeight * .ResolutionY / .ResolutionX
        End If
        Set Flatten = _
            Range.ConvertToBitmapEx( _
                .Mode, , .Transparent, .ResolutionX, _
                cdrNoAntiAliasing, False _
            )
    End With
    Flatten.SetSize Width, Height
End Function

Private Function CreateSideBleed( _
                     ByVal BitmapShape As Shape, _
                     ByVal Bleed As Double, _
                     ByVal Side As enum_Sides _
                 ) As Shape
    
    Dim LeftAdd As Double, RightAdd As Double
    Dim TopAdd As Double, BottomAdd As Double
    Dim ShiftX As Double, ShiftY As Double
    Dim Flip As cdrFlipAxes
    
    Select Case Side
        Case LeftSide
            RightAdd = -(BitmapShape.SizeWidth - Bleed)
            Flip = cdrFlipHorizontal
            ShiftX = -Bleed
        Case RightSide
            LeftAdd = BitmapShape.SizeWidth - Bleed
            Flip = cdrFlipHorizontal
            ShiftX = Bleed
        Case TopSide
            BottomAdd = BitmapShape.SizeHeight - Bleed
            Flip = cdrFlipVertical
            ShiftY = Bleed
        Case BottomSide
            TopAdd = -(BitmapShape.SizeHeight - Bleed)
            Flip = cdrFlipVertical
            ShiftY = -Bleed
    End Select
    
    Set CreateSideBleed = _
        lib_elvin.CropTool( _
            BitmapShape.Duplicate, BitmapShape.LeftX + LeftAdd, _
            BitmapShape.TopY + TopAdd, _
            BitmapShape.RightX + RightAdd, _
            BitmapShape.BottomY + BottomAdd _
        ).FirstShape
    If CreateSideBleed Is Nothing Then Exit Function
    CreateSideBleed.Flip Flip
    CreateSideBleed.Move ShiftX, ShiftY
    CreateSideBleed.Name = "боковой припуск"

End Function

Private Function CreateCornerBleed( _
                     ByVal BitmapShape As Shape, _
                     ByVal Bleed As Double, _
                     ByVal Corner As enum_Corners _
                 ) As Shape
    
    Dim LeftAdd As Double, RightAdd As Double
    Dim TopAdd As Double, BottomAdd As Double
    Dim ShiftX As Double, ShiftY As Double
    
    Select Case Corner
        Case TopLeftCorner
            RightAdd = -(BitmapShape.SizeWidth - Bleed)
            BottomAdd = BitmapShape.SizeHeight - Bleed
            ShiftX = -Bleed
            ShiftY = Bleed
        Case TopRightCorner
            LeftAdd = BitmapShape.SizeWidth - Bleed
            BottomAdd = BitmapShape.SizeHeight - Bleed
            ShiftX = Bleed
            ShiftY = Bleed
        Case BottomLeftCorner
            RightAdd = -(BitmapShape.SizeWidth - Bleed)
            TopAdd = -(BitmapShape.SizeHeight - Bleed)
            ShiftX = -Bleed
            ShiftY = -Bleed
        Case BottomRightCorner
            LeftAdd = BitmapShape.SizeWidth - Bleed
            TopAdd = -(BitmapShape.SizeHeight - Bleed)
            ShiftX = Bleed
            ShiftY = -Bleed
    End Select
    Set CreateCornerBleed = _
        lib_elvin.CropTool( _
            BitmapShape.Duplicate, BitmapShape.LeftX + LeftAdd, _
            BitmapShape.TopY + TopAdd, _
            BitmapShape.RightX + RightAdd, _
            BitmapShape.BottomY + BottomAdd _
        ).FirstShape
    If CreateCornerBleed Is Nothing Then Exit Function
    CreateCornerBleed.Flip cdrFlipBoth
    CreateCornerBleed.Move ShiftX, ShiftY
    CreateCornerBleed.Name = "угловой припуск"
    
End Function
