VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Project1"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''Algae' project 1.01 - speed improvement

'You can use your own bmp (make sure it's 24-bit)
'in one effect by changing the first line in Form_Load

'This program will create a bmp file in same directory if
'none exists.

'To see the effect, uncomment these in the Do While
'in Form_Activate

 'BlitDottedImage FormDib, DotImage, , bErasing
 'MoveDotImage


'there are some other lines you can uncomment in the
'same loop.

  
Private Type Algae
 Center As PrecisionPointAPI
 Spoke As DottedLine
 SpokeCount As Long
 DotsPerSpoke As Long
 radius1 As Single
 radius2 As Single
 CPSelect As ColorProcess
 bErasing As Boolean
 bFilledDots As Boolean
End Type

Dim sinep1!
Dim sinep2!
Dim sinep3!
Dim sinep4!
Dim cosp1!
Dim cosp2!
Dim cosp3!
Dim cosp4!
Dim isinep1!
Dim isinep2!
Dim isinep3!
Dim isinep4!
Dim icosp1!
Dim icosp2!
Dim icosp3!
Dim icosp4!


'some test Algae - commented out in the rendering loop
'in Form_Activate
Dim Wheel1 As Algae
Dim Wheel2 As Algae
Dim Wheel3 As Algae

Private Const LBSW As Long = 1
Private Const UBSW As Long = 100

Dim AlgaeWheel(LBSW To UBSW) As Algae

Dim bFormMouseDown As Boolean

Dim bErasing As Boolean

Dim Elapsed&, LastTic&
Dim FrameCount&
Dim standardspeed!
Dim standardSpeedControl!

'Fps
Dim Tick&
Dim TickSum&
Dim fps!

Dim LRnd&

Enum MultiCP
 ALLSolid
 ALLShift
 ALLInvert
 Mixed
End Enum

Dim MixP1 As MultiCP

Dim bRunning As Boolean

'backbuffer + info
Dim FormDib As AnimSurfaceInfo

'RandomGroup
Dim MaxDotsPerSpoke As Byte
Dim MinSpokeCount1&
Dim SpokeCountVari&
Dim minDotSiz!
Dim sizeVari!
Dim minDefin!
Dim defVarie!

Dim sR!
Dim sG!
Dim sB!
Dim xEnd!
Dim yEnd!

Private Type MoireWheel
 Algae1 As Algae
 Algae2 As Algae
 difAlgae1PointsSpeed As Single
 difAlgae2PointsSpeed As Single
 difAlgaeSpeed As Single
End Type

Dim MWHeel1 As MoireWheel
Dim MWHeel2 As MoireWheel
Dim MWHeel3 As MoireWheel
Dim MWHeel4 As MoireWheel

Dim DotImage As DottedImage

Dim MovingWhichTestPoint&

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Form_Load()

 TruecolorBmpToDottedImage DotImage, "greenfield.bmp"
 
 'Regular vb stuff
 ScaleMode = vbPixels
 Top = 1000
 Left = 1000
 FontSize = 8
 Font = "Verdana"
 FontBold = True
 bRunning = True
 Randomize
 
 'When you press 'R', these variables are read by
 'RandomGroup(), which in turn calls RandomAlgae

 MaxDotsPerSpoke = 6&
 MinSpokeCount1 = 2&
 SpokeCountVari = 6&
 minDotSiz = 1!
 sizeVari = 58!
 minDefin = 0.3!
 defVarie = 18!
  
 bErasing = True

End Sub
Private Sub RandomGroup()

For N = LBSW To UBSW                  '
 RandomAlgae FormDib, AlgaeWheel(N), _
  MaxDotsPerSpoke, _
  MinSpokeCount1, _
  SpokeCountVari, _
  minDotSiz, _
  sizeVari, _
  minDefin, _
  defVarie
Next

'Adjusts brush pressure of each AlgaeWheel(N) based upon
'value of AlgaeWheel(N).CPSelect
Call MixCP(MixP1)

End Sub
Private Sub Form_Activate()

 standardSpeedControl = 2!
 SystemTest
  
 RandomDotImagePos


'Test wheel stuff - uncomment 4 lines
'in the Do While Loop below, to see these
 
 With Wheel1
 .SpokeCount = 3
 .DotsPerSpoke = 25
 .bErasing = True
 .bFilledDots = True
 .radius1 = 20
 .radius2 = 160
 .CPSelect = CSolid
 End With
 ColorPoint Wheel1.Spoke.Point1, 0, 255, 0
 DefIntensDiameter Wheel1.Spoke.Point1, 1.5, 255, 3
 ColorPoint Wheel1.Spoke.Point2, 255, 255, 255
 DefIntensDiameter Wheel1.Spoke.Point2, 1.5, 255, 3
 ProcessDottedLine Wheel1
 
 With Wheel2
 .SpokeCount = 4
 .DotsPerSpoke = 40
 .bErasing = True
 .bFilledDots = True
 .radius1 = 20
 .radius2 = 160
 .CPSelect = CSolid
 End With
 ColorPoint Wheel2.Spoke.Point1, 255, 146, 0
 DefIntensDiameter Wheel2.Spoke.Point1, 1.5, 255, 3
 ColorPoint Wheel2.Spoke.Point2, 255, 255, 0
 DefIntensDiameter Wheel2.Spoke.Point2, 1, 255, 5
 ProcessDottedLine Wheel2

 With Wheel3
 .SpokeCount = 3
 .DotsPerSpoke = 35
 .bErasing = True
 .bFilledDots = True
 .radius1 = 20
 .radius2 = 160
 .CPSelect = CSolid
 End With
 ColorPoint Wheel3.Spoke.Point1, 255, 0, 128
 DefIntensDiameter Wheel3.Spoke.Point1, 1.5, 255, 4.5
 ColorPoint Wheel3.Spoke.Point2, 255, 255, 255
 DefIntensDiameter Wheel3.Spoke.Point2, 1, 255, 3
 ProcessDottedLine Wheel3

'After form has had a chance to resize
Wheel1.Center.px = Rnd * FormDib.Dims.Width
Wheel1.Center.py = Rnd * FormDib.Dims.Height
Wheel2.Center.px = Wheel1.Center.px
Wheel2.Center.py = Wheel1.Center.py
Wheel3.Center.px = Wheel1.Center.px
Wheel3.Center.py = Wheel1.Center.py


'These are initialized after systemtest
Wheel1.Spoke.Point1.iRot = standardspeed / -150&
Wheel1.Spoke.Point2.iRot = standardspeed / -720&

Wheel2.Spoke.Point1.iRot = standardspeed / -350&
Wheel2.Spoke.Point2.iRot = standardspeed / 120&

Wheel3.Spoke.Point1.iRot = standardspeed / 250&
Wheel3.Spoke.Point2.iRot = standardspeed / 220&

RandomGroup

DefaultMoireWheel MWHeel1, 150, 150, 55, 1, 30, 100, 0.3, 0.2, 1, 1, 1.8
DefaultMoireWheel MWHeel2, 450, 150, 15, 0, 10, 120, 0.2, -0.2, -1.91, 1.4, -1.91
DefaultMoireWheel MWHeel3, 240, 350, 55, 0, 20, 120, 0.2, -0.5, 1.1, 1.2, -0.15
DefaultMoireWheel MWHeel4, 500, 350, 65, 1, 50, 90, 0.52, -0.01, 1, 1, -1.25

Do While bRunning
 
 If FormDib.Dims.Width > 0& Then
 
 'BlitDottedImage FormDib, DotImage, , bErasing
 'MoveDotImage
 
 'RandomAlgae FormDib, Wheel3, 12, 2, 12, 1, 18, 0.8, 3
 'DrawAlgae Wheel3
 'DrawAlgae Wheel1
 'DrawAlgae Wheel2
 'DrawAlgae MWHeel1.Algae1
 'DrawAlgae MWHeel1.Algae2
 'DrawAlgae MWHeel2.Algae1
 'DrawAlgae MWHeel2.Algae2
 'DrawAlgae MWHeel3.Algae1
 'DrawAlgae MWHeel3.Algae2
 'DrawAlgae MWHeel4.Algae1
 'DrawAlgae MWHeel4.Algae2
 
 For N = 1 To 13 'I dimmed 100 of these
  DrawAlgae AlgaeWheel(N)
 Next
 
 CalcTick
 CalcFPS
 
 Refresh
 
 CurrentY = 0
 
 ''Useful for debugging
 Print "FPS: " & Round(fps, 1)
 Print "Keys:  R, C, D, S, E,  1 - 4, Space"
 
 EraseSprites FormDib
 
 End If 'FormDib.Dims.Width > 0
 
 DoEvents

Loop

CleanUp

End Sub

Private Sub CalcTick()
 
 Tick = timeGetTime
 Elapsed = Tick - LastTic
 LastTic = Tick

End Sub
Private Sub CalcFPS()
 
 TickSum = TickSum + Elapsed
 If TickSum& > 1000& Then
  fps = 1000& * FrameCount / TickSum
  FrameCount = 0&
  TickSum = 0&
 End If
 
 FrameCount = FrameCount + 1&

End Sub
Private Sub RandomAlgae(Surf As AnimSurfaceInfo, SW1 As Algae, MaxDotsPerSpoke As Byte, MinSpokes&, SpokeCountVariance&, minDotSize!, dotsizeVariance!, defMin!, defVariance!)
Dim bytPress As Byte
Dim bytAlpha As Byte
Dim sDefn!

 With SW1
 .SpokeCount = MinSpokes + Rnd * SpokeCountVariance
 .DotsPerSpoke = Rnd * (MaxDotsPerSpoke - 1) + 1
 .bErasing = bErasing
 If Int(2 * Rnd) = 1 Then
 .bFilledDots = True
 Else
 .bFilledDots = False
 End If
 .radius1 = 50
 .radius2 = .radius1 + Rnd * 150 + 100
 End With
 SW1.Center.px = Rnd * Surf.Dims.Width
 SW1.Center.py = Rnd * Surf.Dims.Height
 
 If SW1.CPSelect = Inverse Then
  bytPress = 255&
 Else
  bytPress = 255& * Rnd '128 + 127& * Rnd
 End If
 sDefn = defMin + Rnd * defVariance '+ 2 - (Rnd) ^ 2
 DefIntensDiameter SW1.Spoke.Point1, sDefn, bytPress, Rnd * dotsizeVariance + minDotSize
 bytAlpha = 255 '* Int(Rnd + 0.5!)
 ColorPoint SW1.Spoke.Point1, bytAlpha, bytAlpha, bytAlpha
 
 If SW1.CPSelect = Inverse Then
  bytPress = 255&
 Else
  bytPress = 255& * Rnd '128 + 127& * Rnd
 End If
 sDefn = defMin + Rnd * defVariance '+ 2 - (Rnd) ^ 2
 DefIntensDiameter SW1.Spoke.Point2, sDefn, bytPress, Rnd * dotsizeVariance + minDotSize
 ColorPoint SW1.Spoke.Point2, Rnd * 255, Rnd * 255, Rnd * 255
 
 ProcessDottedLine SW1
 
 'The rotation speed
 SW1.Spoke.Point1.iRot = standardspeed * Rnd * 0.003
 SW1.Spoke.Point2.iRot = standardspeed * Rnd * 0.01
 
 'initial rotation angle
 SW1.Spoke.Point1.sRot = Rnd * twopi
 SW1.Spoke.Point2.sRot = Rnd * twopi
 
End Sub
Private Sub DefaultMoireWheel(MWHeel As MoireWheel, px!, py!, MinSpoke&, SpokeVar&, rad1!, rad2!, twist1!, twist2!, mltAlg1PtsSpeed!, mltAlg2PtsSpeed!, mltAlg12Speed!)
Dim bytPress As Byte
Dim bytAlpha As Byte
Dim sDefn!, sDia!
Dim ss1!
 
 ss1 = standardspeed / 512&

 'If SW1.CPSelect = Inverse Then
 ' bytPress = 196& + Rnd * 59&
 'Else
  bytPress = 255& '* Rnd
 'End If
 
 sDefn = 1.2  'minDefin + Rnd * defVarie + 2 - (Rnd) ^ 2
 sDia = 3.3
 
 With MWHeel.Algae1
 .SpokeCount = MinSpoke
 .DotsPerSpoke = (rad2 - rad1) / 3!
 .bErasing = bErasing
 'If Int(2 * Rnd) = 1 Then
 .bFilledDots = True
 'Else
 '.bFilledDots = False
 'End If
 .radius1 = rad1
 .radius2 = rad2
 .Center.px = px
 .Center.py = py
 .Spoke.Point1.sRot = 0!
 .Spoke.Point2.sRot = twist1
 .Spoke.Point1.iRot = ss1
 .Spoke.Point2.iRot = ss1 * mltAlg1PtsSpeed
  DefIntensDiameter .Spoke.Point1, sDefn, bytPress, sDia 'Rnd * dotsizeVariance + minDotSize
  DefIntensDiameter .Spoke.Point2, sDefn, bytPress, sDia 'Rnd * dotsizeVariance + minDotSize
 'bytAlpha = 255
 'ColorPoint .Spoke.Point1, bytAlpha, bytAlpha, bytAlpha
 ColorPoint .Spoke.Point1, 0, 0, 0 '255, 255, 255
 ColorPoint .Spoke.Point2, 0, 0, 0 'Rnd * 255
 End With
 ProcessDottedLine MWHeel.Algae1
 
 ss1 = ss1 * mltAlg12Speed
 With MWHeel.Algae2
 .SpokeCount = MinSpoke + SpokeVar
 .DotsPerSpoke = MWHeel.Algae1.DotsPerSpoke '(rad2 - rad1) / 2!
 .bErasing = bErasing
 .bFilledDots = True
 .Center.px = px
 .Center.py = py
 .radius1 = rad1
 .radius2 = rad2
 .Spoke.Point1.sRot = 0!
 .Spoke.Point2.sRot = twist2
 .Spoke.Point1.iRot = ss1
 .Spoke.Point2.iRot = ss1 * mltAlg2PtsSpeed
  'sDefn = 1.3 'minDefin + Rnd * defVarie + 2 - (Rnd) ^ 2
  DefIntensDiameter .Spoke.Point1, sDefn, bytPress, sDia 'Rnd * dotsizeVariance + minDotSize
  DefIntensDiameter .Spoke.Point2, sDefn, bytPress, sDia 'Rnd * dotsizeVariance + minDotSize
 'bytAlpha = 255
 'ColorPoint .Spoke.Point1, bytAlpha, bytAlpha, bytAlpha
 ColorPoint .Spoke.Point1, 0, 0, 0
 ColorPoint .Spoke.Point2, 0, 0, 255
 End With
 ProcessDottedLine MWHeel.Algae2
 
  
End Sub
Private Sub MoveDotImage()
 DotImage.LoLf.px = FormDib.halfW + Cos(cosp1) * 160& 'FormDib.halfW
 DotImage.LoRt.px = FormDib.halfW + Cos(cosp2) * 160& 'FormDib.halfW
 DotImage.HiLf.px = FormDib.halfW + Cos(cosp3) * 160& 'FormDib.halfW
 DotImage.HiRt.px = FormDib.halfW + Cos(cosp4) * 160& 'FormDib.halfW
 DotImage.LoLf.py = FormDib.halfH + Sin(sinep1) * 160& 'FormDib.halfH
 DotImage.LoRt.py = FormDib.halfH + Sin(sinep2) * 160& ' FormDib.halfH
 DotImage.HiLf.py = FormDib.halfH + Sin(sinep3) * 160& 'FormDib.halfH
 DotImage.HiRt.py = FormDib.halfH + Sin(sinep4) * 160& 'FormDib.halfH
 sinep1 = sinep1 + isinep1
 sinep2 = sinep2 + isinep2
 sinep3 = sinep3 + isinep3
 sinep4 = sinep4 + isinep4
 cosp1 = cosp1 + icosp1
 cosp2 = cosp2 + icosp2
 cosp3 = cosp3 + icosp3
 cosp4 = cosp4 + icosp4
End Sub
Private Sub ColorAlgae()
Dim bytAlpha As Byte
For N = LBSW To UBSW                  '
With AlgaeWheel(N)
 bytAlpha = 255 * Int(Rnd + 0.5)
 ColorPoint .Spoke.Point1, bytAlpha, bytAlpha, bytAlpha 'Rnd * 255, Rnd * 255, Rnd * 255
 
 ColorPoint .Spoke.Point2, Rnd * 255, Rnd * 255, Rnd * 255
 
End With
ProcessDottedLine AlgaeWheel(N)
Next
End Sub
Private Sub IntensityAlgae(MinPress As Byte, PressVar As Byte)
For N = LBSW To UBSW                  '
With AlgaeWheel(N)
 If AlgaeWheel(N).CPSelect = Inverse Then
  .Spoke.Point1.intens = 255&
  .Spoke.Point2.intens = 255&
 Else
  .Spoke.Point1.intens = MinPress + Rnd * PressVar
  .Spoke.Point2.intens = MinPress + Rnd * PressVar
 End If
End With
ProcessDottedLine AlgaeWheel(N)
Next
End Sub
Private Sub IntensityAndDefinitionAlgae(MinPressure As Byte, PressureVariance As Byte, minDefinition!, defVariance!)
Dim MinPressure2 As Byte
Dim PressureVar2 As Byte

For N = LBSW To UBSW
With AlgaeWheel(N)
 If .CPSelect = Inverse Then
  MinPressure2 = 255& '176&
  PressureVar2 = 0& 'Rnd * 79&
 .Spoke.Point1.intens = MinPressure2 + Rnd * PressureVar2
 .Spoke.Point2.intens = MinPressure2 + Rnd * PressureVar2
 Else
 .Spoke.Point1.intens = MinPressure + Rnd * PressureVariance
 .Spoke.Point2.intens = MinPressure + Rnd * PressureVariance
 End If
 .Spoke.Point1.defin = minDefinition + Rnd * defVariance
 .Spoke.Point2.defin = minDefinition + Rnd * defVariance
 .bFilledDots = True * Int(Rnd + 0.5!)
 End With
ProcessDottedLine AlgaeWheel(N)
Next

End Sub
Private Sub SizeAlgae(minDotSize!, sizeVariance!)
For N = LBSW To UBSW
With AlgaeWheel(N)
 .Spoke.Point1.diameter = minDotSize + Rnd * sizeVariance
 .Spoke.Point2.diameter = minDotSize + Rnd * sizeVariance
End With
ProcessDottedLine AlgaeWheel(N)
Next
End Sub
Private Sub DrawAlgae(SW1 As Algae)
Dim Loop_Spokes&
Dim Loop_SpokeDots&
Dim sPress!
Dim sDef!
Dim sDia!

 SW1.Spoke.Point1.sAng = SW1.Spoke.Point1.sRot
 SW1.Spoke.Point2.sAng = SW1.Spoke.Point2.sRot
 For Loop_Spokes = 1& To SW1.SpokeCount
  px = SW1.Center.px + SW1.radius1 * Cos(SW1.Spoke.Point1.sAng)
  py = SW1.Center.py + SW1.radius1 * Sin(SW1.Spoke.Point1.sAng)
  xEnd = SW1.Center.px + SW1.radius2 * Cos(SW1.Spoke.Point2.sAng)
  yEnd = SW1.Center.py + SW1.radius2 * Sin(SW1.Spoke.Point2.sAng)
  ix = (xEnd - px) / SW1.DotsPerSpoke
  iy = (yEnd - py) / SW1.DotsPerSpoke
  sR = SW1.Spoke.Point1.sRed
  sG = SW1.Spoke.Point1.sGrn
  sB = SW1.Spoke.Point1.sBlu
  sPress = SW1.Spoke.Point1.intens
  sDef = SW1.Spoke.Point1.defin
  sDia = SW1.Spoke.Point1.diameter
  For Loop_SpokeDots = 1 To SW1.DotsPerSpoke
   AirBrush.chRed = sR
   AirBrush.chGreen = sG
   AirBrush.chBlue = sB
   LumenAirBrush sPress, sDef
   DimensionAirbrush sDia, sDia
   AirBrush.CPMix = SW1.CPSelect
   BlitAirbrush FormDib, px, py, SW1.bErasing, SW1.bFilledDots
   px = px + ix
   py = py + iy
   sR = sR + SW1.Spoke.iRed
   sG = sG + SW1.Spoke.iGrn
   sB = sB + SW1.Spoke.iBlu
   sPress = sPress + SW1.Spoke.iIntens
   sDef = sDef + SW1.Spoke.iDef
   sDia = sDia + SW1.Spoke.iDia
  Next
  SW1.Spoke.Point1.sAng = SW1.Spoke.Point1.sAng + SW1.Spoke.Point1.iAng
  SW1.Spoke.Point2.sAng = SW1.Spoke.Point2.sAng + SW1.Spoke.Point2.iAng
 Next
 SW1.Spoke.Point1.sRot = SW1.Spoke.Point1.sRot + SW1.Spoke.Point1.iRot
 SW1.Spoke.Point2.sRot = SW1.Spoke.Point2.sRot + SW1.Spoke.Point2.iRot
End Sub
Private Sub DefIntensDiameter(ABP As AirBrushPoint, def1!, Intens1 As Byte, diam1!)
 ABP.defin = def1
 ABP.intens = Intens1
 ABP.diameter = diam1
End Sub
Private Sub ColorPoint(ABP As AirBrushPoint, ByVal Red1 As Byte, ByVal Green1 As Byte, ByVal Blue1 As Byte)
 ABP.sRed = Red1
 ABP.sGrn = Green1
 ABP.sBlu = Blue1
End Sub
Private Sub ProcessDottedLine(SPWheel As Algae)
If SPWheel.DotsPerSpoke > 0 Then
 SPWheel.Spoke.iRed = (SPWheel.Spoke.Point2.sRed - SPWheel.Spoke.Point1.sRed) / SPWheel.DotsPerSpoke
 SPWheel.Spoke.iGrn = (SPWheel.Spoke.Point2.sGrn - SPWheel.Spoke.Point1.sGrn) / SPWheel.DotsPerSpoke
 SPWheel.Spoke.iBlu = (SPWheel.Spoke.Point2.sBlu - SPWheel.Spoke.Point1.sBlu) / SPWheel.DotsPerSpoke
 SPWheel.Spoke.iIntens = (SPWheel.Spoke.Point2.intens - SPWheel.Spoke.Point1.intens) / SPWheel.DotsPerSpoke
 SPWheel.Spoke.iDef = (SPWheel.Spoke.Point2.defin - SPWheel.Spoke.Point1.defin) / SPWheel.DotsPerSpoke
 SPWheel.Spoke.iDia = (SPWheel.Spoke.Point2.diameter - SPWheel.Spoke.Point1.diameter) / SPWheel.DotsPerSpoke
 SPWheel.Spoke.Point1.iAng = twopi / SPWheel.SpokeCount
 SPWheel.Spoke.Point2.iAng = twopi / SPWheel.SpokeCount
End If
End Sub

Private Sub Form_Resize()

 'Allocate an array to the Refresh buffer for each object
 AnimPicSurface Form1, FormDib
 
 'Four color gradient
 FCG_ColorTopRight Rnd * 155, Rnd * 155, Rnd * 155
 FCG_ColorLowRight 80, 50, 10
 FCG_ColorTopLeft 255, 255, 255
 FCG_ColorLowLeft 55, 55, 55

 ''Full-size FCG
 WrapFCGToAnimSurf FormDib
 
 Refresh
 
End Sub

Private Sub RandomDotImagePos()
 'DotImage.LoLf.px = Rnd * FormDib.Dims.Width
 ''DotImage.LoRt.px = DotImage.HiRt.px
 'DotImage.LoRt.px = Rnd * FormDib.Dims.Width
 'DotImage.HiLf.px = Rnd * FormDib.Dims.Width
 'DotImage.HiRt.px = Rnd * FormDib.Dims.Width
 'DotImage.LoLf.py = Rnd * FormDib.Dims.Height
 ''DotImage.LoRt.py = DotImage.LoLf.py
 'DotImage.LoRt.py = Rnd * FormDib.Dims.Height
 'DotImage.HiLf.py = Rnd * FormDib.Dims.Height
 'DotImage.HiRt.py = Rnd * FormDib.Dims.Height
 isinep1 = standardspeed * Rnd / 50&
 isinep2 = standardspeed * Rnd / 50&
 isinep3 = standardspeed * Rnd / 50&
 isinep4 = standardspeed * Rnd / 50&
 icosp1 = standardspeed * Rnd / 50&
 icosp2 = standardspeed * Rnd / 50&
 icosp3 = standardspeed * Rnd / 50&
 icosp4 = standardspeed * Rnd / 50&
 sinep1 = twopi * Rnd
 sinep2 = twopi * Rnd
 sinep3 = twopi * Rnd
 sinep4 = twopi * Rnd
 cosp1 = twopi * Rnd
 cosp2 = twopi * Rnd
 cosp3 = twopi * Rnd
 cosp4 = twopi * Rnd
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
 Case vbKeyEscape
  'CleanUp
  bRunning = False
 
 Case vbKeyR
  RandomGroup
  RandomDotImagePos
  FCG_ColorTopRight Rnd * 155, Rnd * 155, Rnd * 155
  WrapFCGToAnimSurf FormDib
 
 Case vbKeyE
  bErasing = Not bErasing
  For N = LBSW To UBSW
  AlgaeWheel(N).bErasing = bErasing
  Next
 
 Case vbKeySpace
  WrapFCGToAnimSurf FormDib
 
 Case vbKeyD '              min, variance
  IntensityAndDefinitionAlgae 0, 255, minDefin, defVarie
  '                                   min, variance
 Case vbKeyC
  ColorAlgae
 Case vbKeyS
  SizeAlgae 0.5, 28
 Case 49 To 52
  MixP1 = KeyCode - 49
  MixCP MixP1
  IntensityAlgae 0, 255
 End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 bRunning = False
End Sub

Private Sub MixCP(MCP1 As MultiCP)
For N = LBSW To UBSW
With AlgaeWheel(N)
 If MCP1 = ALLInvert Then
 .CPSelect = Inverse
 ElseIf MCP1 = ALLShift Then
 .CPSelect = CShift
 ElseIf MCP1 = ALLSolid Then
 .CPSelect = CSolid
 Else
  LRnd = Rnd * 2
  If LRnd <> 0 Then LRnd = Rnd * 2
  If LRnd = 0 Then
  .CPSelect = CSolid
  ElseIf LRnd = 1 Then
  .CPSelect = Inverse
  Else
  .CPSelect = CShift
  End If
 End If
End With
Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ptd1!
Dim ptd2!
Dim ptd3!
Dim ptd4!
Dim minim!

 MovingWhichTestPoint = 0
 
 sR = X - DotImage.LoLf.px
 sG = Y - DotImage.LoLf.py
 ptd1 = sR * sR + sG * sG
 sR = X - DotImage.LoRt.px
 sG = Y - DotImage.LoRt.py
 ptd2 = sR * sR + sG * sG
 sR = X - DotImage.HiLf.px
 sG = Y - DotImage.HiLf.py
 ptd3 = sR * sR + sG * sG
 sR = X - DotImage.HiRt.px
 sG = Y - DotImage.HiRt.py
 ptd4 = sR * sR + sG * sG
 
 minim = 90!
 If ptd1 < minim! Then
  minim = ptd1
  MovingWhichTestPoint = 1&
 End If
 If ptd2 < minim Then
  minim = ptd2
  MovingWhichTestPoint = 2&
 End If
 If ptd3 < minim Then
  minim = ptd3
  MovingWhichTestPoint = 3&
 End If
 If ptd4 < minim Then
  minim = ptd4
  MovingWhichTestPoint = 4&
 End If
 bFormMouseDown = True
 
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bFormMouseDown Then
 'Wheel1.Center.px = X
 'Wheel1.Center.py = Y
 'Wheel2.Center.px = X
 'Wheel2.Center.py = Y
 Select Case MovingWhichTestPoint
 Case 1
 DotImage.LoLf.px = X
 DotImage.LoLf.py = Y
 Case 2
 DotImage.LoRt.px = X
 DotImage.LoRt.py = Y
 Case 3
 DotImage.HiLf.px = X
 DotImage.HiLf.py = Y
 Case 4
 DotImage.HiRt.px = X
 DotImage.HiRt.py = Y
 End Select
 For N = LBSW To UBSW
 'AlgaeWheel(N).Center.px = X
 'AlgaeWheel(N).Center.py = Y
 Next
 'DrawFCGToAnimSurf FormDib, X - 20, Y, 50, 50
End If

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 bFormMouseDown = False
End Sub
Private Sub SystemTest()
Dim NumPasses&

 Elapsed = 0
  
 LastTic = timeGetTime
 
 While Elapsed < 150&
  LoopTestCode
  Elapsed = timeGetTime - LastTic
  NumPasses = NumPasses + 1&
 Wend
 
 standardspeed = (Elapsed / NumPasses) * standardSpeedControl
 
 If standardspeed > 15 Then standardspeed = 15
 
End Sub
Private Sub LoopTestCode()
Dim N102&, N&, sng1!
Dim RA1(999) As Long
 
 For N102 = 1 To 2

  For N = 0 To 999
  RA1(N) = Sqr(RA1(Rnd * 999)) ^ 2
  Next
 
 Next N102
 
End Sub

'This should be called at exit.
'If you declare any more surfaces, put them in here,
'using the format shown
Private Sub CleanUp()
 CopyMemory ByVal VarPtrArray(FormDib.Dib), 0&, 4
 CopyMemory ByVal VarPtrArray(FormDib.LDib), 0&, 4
 End
End Sub

