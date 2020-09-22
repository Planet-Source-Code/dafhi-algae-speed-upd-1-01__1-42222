Attribute VB_Name = "modAirbrush"
Option Explicit

Public Type AirBrushPoint
 sRed As Single
 sGrn As Single
 sBlu As Single
 intens As Single
 defin As Single
 diameter As Single
 iAng As Single
 sAng As Single
 iRot As Single
 sRot As Single
End Type

Private Type AirBrushGrid
 Wide As Long
 High As Long
 HighM1 As Long 'Minus 1
 RightM1 As Long 'Minus 1
 AddHalfW As Long
 AddHalfH As Long
 InvertAndOffSet As Long
End Type

Private Type AirBrushSystemVariables
 SpriteGrid As AirBrushGrid
 intens As Single
 int_x_def As Single
 ix As Single
 iy As Single
 wall_x As Single
 floor_y As Single
 prev_wide As Single
 prev_high As Single
 DrawBot As Long
 DrawTop As Long
 DrawRight As Long
 DrawLeft As Long
 ClipLeft As Long
 ClipTop As Long
 ClipBot As Long
 blit_x As Single
 blit_y As Single
 px As Single
 py As Single
 pz As Single
End Type

Private Type AirBrushPrecisionDims
 wWide As Single
 hHigh As Single
 wDiv2 As Single
 hDiv2 As Single
End Type

Public Type AirbrushStruct
 chRed As Byte 'ch refers to '8 bits'
 chGreen As Byte
 chBlue As Byte
 chIntensity As Byte
 CPMix As ColorProcess
 I2 As Long
 definition As Single
 Dims As AirBrushPrecisionDims
 IProcess As AirBrushSystemVariables
 dx As Single
 dy As Single
 dz As Single
End Type

Private Type PrecisionPointAPI
 px As Single
 py As Single
End Type

Enum ColorProcess
 CSolid
 CShift
 Inverse
End Enum

Public Type DottedLine
 Point1 As AirBrushPoint
 Point2 As AirBrushPoint
 iRed As Single
 iGrn As Single
 iBlu As Single
 iIntens As Single
 iDef As Single
 iDia As Single
End Type

'==================================
' with subs that are called frequently, it's computationally
' faster to dim variables outside.  Some of these variables
' are discussed in the subs that use them.

' AirbrushClipper
Dim Left_&
Dim Top_&
Dim Right_&
Dim Bot_&

' BlitAirbrush
Dim sngMaskAlphaTest!
Dim intensity!
Dim sR!
Dim sG!
Dim sB!
Dim offset_x!
Dim offset_y!
Dim delta_y!
Dim delta_ySq!
Dim delta_x!
Dim baseleft!
Dim Bright&

' SysDimensionAirbrush
'Dim half!
'Dim Rounded&
'Dim pp5!
'Dim mult2!
'Dim add2!
'Dim left_side_scale!
'Dim right_side_scale!
'Dim top_edge_scale!
'Dim bottom_edge_scale!

Dim DrawWidth_&
Dim Loca&
Dim Loca1&
Dim Loca2&

Dim AddDrawWidth&
Dim AddDrawWidthBytes&

Dim DrawTop&
Dim DrawBot&
Dim DrawRight&

'ColorShift
Dim iR!
Dim iG!
Dim iB!

Dim maximu!
Dim minimu!
Dim iSubt!
Dim bytMaxMin_diff As Byte

Type BGRAQUAD
 Blue As Byte
 Green As Byte
 Red As Byte
 Alpha As Byte
End Type

Type SAFEARRAY1D
 cDims As Integer
 fFeatures As Integer
 cbElements As Long
 cLocks As Long
 pvData As Long
 cElements As Long
 lLbound As Long
End Type

Type AnimSurfaceInfo
 Dib() As BGRAQUAD
 LDib() As Long
 TopRight As PointAPI
 Dims As DimsAPI
 halfW As Single
 halfH As Single
 CBWidth As Long
 SA1D As SAFEARRAY1D
 SA1D_L As SAFEARRAY1D
 EraseDib() As Long
 LBotLeftErase() As Long
 LTopLeftErase() As Long
 LEraseWidth() As Long
 EraseSpriteCount As Long
End Type

Public AirBrush As AirbrushStruct
 
Type PointAPI_3D
 px As Single
 py As Single
 pz As Single
End Type

Public Type DottedImage
 DScanLines() As BGRAQUAD
 UB As Long
 LB As Long
 ImgWide As Long
 ImgHigh As Long
 LoLf As PointAPI_3D
 LoRt As PointAPI_3D
 HiLf As PointAPI_3D
 HiRt As PointAPI_3D
End Type

Public px!
Public py!
Public ix!
Public iy!

Dim RX&
Dim RY&
Dim RXL&
Dim RXR&
Dim RYT&
Dim RYB&

Public Sub BlitAirbrush(Surf As AnimSurfaceInfo, X!, Y!, Optional bdoErase As Boolean = True, Optional bSolidDot As Boolean)
Dim sR2!
Dim sG2!
Dim sB2!
  
  AirBrush.IProcess.iy! = AirBrush.IProcess.int_x_def / (AirBrush.Dims.wWide / 2)
  
  px = X
  py = Y
  RX = RealRound2(px)
  RY = RealRound2(py)
  RXL = RealRound2(px - AirBrush.Dims.wDiv2)
  RXR = RealRound2(px + AirBrush.Dims.wDiv2)
  RYT = RealRound2(py + AirBrush.Dims.hDiv2)
  RYB = RealRound2(py - AirBrush.Dims.hDiv2)
 
  Left_& = RXL
  Top_& = RYT
 
  Right_& = Left_& + RXR - RXL
  Bot_& = Top_& - RYT + RYB
  
  If Bot_ < 0& Then
   AirBrush.IProcess.ClipBot& = -Bot_&
  Else
   AirBrush.IProcess.ClipBot& = 0&
  End If
  AirBrush.IProcess.DrawBot& = Bot_& + AirBrush.IProcess.ClipBot
 
  If Left_& < 0& Then
   AirBrush.IProcess.ClipLeft& = -Left_&
  Else
   AirBrush.IProcess.ClipLeft& = 0&
  End If
  AirBrush.IProcess.DrawLeft& = Left_& + AirBrush.IProcess.ClipLeft
  
  If Top_& - Surf.Dims.Height < 0& Then
   AirBrush.IProcess.DrawTop& = Top_&
  Else
   AirBrush.IProcess.DrawTop& = Surf.TopRight.Y
  End If

  If Right_ > Surf.TopRight.X Then
   AirBrush.IProcess.DrawRight = Surf.TopRight.X
  Else
   AirBrush.IProcess.DrawRight = Right_
  End If
 
  Left_ = AirBrush.IProcess.DrawTop - AirBrush.IProcess.DrawBot
 
  AirBrush.IProcess.DrawBot = AirBrush.IProcess.DrawBot * Surf.Dims.Width + AirBrush.IProcess.DrawLeft
  AirBrush.IProcess.DrawTop = AirBrush.IProcess.DrawBot + Left_ * Surf.Dims.Width
    
  AddDrawWidth = AirBrush.IProcess.DrawRight - AirBrush.IProcess.DrawLeft
  
  '=====================
  
  'how far brush center is from pixel center affects outcome.
  baseleft = (RXL - px + AirBrush.IProcess.ClipLeft) * AirBrush.IProcess.iy
  delta_y = (RYB - py + AirBrush.IProcess.ClipBot) * AirBrush.IProcess.iy
 
  If bdoErase Then
  
  Surf.EraseSpriteCount = Surf.EraseSpriteCount + 1&
   
  ReDim Preserve Surf.LBotLeftErase(1& To Surf.EraseSpriteCount)
  ReDim Preserve Surf.LTopLeftErase(1& To Surf.EraseSpriteCount)
  ReDim Preserve Surf.LEraseWidth(1& To Surf.EraseSpriteCount)
  
  Surf.LBotLeftErase(Surf.EraseSpriteCount) = AirBrush.IProcess.DrawBot
  Surf.LTopLeftErase(Surf.EraseSpriteCount) = AirBrush.IProcess.DrawTop ' / 4& 'Surf.LBotLeftErase(Surf.EraseSpriteCount) + (AirBrush.IProcess.DrawTop - AirBrush.IProcess.DrawBot) * Surf.Dims.Width
  Surf.LEraseWidth(Surf.EraseSpriteCount) = AddDrawWidth
  
  If AirBrush.CPMix = CSolid Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR = sR! + intensity! * (AirBrush.chRed - sR!)
     sG = sG! + intensity! * (AirBrush.chGreen - sG!)
     sB = sB! + intensity! * (AirBrush.chBlue - sB!)
     Surf.Dib(DrawX).Blue = sB
     Surf.Dib(DrawX).Green = sG
     Surf.Dib(DrawX).Red = sR
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      Surf.Dib(DrawX).Blue = sB + intensity! * (AirBrush.chBlue - sB!)
      Surf.Dib(DrawX).Green = sG + intensity! * (AirBrush.chGreen - sG!)
      Surf.Dib(DrawX).Red = sR + intensity! * (AirBrush.chRed - sR!)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If
  
  ElseIf AirBrush.CPMix = Inverse Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y! 'speeds up what's inside the next loop
   delta_x! = baseleft! 'starting at the left side with each new row.
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > AirBrush.chIntensity Then Bright = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR2 = 255! - sR
     sG2 = 255! - sG
     sB2 = 255! - sB
     Surf.Dib(DrawX).Blue = sB + intensity! * (sB2 - sB!)
     Surf.Dib(DrawX).Green = sG + intensity! * (sG2 - sG!)
     Surf.Dib(DrawX).Red = sR + intensity! * (sR2 - sR!)
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      sR2 = 255& - sR
      sG2 = 255& - sG
      sB2 = 255& - sB
      Surf.Dib(DrawX).Blue = sB + intensity! * (sB2 - sB!)
      Surf.Dib(DrawX).Green = sG + intensity! * (sG2 - sG!)
      Surf.Dib(DrawX).Red = sR + intensity! * (sR2 - sR!)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If
  
  ElseIf AirBrush.CPMix = CShift Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y! 'speeds up what's inside the next loop
   delta_x! = baseleft! 'starting at the left side with each new row.
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > AirBrush.chIntensity Then Bright = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     iB! = Surf.Dib(DrawX).Blue
     iG! = Surf.Dib(DrawX).Green
     iR! = Surf.Dib(DrawX).Red
     ColorShift
     Surf.Dib(DrawX).Blue = iB
     Surf.Dib(DrawX).Green = iG
     Surf.Dib(DrawX).Red = iR
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      iB! = Surf.Dib(DrawX).Blue
      iG! = Surf.Dib(DrawX).Green
      iR! = Surf.Dib(DrawX).Red
      ColorShift
      Surf.Dib(DrawX).Blue = iB
      Surf.Dib(DrawX).Green = iG
      Surf.Dib(DrawX).Red = iR
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If
  
  End If 'AirBrush.CPMix = CSolid
 
  
  Else 'not erasing
  
    
  If AirBrush.CPMix = CSolid Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     Surf.Dib(DrawX).Blue = sB + intensity! * (AirBrush.chBlue - sB!)
     Surf.Dib(DrawX).Green = sG + intensity! * (AirBrush.chGreen - sG!)
     Surf.Dib(DrawX).Red = sR + intensity! * (AirBrush.chRed - sR!)
     Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      Surf.Dib(DrawX).Blue = sB + intensity! * (AirBrush.chBlue - sB!)
      Surf.Dib(DrawX).Green = sG + intensity! * (AirBrush.chGreen - sG!)
      Surf.Dib(DrawX).Red = sR + intensity! * (AirBrush.chRed - sR!)
      Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If 'outlined
 
  ElseIf AirBrush.CPMix = Inverse Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > AirBrush.chIntensity Then Bright = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR2 = 255& - sR
     sG2 = 255& - sG
     sB2 = 255& - sB
     Surf.Dib(DrawX).Blue = sB + intensity! * (sB2 - sB!)
     Surf.Dib(DrawX).Green = sG + intensity! * (sG2 - sG!)
     Surf.Dib(DrawX).Red = sR + intensity! * (sR2 - sR!)
     Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      sR2 = 255& - sR
      sG2 = 255& - sG
      sB2 = 255& - sB
      Surf.Dib(DrawX).Blue = sB + intensity! * (sB2 - sB!)
      Surf.Dib(DrawX).Green = sG + intensity! * (sG2 - sG!)
      Surf.Dib(DrawX).Red = sR + intensity! * (sR2 - sR!)
      Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If 'outlined
  
  ElseIf AirBrush.CPMix = CShift Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > AirBrush.chIntensity Then Bright = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     iB! = Surf.Dib(DrawX).Blue
     iG! = Surf.Dib(DrawX).Green
     iR! = Surf.Dib(DrawX).Red
     ColorShift
     Surf.Dib(DrawX).Blue = iB
     Surf.Dib(DrawX).Green = iG
     Surf.Dib(DrawX).Red = iR
     Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      iB! = Surf.Dib(DrawX).Blue
      iG! = Surf.Dib(DrawX).Green
      iR! = Surf.Dib(DrawX).Red
      ColorShift
      Surf.Dib(DrawX).Blue = iB
      Surf.Dib(DrawX).Green = iG
      Surf.Dib(DrawX).Red = iR
      Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If 'bSolid
  
  End If 'AirBrush.CPMix
  
  End If 'erasing
  
End Sub

Public Sub BlitDottedImage(Surf As AnimSurfaceInfo, DotImage As DottedImage, Optional bSolidDot As Boolean = True, Optional bdoErase As Boolean = True)
Dim sR2!
Dim sG2!
Dim sB2!
Dim iix!
Dim iiy!
Dim X4&
Dim Y4&
Dim Elem&
Dim ixL!
Dim iyL!
Dim pxL!
Dim pyL!
Dim ixT!
Dim iyT!

  
  'DotImage.LoLf.px = Rnd * Surf.Dims.Width
  'DotImage.LoRt.px = DotImage.HiRt.px 'Rnd * Surf.Dims.Width
  'DotImage.HiLf.px = Rnd * Surf.Dims.Width
  'DotImage.HiRt.px = Rnd * Surf.Dims.Width
  'DotImage.LoLf.py = Rnd * Surf.Dims.Height
  'DotImage.LoRt.py = DotImage.LoLf.py 'Rnd * Surf.Dims.Height
  'DotImage.HiLf.py = Rnd * Surf.Dims.Height
  'DotImage.HiRt.py = Rnd * Surf.Dims.Height
  
  pxL = DotImage.LoLf.px
  pyL = DotImage.LoLf.py
  
  If DotImage.ImgWide > 0& Then
  Elem = DotImage.ImgWide - 1&
  ix = (DotImage.LoRt.px - pxL) / Elem 'DotImage.ImgWide
  iy = (DotImage.LoRt.py - pyL) / Elem 'DotImage.ImgWide
  
  ixT = (DotImage.HiRt.px - DotImage.HiLf.px) / Elem 'DotImage.ImgWide
  iyT = (DotImage.HiRt.py - DotImage.HiLf.py) / Elem 'DotImage.ImgWide
  
  If DotImage.ImgHigh > 0& Then
   Elem = DotImage.ImgHigh - 1&
   ixL = (DotImage.HiLf.px - pxL) / Elem 'DotImage.ImgHigh
   iyL = (DotImage.HiLf.py - pyL) / Elem 'DotImage.ImgHigh
   iix = (ixT - ix) / Elem
   iiy = (iyT - iy) / Elem
  
  LumenAirBrush 255, 1.9
  DimensionAirbrush 4, 4
  
  AirBrush.CPMix = CSolid
  
  Elem = 0
  
  For Y4 = 1 To DotImage.ImgHigh Step 1&
  px = pxL
  py = pyL
  For X4 = 1 To DotImage.ImgWide Step 1&
  
  AirBrush.chBlue = DotImage.DScanLines(Elem).Blue
  AirBrush.chGreen = DotImage.DScanLines(Elem).Green
  AirBrush.chRed = DotImage.DScanLines(Elem).Red
  
  'AirbrushClipper
  '=====================
 
  AirBrush.IProcess.iy! = AirBrush.IProcess.int_x_def / (AirBrush.Dims.wWide / 2)
  
  px = px
  py = py
  RX = RealRound2(px)
  RY = RealRound2(py)
  RXL = RealRound2(px - AirBrush.Dims.wDiv2)
  RXR = RealRound2(px + AirBrush.Dims.wDiv2)
  RYT = RealRound2(py + AirBrush.Dims.hDiv2)
  RYB = RealRound2(py - AirBrush.Dims.hDiv2)
 
  Left_& = RXL
  Top_& = RYT
 
  Right_& = Left_& + RXR - RXL
  Bot_& = Top_& - RYT + RYB
  
  If Bot_ < 0& Then
   AirBrush.IProcess.ClipBot& = -Bot_&
  Else
   AirBrush.IProcess.ClipBot& = 0&
  End If
  AirBrush.IProcess.DrawBot& = Bot_& + AirBrush.IProcess.ClipBot
 
  If Left_& < 0& Then
   AirBrush.IProcess.ClipLeft& = -Left_&
  Else
   AirBrush.IProcess.ClipLeft& = 0&
  End If
  AirBrush.IProcess.DrawLeft& = Left_& + AirBrush.IProcess.ClipLeft
  
  If Top_& - Surf.Dims.Height < 0& Then
   AirBrush.IProcess.DrawTop& = Top_&
  Else
   AirBrush.IProcess.DrawTop& = Surf.TopRight.Y
  End If

  If Right_ > Surf.TopRight.X Then
   AirBrush.IProcess.DrawRight = Surf.TopRight.X
  Else
   AirBrush.IProcess.DrawRight = Right_
  End If
 
  Left_ = AirBrush.IProcess.DrawTop - AirBrush.IProcess.DrawBot
 
  AirBrush.IProcess.DrawBot = AirBrush.IProcess.DrawBot * Surf.Dims.Width + AirBrush.IProcess.DrawLeft
  AirBrush.IProcess.DrawTop = AirBrush.IProcess.DrawBot + Left_ * Surf.Dims.Width
    
  AddDrawWidth = AirBrush.IProcess.DrawRight - AirBrush.IProcess.DrawLeft
  
  '=====================
  
  'how far brush center is from pixel center affects outcome.
  baseleft = (RXL - px + AirBrush.IProcess.ClipLeft) * AirBrush.IProcess.iy
  delta_y = (RYB - py + AirBrush.IProcess.ClipBot) * AirBrush.IProcess.iy
  
  If bdoErase Then
  
  Surf.EraseSpriteCount = Surf.EraseSpriteCount + 1&
   
  ReDim Preserve Surf.LBotLeftErase(1& To Surf.EraseSpriteCount)
  ReDim Preserve Surf.LTopLeftErase(1& To Surf.EraseSpriteCount)
  ReDim Preserve Surf.LEraseWidth(1& To Surf.EraseSpriteCount)
  
  Surf.LBotLeftErase(Surf.EraseSpriteCount) = AirBrush.IProcess.DrawBot
  Surf.LTopLeftErase(Surf.EraseSpriteCount) = AirBrush.IProcess.DrawTop ' / 4& 'Surf.LBotLeftErase(Surf.EraseSpriteCount) + (AirBrush.IProcess.DrawTop - AirBrush.IProcess.DrawBot) * Surf.Dims.Width
  Surf.LEraseWidth(Surf.EraseSpriteCount) = AddDrawWidth
  
  If AirBrush.CPMix = CSolid Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR = sR! + intensity! * (AirBrush.chRed - sR!)
     sG = sG! + intensity! * (AirBrush.chGreen - sG!)
     sB = sB! + intensity! * (AirBrush.chBlue - sB!)
     Surf.Dib(DrawX).Blue = sB
     Surf.Dib(DrawX).Green = sG
     Surf.Dib(DrawX).Red = sR
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      Surf.Dib(DrawX).Blue = sB + intensity! * (AirBrush.chBlue - sB!)
      Surf.Dib(DrawX).Green = sG + intensity! * (AirBrush.chGreen - sG!)
      Surf.Dib(DrawX).Red = sR + intensity! * (AirBrush.chRed - sR!)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If
  
  ElseIf AirBrush.CPMix = Inverse Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y! 'speeds up what's inside the next loop
   delta_x! = baseleft! 'starting at the left side with each new row.
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > AirBrush.chIntensity Then Bright = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR2 = 255! - sR
     sG2 = 255! - sG
     sB2 = 255! - sB
     Surf.Dib(DrawX).Blue = sB + intensity! * (sB2 - sB!)
     Surf.Dib(DrawX).Green = sG + intensity! * (sG2 - sG!)
     Surf.Dib(DrawX).Red = sR + intensity! * (sR2 - sR!)
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      sR2 = 255& - sR
      sG2 = 255& - sG
      sB2 = 255& - sB
      Surf.Dib(DrawX).Blue = sB + intensity! * (sB2 - sB!)
      Surf.Dib(DrawX).Green = sG + intensity! * (sG2 - sG!)
      Surf.Dib(DrawX).Red = sR + intensity! * (sR2 - sR!)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If
  
  ElseIf AirBrush.CPMix = CShift Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y! 'speeds up what's inside the next loop
   delta_x! = baseleft! 'starting at the left side with each new row.
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > AirBrush.chIntensity Then Bright = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     iB! = Surf.Dib(DrawX).Blue
     iG! = Surf.Dib(DrawX).Green
     iR! = Surf.Dib(DrawX).Red
     ColorShift
     Surf.Dib(DrawX).Blue = iB
     Surf.Dib(DrawX).Green = iG
     Surf.Dib(DrawX).Red = iR
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      iB! = Surf.Dib(DrawX).Blue
      iG! = Surf.Dib(DrawX).Green
      iR! = Surf.Dib(DrawX).Red
      ColorShift
      Surf.Dib(DrawX).Blue = iB
      Surf.Dib(DrawX).Green = iG
      Surf.Dib(DrawX).Red = iR
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If
  
  End If 'AirBrush.CPMix = CSolid
 
  
  Else 'not erasing
  
    
  If AirBrush.CPMix = CSolid Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     Surf.Dib(DrawX).Blue = sB + intensity! * (AirBrush.chBlue - sB!)
     Surf.Dib(DrawX).Green = sG + intensity! * (AirBrush.chGreen - sG!)
     Surf.Dib(DrawX).Red = sR + intensity! * (AirBrush.chRed - sR!)
     Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      Surf.Dib(DrawX).Blue = sB + intensity! * (AirBrush.chBlue - sB!)
      Surf.Dib(DrawX).Green = sG + intensity! * (AirBrush.chGreen - sG!)
      Surf.Dib(DrawX).Red = sR + intensity! * (AirBrush.chRed - sR!)
      Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If 'outlined
 
  ElseIf AirBrush.CPMix = Inverse Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > AirBrush.chIntensity Then Bright = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR2 = 255& - sR
     sG2 = 255& - sG
     sB2 = 255& - sB
     Surf.Dib(DrawX).Blue = sB + intensity! * (sB2 - sB!)
     Surf.Dib(DrawX).Green = sG + intensity! * (sG2 - sG!)
     Surf.Dib(DrawX).Red = sR + intensity! * (sR2 - sR!)
     Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      sR2 = 255& - sR
      sG2 = 255& - sG
      sB2 = 255& - sB
      Surf.Dib(DrawX).Blue = sB + intensity! * (sB2 - sB!)
      Surf.Dib(DrawX).Green = sG + intensity! * (sG2 - sG!)
      Surf.Dib(DrawX).Red = sR + intensity! * (sR2 - sR!)
      Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If 'outlined
  
  ElseIf AirBrush.CPMix = CShift Then
  
  If bSolidDot Then
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > AirBrush.chIntensity Then Bright = AirBrush.chIntensity
     intensity = AirBrush.IProcess.intens * Bright&
     iB! = Surf.Dib(DrawX).Blue
     iG! = Surf.Dib(DrawX).Green
     iR! = Surf.Dib(DrawX).Red
     ColorShift
     Surf.Dib(DrawX).Blue = iB
     Surf.Dib(DrawX).Green = iG
     Surf.Dib(DrawX).Red = iR
     Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  Else 'circle outline
  
  For DrawY& = AirBrush.IProcess.DrawBot& To AirBrush.IProcess.DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = baseleft!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight& Step 1&
    Bright& = AirBrush.IProcess.int_x_def - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > AirBrush.chIntensity Then Bright& = AirBrush.I2 - Bright
     If Bright > 0& Then
      intensity = AirBrush.IProcess.intens * Bright&
      iB! = Surf.Dib(DrawX).Blue
      iG! = Surf.Dib(DrawX).Green
      iR! = Surf.Dib(DrawX).Red
      ColorShift
      Surf.Dib(DrawX).Blue = iB
      Surf.Dib(DrawX).Green = iG
      Surf.Dib(DrawX).Red = iR
      Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
     End If
    End If
    delta_x! = delta_x! + AirBrush.IProcess.iy!
   Next DrawX
   delta_y! = delta_y! + AirBrush.IProcess.iy!
  Next DrawY
  
  End If 'bSolid
  
  End If 'AirBrush.CPMix
  
  End If 'erasing
  
  Elem = Elem + 1&
  px = px + ix
  py = py + iy
  Next 'px
  ix = ix + iix
  iy = iy + iiy
  pxL = pxL + ixL
  pyL = pyL + iyL
  Next 'py
  
  End If
  
  End If

End Sub

Public Sub ColorShift()

 If iR < iB Then
  If iR < iG Then
   If iG < iB Then
    bytMaxMin_diff = iB - iR
    iG = iG - bytMaxMin_diff * intensity
    If iG < iR Then
     iSubt = iR - iG
     iG = iR
     iR = iR + iSubt
    End If
   Else
    bytMaxMin_diff = iG - iR
    iB = iB + bytMaxMin_diff * intensity
    If iB > iG Then
     iSubt = iB - iG
     iB = iG
     iG = iG - iSubt
    End If
   End If
  Else
   bytMaxMin_diff = iB - iG
   iR = iR + bytMaxMin_diff * intensity
   If iR > iB Then
    iSubt = iR - iB
    iR = iB
    iB = iB - iSubt
   End If
  End If
 ElseIf iR > iG Then
  If iB < iG Then
   bytMaxMin_diff = iR - iB
   iG = iG + bytMaxMin_diff * intensity
   If iG > iR Then
    iSubt = iG - iR
    iG = iR
    iR = iR - iSubt
   End If
  Else
   bytMaxMin_diff = iR - iG
   iB = iB - bytMaxMin_diff * intensity
   If iB < iG Then
    iSubt = iG - iB
    iB = iG
    iG = iG + iSubt
   End If
  End If
 Else
  bytMaxMin_diff = iG - iB
  iR = iR - bytMaxMin_diff * intensity
  If iR < iB Then
   iSubt = iB - iR
   iR = iB
   iB = iB + iSubt
  End If
 End If
 
End Sub

Private Sub AirbrushClipper(Surf As AnimSurfaceInfo)
 
 'This sub writes to AirBrush.IProcess.DrawBot, ..Top, ..Left,
 ' ..Right, ..ClipBot, ..ClipLeft
 
 'Blitairbrush uses these dimensions to blit within array
 'bounds.  It's an improvement over inbounds test per pixel.
 
 'You may note that .ClipTop is also a member of AirBrush.Iprocess
 'It is only used for processing in this sub
 
 Left_& = RealRound2(AirBrush.IProcess.blit_x) - AirBrush.IProcess.SpriteGrid.AddHalfW
 Top_& = Surf.Dims.Height - RealRound2(AirBrush.IProcess.blit_y) + AirBrush.IProcess.SpriteGrid.AddHalfH
 
 Right_& = Left_& + AirBrush.IProcess.SpriteGrid.RightM1
 Bot_& = Top_& - AirBrush.IProcess.SpriteGrid.HighM1
 
 AirBrush.IProcess.ClipTop& = Top_& - Surf.Dims.Height
  
 If Bot_ < 0& Then
  AirBrush.IProcess.ClipBot& = -Bot_& '* Surf.Dims.Width
  AirBrush.IProcess.DrawBot& = 0&
 Else
  AirBrush.IProcess.ClipBot = 0&
  AirBrush.IProcess.DrawBot& = Bot_& '* Surf.Dims.Width
 End If
 
 If Left_& < 0& Then
  AirBrush.IProcess.ClipLeft& = -Left_&
  AirBrush.IProcess.DrawLeft& = 0& '+ AirBrush.IProcess.DrawBot
 Else
  AirBrush.IProcess.ClipLeft& = 0&
  AirBrush.IProcess.DrawLeft& = Left_& '+ AirBrush.IProcess.DrawBot
 End If
 
 If AirBrush.IProcess.ClipTop& < 0& Then
  AirBrush.IProcess.DrawTop& = Top_&
 Else
  AirBrush.IProcess.DrawTop& = Surf.TopRight.Y
 End If

 If Right_ > Surf.TopRight.X Then
  AirBrush.IProcess.DrawRight = Surf.TopRight.X
 Else
  AirBrush.IProcess.DrawRight = Right_
 End If
 
 Left_ = AirBrush.IProcess.DrawTop - AirBrush.IProcess.DrawBot
 
 AirBrush.IProcess.DrawBot = AirBrush.IProcess.DrawBot * Surf.Dims.Width + AirBrush.IProcess.DrawLeft '* 4&
 AirBrush.IProcess.DrawTop = AirBrush.IProcess.DrawBot + Left_ * Surf.Dims.Width

End Sub


Public Sub EraseSprites(Surf As AnimSurfaceInfo)
Dim LBotLeft&
Dim LTopLeft&
Dim LEraseWide&

 For N& = 1& To Surf.EraseSpriteCount&
  
  LBotLeft = Surf.LBotLeftErase(N)
  LTopLeft = Surf.LTopLeftErase(N)
  LEraseWide = Surf.LEraseWidth(N)
  For Loca& = LBotLeft To LTopLeft Step Surf.Dims.Width
   AddDrawWidth = Loca + LEraseWide
   For DrawX = Loca To AddDrawWidth
    Surf.LDib(DrawX) = Surf.EraseDib(DrawX)
   Next
  Next
  
 Next N& 'Next Sprite
 
 Surf.EraseSpriteCount = 0&

End Sub

'These can be called by user to quickly set specified elements
Public Sub DimensionAirbrush(Wide!, High!)
 AirBrush.Dims.wWide = Wide
 AirBrush.Dims.hHigh = High
 AirBrush.Dims.wDiv2 = Wide / 2
 AirBrush.Dims.hDiv2 = High / 2
End Sub
Public Sub ColorAirBrush(ByVal chRed As Byte, ByVal chGreen As Byte, ByVal chBlue As Byte)
 AirBrush.chBlue = chBlue
 AirBrush.chGreen = chGreen
 AirBrush.chRed = chRed
End Sub
Public Sub LumenAirBrush(ByVal chIntensity As Byte, definition!)
Dim tmpIntensity!
 
 AirBrush.definition = definition
 AirBrush.chIntensity = chIntensity
 
 AirBrush.IProcess.int_x_def = AirBrush.definition * AirBrush.chIntensity
 
 tmpIntensity = AirBrush.chIntensity / B255
 AirBrush.IProcess.intens = tmpIntensity / B255
 AirBrush.I2 = AirBrush.chIntensity * 2&
 
End Sub


'Here are my rounding functions.  Not really exciting, but necessary.
'VB says that Round(0.5) = 0 and Round(1.5) = 2.
Public Function RealRound(ByVal sngValue!) As Long
Dim diff!
 'This function rounds .5 up
 
 RealRound = Int(sngValue)
 diff = sngValue - RealRound
 If diff >= 0.5! Then RealRound = RealRound + 1&

End Function
Public Function RealRound2(ByVal sngValue!) As Long
Dim diff!
 'This function rounds .5 down
 
 RealRound2 = Int(sngValue)
 diff = sngValue - RealRound2
 If diff > 0.5! Then RealRound2 = RealRound2 + 1&

End Function

