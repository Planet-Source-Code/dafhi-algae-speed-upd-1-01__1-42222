Attribute VB_Name = "modFCG"
Option Explicit

Private Type RectDims
 Width As Long
 Height As Long
 WideM1 As Long
 HighM1 As Long
 LowLeftPos As PointAPI
 TopRightPos As PointAPI
End Type

Private Type RGBDiffAPI
 sRed As Single
 sGrn As Single
 sBlu As Single
End Type

Private Type FCG_System
 LeftRGBdelta As RGBDiffAPI
 TopRGBdelta As RGBDiffAPI
 BottomRGBdelta As RGBDiffAPI
 LeftRGBi As RGBDiffAPI
 TopRGBi As RGBDiffAPI
 BottomRGBi As RGBDiffAPI
 Dims As RectDims
End Type

Public Type FCGRect
 LowLeft As RGBTriple
 LowRight As RGBTriple
 TopLeft As RGBTriple
 TopRight As RGBTriple
 IProcess As FCG_System
End Type

Public FourColorGradient As FCGRect

'DrawFCG and DrawFCG2
Dim iBlue!
Dim iGreen!
Dim iRed!
Dim iiBlu!
Dim iiGrn!
Dim iiRed!
Dim iiBlu2!
Dim iiGrn2!
Dim iiRed2!
Dim sngBlue!
Dim sngGreen!
Dim sngRed!
Dim vertBlue!
Dim vertGreen!
Dim vertRed!

Dim DrawTop&
Dim DrawBot&
Dim DrawLeft&
Dim DrawRight&

Dim DrawX2&
Dim VertLDib&

Dim AddDrawWidth&

Public Sub FCG_ColorLowLeft(Red&, Green&, Blue&)
 FourColorGradient.LowLeft.Red = Red
 FourColorGradient.LowLeft.Green = Green
 FourColorGradient.LowLeft.Blue = Blue
End Sub
Public Sub FCG_ColorLowRight(Red&, Green&, Blue&)
 FourColorGradient.LowRight.Red = Red
 FourColorGradient.LowRight.Green = Green
 FourColorGradient.LowRight.Blue = Blue
End Sub
Public Sub FCG_ColorTopLeft(Red&, Green&, Blue&)
 FourColorGradient.TopLeft.Red = Red
 FourColorGradient.TopLeft.Green = Green
 FourColorGradient.TopLeft.Blue = Blue
End Sub
Public Sub FCG_ColorTopRight(Red&, Green&, Blue&)
 FourColorGradient.TopRight.Red = Red
 FourColorGradient.TopRight.Green = Green
 FourColorGradient.TopRight.Blue = Blue
End Sub

Public Sub SetLowLeftCorner(X&, Y&)
 FourColorGradient.IProcess.Dims.LowLeftPos.X = X
 FourColorGradient.IProcess.Dims.LowLeftPos.Y = Y
End Sub
Public Sub SetFCGDims(Width&, Height&)
With FourColorGradient.IProcess.Dims
 .Width = Width
 .Height = Height
 .WideM1 = Width - 1
 .HighM1 = Height - 1
 .TopRightPos.X = .LowLeftPos.X + .WideM1
 .TopRightPos.Y = .LowLeftPos.Y + .HighM1
End With
End Sub


Public Sub WrapFCGToAnimSurf(Surf As AnimSurfaceInfo)
 
  If Surf.Dims.Width > 0& Then
  
  SetLowLeftCorner 0, 0
  SetFCGDims Surf.Dims.Width, Surf.Dims.Height
  
  With FourColorGradient
  .IProcess.BottomRGBdelta.sBlu = .LowRight.Blue - .LowLeft.Blue
  .IProcess.BottomRGBdelta.sGrn = .LowRight.Green - .LowLeft.Green
  .IProcess.BottomRGBdelta.sRed = .LowRight.Red - .LowLeft.Red

  .IProcess.TopRGBdelta.sBlu = .TopRight.Blue - .TopLeft.Blue
  .IProcess.TopRGBdelta.sGrn = .TopRight.Green - .TopLeft.Green
  .IProcess.TopRGBdelta.sRed = .TopRight.Red - .TopLeft.Red
  
  .IProcess.LeftRGBdelta.sBlu = .TopLeft.Blue - .LowLeft.Blue
  .IProcess.LeftRGBdelta.sGrn = .TopLeft.Green - .LowLeft.Green
  .IProcess.LeftRGBdelta.sRed = .TopLeft.Red - .LowLeft.Red
  End With
  
  With FourColorGradient.IProcess
   
   If .Dims.WideM1 > 0 Then
   .TopRGBi.sBlu = .TopRGBdelta.sBlu / .Dims.WideM1
   .TopRGBi.sGrn = .TopRGBdelta.sGrn / .Dims.WideM1
   .TopRGBi.sRed = .TopRGBdelta.sRed / .Dims.WideM1
   .BottomRGBi.sBlu = .BottomRGBdelta.sBlu / .Dims.WideM1
   .BottomRGBi.sGrn = .BottomRGBdelta.sGrn / .Dims.WideM1
   .BottomRGBi.sRed = .BottomRGBdelta.sRed / .Dims.WideM1
   End If
   
   If .Dims.HighM1 > 0 Then
   .LeftRGBi.sBlu = .LeftRGBdelta.sBlu / .Dims.HighM1
   .LeftRGBi.sGrn = .LeftRGBdelta.sGrn / .Dims.HighM1
   .LeftRGBi.sRed = .LeftRGBdelta.sRed / .Dims.HighM1
    iiRed = (.TopRGBi.sRed - .BottomRGBi.sRed) / .Dims.HighM1
    iiGrn = (.TopRGBi.sGrn - .BottomRGBi.sGrn) / .Dims.HighM1
    iiBlu = (.TopRGBi.sBlu - .BottomRGBi.sBlu) / .Dims.HighM1
   End If
  
   iRed = .BottomRGBi.sRed
   iGreen = .BottomRGBi.sGrn
   iBlue = .BottomRGBi.sBlu
   
   DrawLeft = .Dims.TopRightPos.Y - .Dims.LowLeftPos.Y
   DrawBot = .Dims.LowLeftPos.Y * Surf.Dims.Width + .Dims.LowLeftPos.X
   DrawTop = DrawBot + DrawLeft * Surf.Dims.Width
   
   AddDrawWidth = .Dims.TopRightPos.X - .Dims.LowLeftPos.X
   'AddDrawWidthBytes = AddDrawWidth * 4&
   
  End With
  
  vertBlue = FourColorGradient.LowLeft.Blue
  vertGreen = FourColorGradient.LowLeft.Green
  vertRed = FourColorGradient.LowLeft.Red
      
  For DrawY = DrawBot To DrawTop Step Surf.Dims.Width
   DrawRight = DrawY + AddDrawWidth
   sngBlue = vertBlue
   sngGreen = vertGreen
   sngRed = vertRed
   vertBlue = vertBlue + FourColorGradient.IProcess.LeftRGBi.sBlu
   vertGreen! = vertGreen + FourColorGradient.IProcess.LeftRGBi.sGrn
   vertRed! = vertRed + FourColorGradient.IProcess.LeftRGBi.sRed
   For DrawX = DrawY To DrawRight Step 1&
    Surf.Dib(DrawX).Blue = sngBlue
    Surf.Dib(DrawX).Green = sngGreen
    Surf.Dib(DrawX).Red = sngRed
    sngBlue = sngBlue + iBlue
    sngGreen = sngGreen + iGreen
    sngRed = sngRed + iRed
   Next DrawX
   iBlue = iBlue + iiBlu
   iGreen = iGreen + iiGrn
   iRed = iRed + iiRed
  Next DrawY
  
  DrawTop = Surf.SA1D_L.cElements - 1
  For DrawBot = 0 To DrawTop
   Surf.EraseDib(DrawBot) = Surf.LDib(DrawBot)
  Next
  
  End If

End Sub

Public Sub DrawFCGToAnimSurf(Surf As AnimSurfaceInfo, ByVal Left&, ByVal Top&, Width&, Height&)
Dim ClipLeft&
Dim ClipBot&
  
  SetLowLeftCorner Left, Surf.Dims.Height - Top - Height
  SetFCGDims Width, Height
  
  If Left < 0 Then
   ClipLeft = -Left
   DrawLeft = 0
  Else
   DrawLeft = Left '* 4&
  End If
  
  If FourColorGradient.IProcess.Dims.LowLeftPos.Y < 0 Then
   DrawBot = 0
   ClipBot = -FourColorGradient.IProcess.Dims.LowLeftPos.Y
  Else
   DrawBot = FourColorGradient.IProcess.Dims.LowLeftPos.Y * Surf.Dims.Width
  End If
  DrawBot = DrawBot + DrawLeft
    
  If FourColorGradient.IProcess.Dims.TopRightPos.Y > Surf.TopRight.Y Then
   DrawTop = Surf.TopRight.Y * Surf.Dims.Width
  Else
   DrawTop = FourColorGradient.IProcess.Dims.TopRightPos.Y * Surf.Dims.Width
  End If
  DrawTop = DrawTop + DrawLeft
  
  If FourColorGradient.IProcess.Dims.TopRightPos.X > Surf.TopRight.X Then
   DrawRight = Surf.Dims.Width - 1&
  Else
   DrawRight = FourColorGradient.IProcess.Dims.TopRightPos.X
  End If
    
  AddDrawWidth = DrawRight - DrawLeft
  
  With FourColorGradient
  .IProcess.BottomRGBdelta.sBlu = .LowRight.Blue - .LowLeft.Blue
  .IProcess.BottomRGBdelta.sGrn = .LowRight.Green - .LowLeft.Green
  .IProcess.BottomRGBdelta.sRed = .LowRight.Red - .LowLeft.Red

  .IProcess.TopRGBdelta.sBlu = .TopRight.Blue - .TopLeft.Blue
  .IProcess.TopRGBdelta.sGrn = .TopRight.Green - .TopLeft.Green
  .IProcess.TopRGBdelta.sRed = .TopRight.Red - .TopLeft.Red
  
  .IProcess.LeftRGBdelta.sBlu = .TopLeft.Blue - .LowLeft.Blue
  .IProcess.LeftRGBdelta.sGrn = .TopLeft.Green - .LowLeft.Green
  .IProcess.LeftRGBdelta.sRed = .TopLeft.Red - .LowLeft.Red
  End With
  
  With FourColorGradient.IProcess
   
   If .Dims.WideM1 > 0& Then
   .TopRGBi.sBlu = .TopRGBdelta.sBlu / .Dims.WideM1
   .TopRGBi.sGrn = .TopRGBdelta.sGrn / .Dims.WideM1
   .TopRGBi.sRed = .TopRGBdelta.sRed / .Dims.WideM1
   .BottomRGBi.sBlu = .BottomRGBdelta.sBlu / .Dims.WideM1
   .BottomRGBi.sGrn = .BottomRGBdelta.sGrn / .Dims.WideM1
   .BottomRGBi.sRed = .BottomRGBdelta.sRed / .Dims.WideM1
   End If
   
   If .Dims.HighM1 > 0& Then
   .LeftRGBi.sBlu = .LeftRGBdelta.sBlu / .Dims.HighM1
   .LeftRGBi.sGrn = .LeftRGBdelta.sGrn / .Dims.HighM1
   .LeftRGBi.sRed = .LeftRGBdelta.sRed / .Dims.HighM1
    iiRed = (.TopRGBi.sRed - .BottomRGBi.sRed) / .Dims.HighM1
    iiGrn = (.TopRGBi.sGrn - .BottomRGBi.sGrn) / .Dims.HighM1
    iiBlu = (.TopRGBi.sBlu - .BottomRGBi.sBlu) / .Dims.HighM1
   End If
  
   iRed = .BottomRGBi.sRed + iiRed * ClipBot
   iGreen = .BottomRGBi.sGrn + iiGrn * ClipBot
   iBlue = .BottomRGBi.sBlu + iiBlu * ClipBot
   
   sngBlue = ClipLeft * .BottomRGBi.sBlu
   sngGreen = ClipLeft * .BottomRGBi.sGrn
   sngRed = ClipLeft * .BottomRGBi.sRed
   
   iiBlu2 = .LeftRGBi.sBlu + _
    (ClipLeft * .TopRGBi.sBlu - sngBlue) / (.Dims.HighM1)
   iiGrn2 = .LeftRGBi.sGrn + _
    (ClipLeft * .TopRGBi.sGrn - sngGreen) / (.Dims.HighM1)
   iiRed2 = .LeftRGBi.sRed + _
    (ClipLeft * .TopRGBi.sRed - sngRed) / (.Dims.HighM1)
  
  End With
  
  With FourColorGradient
  vertBlue = .LowLeft.Blue + sngBlue + ClipBot * iiBlu2
  vertGreen = .LowLeft.Green + sngGreen + ClipBot * iiGrn2
  vertRed = .LowLeft.Red + sngRed + ClipBot * iiRed2
  End With
  
  For DrawY = DrawBot To DrawTop Step Surf.Dims.Width
   DrawRight = DrawY + AddDrawWidth
   sngBlue = vertBlue
   sngGreen = vertGreen
   sngRed = vertRed
   vertBlue = vertBlue + iiBlu2
   vertGreen! = vertGreen + iiGrn2
   vertRed! = vertRed + iiRed2
   DrawX2 = DrawLeft
   For DrawX = DrawY To DrawRight Step 1&
    Surf.Dib(DrawX).Blue = sngBlue
    Surf.Dib(DrawX).Green = sngGreen
    Surf.Dib(DrawX).Red = sngRed
    Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    sngBlue = sngBlue + iBlue
    sngGreen = sngGreen + iGreen
    sngRed = sngRed + iRed
    DrawX2 = DrawX2 + 1&
   Next
   iBlue = iBlue + iiBlu
   iGreen = iGreen + iiGrn
   iRed = iRed + iiRed
  Next
  
End Sub

