Attribute VB_Name = "modFileHandler"
Option Explicit

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Dim BMPadBytes&
Dim BMPFileHeader As BITMAPFILEHEADER   'Holds the file header
Dim BMPInfoHeader As BITMAPINFOHEADER   'Holds the info header
Dim BMPData() As Byte                   'Holds the pixel data


Dim XLng& '1d.  XLng marks the position.  YLng is a reference.
Dim YLng&
Dim AddDrawWidthBytes&
Dim DrawRight&
Dim TopLeft&
Dim WidthBytes&
Dim X_Max&
Dim Y_Max&
Dim ClipTop&

Public Sub TruecolorBmpToAnimSurf(Surf As AnimSurfaceInfo, strFilename$)

 TrueColorBMPToData strFilename
 
 'These are used to reference 1d array file bytes
 WidthBytes& = BMPInfoHeader.biWidth * 3& + BMPadBytes

 If WidthBytes > 0 Then
 
  If BMPInfoHeader.biWidth > Surf.Dims.Width Then
   X_Max = Surf.TopRight.X
  Else
   X_Max = BMPInfoHeader.biWidth - 1
  End If
  
  If BMPInfoHeader.biHeight > Surf.Dims.Height Then
   Y_Max = Surf.TopRight.Y
   DrawY = 0
  Else
   Y_Max = BMPInfoHeader.biHeight - 1
   DrawY = Surf.Dims.Height - BMPInfoHeader.biHeight
   ClipTop = DrawY
  End If
 
  DrawRight = X_Max * 3
  TopLeft = WidthBytes * Y_Max
  X_Max = ClipTop * Surf.Dims.Width
  For YLng& = 0& To TopLeft Step WidthBytes
   DrawX = X_Max
   AddDrawWidthBytes& = YLng& + DrawRight&
   For XLng& = YLng& To AddDrawWidthBytes& Step 3&
    Blue = BMPData(XLng&)
    Green = BMPData(XLng& + 1&)
    Red = BMPData(XLng& + 2&)
    Surf.Dib(DrawX).Blue = Blue
    Surf.Dib(DrawX).Green = Green
    Surf.Dib(DrawX).Red = Red
    DrawX = DrawX + 1&
   Next XLng
   X_Max = X_Max + Surf.Dims.Width
  Next YLng

 End If 'Widthbytes > 0
 
 DrawY = Surf.SA1D_L.cElements - 1
 For DrawX = 0 To DrawY
  Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
 Next
 
End Sub

Public Sub TruecolorBmpToDottedImage(DotImage As DottedImage, strFilename$)
Dim TrackX&
Dim TrackY&
Dim ScanLineWidthBytes&

 TrueColorBMPToData strFilename
  
 'These are used to reference 1d array file bytes
 WidthBytes& = BMPInfoHeader.biWidth * 3& + BMPadBytes

 If WidthBytes > 0 Then
 
  DotImage.ImgWide = BMPInfoHeader.biWidth
  DotImage.ImgHigh = BMPInfoHeader.biHeight
  
  DotImage.LB = 0&
  DotImage.UB = DotImage.ImgHigh * DotImage.ImgWide * 4&
  
  ReDim DotImage.DScanLines(DotImage.LB To DotImage.UB)
  
  X_Max = DotImage.ImgWide - 1
  DrawRight = X_Max * 3
  
  TopLeft = WidthBytes * (DotImage.ImgHigh - 1&)
  ScanLineWidthBytes = DotImage.ImgWide '* 4&
  For YLng& = 0& To TopLeft Step WidthBytes
   DrawX = TrackY
   AddDrawWidthBytes& = YLng& + DrawRight&
   For XLng& = YLng& To AddDrawWidthBytes& Step 3&
    Blue = BMPData(XLng&)
    Green = BMPData(XLng& + 1&)
    Red = BMPData(XLng& + 2&)
    DotImage.DScanLines(DrawX).Blue = Blue
    DotImage.DScanLines(DrawX).Green = Green
    DotImage.DScanLines(DrawX).Red = Red
    DrawX = DrawX + 1&
   Next XLng
   TrackY = TrackY + ScanLineWidthBytes
  Next YLng

 End If 'Widthbytes > 0
  
End Sub

Private Sub TrueColorBMPToData(strFilename$)
 Open (App.Path & "\" & strFilename) For Binary As #1
   Get #1, 1, BMPFileHeader
   Get #1, , BMPInfoHeader
   With BMPInfoHeader
    N = 3 * .biWidth '(red, green, blue) * width
    BMPadBytes = ((N + 3) And &HFFFFFFFC) - N
    ReDim BMPData(.biHeight * (BMPadBytes + .biWidth * .biBitCount / 8))
   End With
   Get #1, , BMPData
 Close #1
End Sub
