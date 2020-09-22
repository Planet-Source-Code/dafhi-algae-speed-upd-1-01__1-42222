Attribute VB_Name = "modBuf"
Option Explicit

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As BGRAQUAD
End Type

Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type


Public BM As BITMAP

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)
Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any) As Long
Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc&, pBitmapInfo As BITMAPINFO, ByVal un&, lplpVoid&, ByVal handle&, ByVal dw&) As Long
Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, IPic As IPicture) As Long

Function CreatePicture(ByVal nWidth&, ByVal nHeight&, ByVal BitDepth&) As Picture
Dim Pic As PicBmp, IID_IDispatch As GUID
Dim BMI As BITMAPINFO
With BMI.bmiHeader
.biSize = Len(BMI.bmiHeader)
.biWidth = nWidth
.biHeight = nHeight
.biPlanes = 1
.biBitCount = BitDepth
End With
Pic.hBmp = CreateDIBSection(0, BMI, 0, 0, 0, 0)
IID_IDispatch.Data1 = &H20400: IID_IDispatch.Data4(0) = &HC0: IID_IDispatch.Data4(7) = &H46
Pic.Size = Len(Pic)
Pic.Type = vbPicTypeBitmap
OleCreatePictureIndirect Pic, IID_IDispatch, 1, CreatePicture
If CreatePicture = 0 Then Set CreatePicture = Nothing
End Function

Public Sub AnimPicSurface(Obj As Object, Surf As AnimSurfaceInfo)
 
 Surf.Dims.Width = Obj.ScaleWidth
 Surf.Dims.Height = Obj.ScaleHeight
 Surf.TopRight.X = Surf.Dims.Width - 1
 Surf.TopRight.Y = Surf.Dims.Height - 1
 
 Surf.halfW = Surf.Dims.Width / 2&
 Surf.halfH = Surf.Dims.Height / 2&
 
 Surf.EraseSpriteCount = 0
 
 'Destroy any pointer this array may have had
 CopyMemory ByVal VarPtrArray(Surf.Dib), 0&, 4
 
 If Surf.Dims.Height > 0 Then
 
 'Allocate memory to the Refresh buffer
 Obj.Picture = CreatePicture(Surf.Dims.Width, Surf.Dims.Height, 32)
 GetObjectAPI Obj.Picture, Len(BM), BM
 With Surf.SA1D
 .cbElements = 4
 .cDims = 1
 .lLbound = 0
 .cElements = BM.bmHeight * BM.bmWidth
 .pvData = BM.bmBits
 End With
 CopyMemory ByVal VarPtrArray(Surf.Dib), VarPtr(Surf.SA1D), 4
 
 With Surf.SA1D_L
 .cbElements = 4
 .cDims = 1
 .lLbound = 0
 .cElements = BM.bmHeight * BM.bmWidth
 .pvData = BM.bmBits
 ReDim Surf.EraseDib(.cElements)
 End With
 CopyMemory ByVal VarPtrArray(Surf.LDib), VarPtr(Surf.SA1D_L), 4
 
 Surf.CBWidth = BM.bmWidthBytes
 End If
 
End Sub

