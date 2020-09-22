Attribute VB_Name = "modPublic"
Option Explicit

Public N& 'Loops

Public DrawX&
Public DrawY&

Public Type PointAPI
 X As Long
 Y As Long
End Type

Public Type DimsAPI
 Width As Long
 Height As Long
End Type

Public Type RGBTriple
 Red As Long
 Green As Long
 Blue As Long
End Type

Type PrecisionPointAPI
 px As Single
 py As Single
End Type

Public Blue&
Public Green%
Public Red%
Public Alpha&

Public Const BYT2 As Byte = 2
Public Const B255 As Byte = 255

Public Const pi As Single = 3.14159265
Public Const twopi As Single = 2 * pi



'=================================================
'BGR2 = Point(X, Y)
'Blue = (BGR2 And 16711680) / 65536
'Green = (BGR2 And 65280) / 256&
'Red = BGR2 And B255

