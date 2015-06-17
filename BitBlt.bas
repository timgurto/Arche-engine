Attribute VB_Name = "X__external_mdlBitBlt"
'***EXTERNAL MODULE***
'Bit Block Transfer Functions
'KPD Team
'*********************


Option Explicit

Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) _
    As Long

Declare Function StretchBlt Lib "gdi32" _
    (ByVal hdc As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, _
    ByVal dwRop As Long) _
As Long

Declare Function TransparentBlt Lib "msimg32.dll" _
    (ByVal hdc As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, _
    ByVal crTransparent As Long) _
As Boolean


Public Declare Function Ellipse Lib "gdi32.dll" _
   (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
   ByVal X2 As Long, ByVal Y2 As Long) As Long


Declare Function Rectangle Lib "gdi32.dll" ( _
    ByVal hdc As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long) As Long

