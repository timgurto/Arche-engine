Attribute VB_Name = "mdlBitBlt"
'***EXTERNAL MODULE***
'Bit Block Transfer Functions
'KPD Team
'*********************


Option Explicit

'Declare Function BitBlt Lib "gdi32" _
'    (ByVal hDestDC As Long, _
'    ByVal x As Long, _
'    ByVal y As Long, _
'    ByVal nWidth As Long, _
'    ByVal nHeight As Long, _
'    ByVal hSrcDC As Long, _
'    ByVal xSrc As Long, _
'    ByVal ySrc As Long, _
'    ByVal dwRop As Long) _
'    As Long

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


'***EXTERNAL MODULE SUPPLEMENT***
'SimpleTransparentBlt
'Added by TCT Software
'********************************

'Public Sub SimpleTransparentBlt(hdc As Long, Destination As recCoords, Dimensions As recCoords, hSrcDC As Long, Source As recCoords, Optional crTransparent As Long = &HFF0000)
'    Dim TransparentBltResult As Boolean
'    Do
'        TransparentBltResult = TransparentBlt(hdc, Destination.x, Destination.y, Dimensions.x, Dimensions.y, hSrcDC, Source.x, Source.y, Dimensions.x, Dimensions.y, crTransparent)
'    Loop While Not TransparentBltResult
'End Sub


