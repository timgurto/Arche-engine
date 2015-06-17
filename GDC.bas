Attribute VB_Name = "GDC"
'***EXTERNAL MODULE***
'GDC Functions
'Tim Miron
'*********************


Option Explicit

'The following API calls are for:

'blitting
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'code timer
Declare Function GetTickCount Lib "kernel32" () As Long

'creating buffers / loading sprites
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'loading sprites
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'cleanup
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'end of copy-paste here...



Public Function LoadGraphicDC(sFileName As String) As Long

'cheap error handling
On Error Resume Next

'temp variable to hold our DC address
Dim LoadGraphicDCTEMP As Long

'create the DC address compatible with
'the DC of the screen
LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))

'load the graphic file into the DC...
SelectObject LoadGraphicDCTEMP, LoadPicture(sFileName)

'return the address of the file
LoadGraphicDC = LoadGraphicDCTEMP

'******
'DeleteDC LoadGraphicDCTEMP
'DeleteObject LoadGraphicDCTEMP
'******
End Function

'end of copy-paste here...
