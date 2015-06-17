Attribute VB_Name = "mdlMisc"
Option Explicit

Public Function makeDC(file As String)
makeDC = CreateCompatibleDC(0)
makeDC = LoadGraphicDC(gamePath & " Data\Images\" & file)
End Function

Public Sub sound(n As Integer)
Dim x As Long
If Not DEBUG_MODE Then x = sndPlaySound(gamePath & " Data\Sounds\s" & n & ".wav", sndAsync)
End Sub

Public Function increment(ByRef n As Variant)
n = n + 1
End Function

Public Function makeCoords(x As Integer, y As Integer) As typcoords
makeCoords.x = x
makeCoords.y = y
End Function

Public Function min(x As Variant, y As Variant)
min = IIf(x < y, x, y)
End Function

Public Function max(x As Variant, y As Variant)
max = IIf(x > y, x, y)
End Function

Public Function distance(a As typcoords, b As typcoords) As Double
distance = Sqr((a.x - b.x) ^ 2 + (a.y - b.y) ^ 2)
End Function

Public Function addCoords(a As typcoords, b As typcoords) As typcoords
addCoords.x = a.x + b.x
addCoords.y = a.y + b.y
End Function

Public Function subCoords(a As typcoords, b As typcoords) As typcoords
subCoords.x = a.x - b.x
subCoords.y = a.y - b.y
End Function

Public Function muxCoords(a As typcoords, n As Integer) As typcoords
muxCoords.x = a.x * n
muxCoords.y = a.y * n
End Function

Public Function collision(loc1 As typcoords, dim1 As typcoords, loc2 As typcoords, dim2 As typcoords)
collision = _
((loc1.x <= loc2.x + dim2.x) And _
(loc2.x <= loc1.x + dim1.x)) And _
((loc1.y <= loc2.y + dim2.y) And _
(loc2.y <= loc1.y + dim1.y))

End Function

Public Function str2Bool(str As String) As Boolean
Dim s As String
s = UCase(str)
Select Case s
Case "T"
    str2Bool = True
Case "TRUE"
    str2Bool = True
Case "F"
    str2Bool = False
Case "FALSE"
    str2Bool = False
Case Else
    str2Bool = False
    If DEBUG_MODE Then MsgBox "Attempting to read '" & str & "' from a file as a boolean.  Returning False by default."
End Select
End Function





