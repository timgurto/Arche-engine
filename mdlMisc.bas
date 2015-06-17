Attribute VB_Name = "mdlMisc"
Option Explicit

Public Function makeDC(file As String)
      makeDC = CreateCompatibleDC(0)
      makeDC = LoadGraphicDC(App.Path & "\Images\" & file)
End Function

Public Function increment(ByRef n As Variant)
n = n + 1
End Function

Public Function makeCoords(x As Integer, y As Integer) As typcoords
makeCoords.x = x
makeCoords.y = y
End Function

Public Function unitSize(i As Integer, j As Integer) As Integer
Dim u As typUnit, v As typUnit
Dim s As typUnitType, t As typUnitType
Dim tar As typcoords
u = unit(i)
v = unit(j)
t = unitType(u.type)
s = unitType(v.type)
'tar = muxCoords(unitType(unit(u.targetUnit).type).dimensions, -1)
'unitSize = distance(t.dimensions, tar) / 2
unitSize = 1.5 * max(max(t.dimensions.x, t.dimensions.y), max(s.dimensions.x, s.dimensions.y))
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

Public Function moveUp(n As Integer) As typcoords
moveUp.x = 0
moveUp.y = -1 * n
End Function

Public Function moveDown(n As Integer) As typcoords
moveDown.x = 0
moveDown.y = 1 * n
End Function

Public Function moveLeft(n As Integer) As typcoords
moveLeft.x = -1 * n
moveLeft.y = 0
End Function

Public Function moveRight(n As Integer) As typcoords
moveRight.x = 1 * n
moveRight.y = 0
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

Public Function pointCollidesWithUnit(loc As typcoords, ByRef u As Integer) As Boolean
pointCollidesWithUnit = False
For u = 0 To activeUnits - 1
   If collision(loc, makeCoords(1, 1), screenCoords(unit(u)), unitType(unit(u).type).dimensions) Then
      pointCollidesWithUnit = True
      Exit For
   End If
Next u
End Function

Public Function screenCoords(dudeInQuestion As typUnit) As typcoords
screenCoords = makeCoords(dudeInQuestion.location.x - 0.5 * unitType(dudeInQuestion.type).dimensions.x, dudeInQuestion.location.y - 0.875 * unitType(dudeInQuestion.type).dimensions.y)
End Function
