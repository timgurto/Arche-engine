Attribute VB_Name = "mdlMisc"
Option Explicit

Public Function increment(ByRef n As Variant)
n = n + 1
End Function

Public Function makeCoords(x As Integer, y As Integer) As typCoords
makeCoords.x = x
makeCoords.y = y
End Function

Public Function distance(a As typCoords, b As typCoords) As Double
distance = Sqr((a.x - b.x) ^ 2 + (a.y - b.y) ^ 2)
End Function

Public Function moveUp(n As Integer) As typCoords
moveUp.x = 0
moveUp.y = -1 * n
End Function

Public Function moveDown(n As Integer) As typCoords
moveDown.x = 0
moveDown.y = 1 * n
End Function

Public Function moveLeft(n As Integer) As typCoords
moveLeft.x = -1 * n
moveLeft.y = 0
End Function

Public Function moveRight(n As Integer) As typCoords
moveRight.x = 1 * n
moveRight.y = 0
End Function

Public Function addCoords(a As typCoords, b As typCoords) As typCoords
addCoords.x = a.x + b.x
addCoords.y = a.y + b.y
End Function

Public Function subCoords(a As typCoords, b As typCoords) As typCoords
subCoords.x = a.x - b.x
subCoords.y = a.y - b.y
End Function

Public Function muxCoords(a As typCoords, n As Integer) As typCoords
muxCoords.x = a.x * n
muxCoords.y = a.y * n
End Function

Public Function collision(loc1 As typCoords, dim1 As typCoords, loc2 As typCoords, dim2 As typCoords)
collision = _
((loc1.x <= loc2.x + dim2.x) And _
(loc2.x <= loc1.x + dim1.x)) And _
((loc1.y <= loc2.y + dim2.y) And _
(loc2.y <= loc1.y + dim1.y))

End Function

Public Function pointCollidesWithUnit(loc As typCoords, ByRef u As Integer) As Boolean
pointCollidesWithUnit = False
For u = 0 To activeUnits - 1
   If collision(loc, makeCoords(1, 1), screenCoords(unit(u)), unitType(unit(u).type).dimensions) Then
      pointCollidesWithUnit = True
      Exit For
   End If
Next u
End Function

Public Function screenCoords(dudeInQuestion As typUnit) As typCoords
screenCoords = makeCoords(dudeInQuestion.location.x - 0.5 * unitType(dudeInQuestion.type).dimensions.x, dudeInQuestion.location.y - 0.875 * unitType(dudeInQuestion.type).dimensions.y)
End Function
