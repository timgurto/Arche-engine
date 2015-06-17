Attribute VB_Name = "mdlMisc"
Option Explicit

Public Function makeCoords(X As Integer, Y As Integer) As typCoords
makeCoords.X = X
makeCoords.Y = Y
End Function

Public Function distance(a As typCoords, b As typCoords) As Double
distance = Sqr((a.X - b.X) ^ 2 + (a.Y - b.Y) ^ 2)
End Function

Public Function moveUp(n As Integer) As typCoords
moveUp.X = 0
moveUp.Y = -1 * n
End Function

Public Function moveDown(n As Integer) As typCoords
moveDown.X = 0
moveDown.Y = 1 * n
End Function

Public Function moveLeft(n As Integer) As typCoords
moveLeft.X = -1 * n
moveLeft.Y = 0
End Function

Public Function moveRight(n As Integer) As typCoords
moveRight.X = 1 * n
moveRight.Y = 0
End Function

Public Function addCoords(a As typCoords, b As typCoords) As typCoords
addCoords.X = a.X + b.X
addCoords.Y = a.Y + b.Y
End Function

Public Function subCoords(a As typCoords, b As typCoords) As typCoords
subCoords.X = a.X - b.X
subCoords.Y = a.Y - b.Y
End Function

Public Function muxCoords(a As typCoords, n As Integer) As typCoords
muxCoords.X = a.X * n
muxCoords.Y = a.Y * n
End Function

Public Function collision(loc1 As typCoords, dim1 As typCoords, loc2 As typCoords, dim2 As typCoords)
collision = _
((loc1.X <= loc2.X + dim2.X) And _
(loc2.X <= loc1.X + dim1.X)) And _
((loc1.Y <= loc2.Y + dim2.Y) And _
(loc2.Y <= loc1.Y + dim1.Y))

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
screenCoords = makeCoords(dudeInQuestion.location.X - 0.5 * unitType(dudeInQuestion.type).dimensions.X, dudeInQuestion.location.Y - 0.875 * unitType(dudeInQuestion.type).dimensions.Y)
End Function
