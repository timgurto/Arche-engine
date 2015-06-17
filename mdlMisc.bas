Attribute VB_Name = "mdlMisc"
Option Explicit

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

Public Function collision(loc1 As typCoords, dim1 As typCoords, loc2 As typCoords, dim2 As typCoords)
collision = _
((loc1.x <= loc2.x + dim2.x) And _
(loc2.x <= loc1.x + dim1.x)) And _
((loc1.y <= loc2.y + dim2.y) And _
(loc2.y <= loc1.y + dim1.y))

End Function

'Public Function collidesWithUnit(loc As typCoords, dimensions As typCoords)

Public Function screenCoords(dudeInQuestion As typUnit) As typCoords
screenCoords = makeCoords(dudeInQuestion.location.x - 0.5 * unitType(dudeInQuestion.type).dimensions.x, dudeInQuestion.location.y - 0.875 * unitType(dudeInQuestion.type).dimensions.y)
End Function

Public Function findPath(n As Integer) As typCoords
Dim i As Integer
i = 0

'move horizontally
If Abs(unit(n).location.x - unit(n).target.x) >= unitType(unit(n).type).speed Then
   If unit(n).location.x < unit(n).target.x Then
      unit(n).direction = dirR
      findPath = moveRight(unitType(unit(n).type).speed)
   Else
      unit(n).direction = dirL
      findPath = moveLeft(unitType(unit(n).type).speed)
   End If

'move vertically
ElseIf Abs(unit(n).location.y - unit(n).target.y) >= unitType(unit(n).type).speed Then
   If unit(n).location.y < unit(n).target.y Then
      unit(n).direction = dirD
      findPath = moveDown(unitType(unit(n).type).speed)
   Else
      unit(n).direction = dirU
      findPath = moveUp(unitType(unit(n).type).speed)
   End If
End If

'other checks

For i = 0 To activeUnits - 1
   If i <> n Then
      If collision(addCoords(screenCoords(unit(n)), findPath), unitType(unit(n).type).dimensions, screenCoords(unit(i)), unitType(unit(i).type).dimensions) Then
         'unit(n).frame = 0
         If Not keepWalkingOnCollision Then unit(n).freezeFrame = True 'unit(n).frame = unit(n).frame - 1
         findPath = makeCoords(0, 0)
         
      End If
   End If
Next i
      

End Function

Public Function findPathA(ByRef movingUnit As typUnit) As typCoords 'broken; when working, makes for shakey diagonals
Dim u, d, l, r As Double
Dim min As Double

u = distance(addCoords(movingUnit.location, moveUp(unitType(movingUnit.type).speed)), movingUnit.target)
d = distance(addCoords(movingUnit.location, moveDown(unitType(movingUnit.type).speed)), movingUnit.target)
l = distance(addCoords(movingUnit.location, moveLeft(unitType(movingUnit.type).speed)), movingUnit.target)
r = distance(addCoords(movingUnit.location, moveRight(unitType(movingUnit.type).speed)), movingUnit.target)

min = u
movingUnit.direction = dirU
findPathA = moveUp(unitType(movingUnit.type).speed)
If d < min Then
   min = d
   findPathA = moveDown(unitType(movingUnit.type).speed)
   movingUnit.direction = dirD
End If
If l < min Then
   min = l
   findPathA = moveLeft(unitType(movingUnit.type).speed)
   movingUnit.direction = dirL
End If
If r < min Then
   min = r
   findPathA = moveRight(unitType(movingUnit.type).speed)
   movingUnit.direction = dirR
End If

End Function
