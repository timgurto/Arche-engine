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

Public Function findPath(ByRef movingUnit As typUnit) As typCoords

'move horizontally
If Abs(movingUnit.location.x - movingUnit.target.x) >= unitType(movingUnit.type).speed Then
   If movingUnit.location.x < movingUnit.target.x Then
      movingUnit.direction = dirR
      findPath = moveRight(unitType(movingUnit.type).speed)
   Else
      movingUnit.direction = dirL
      findPath = moveLeft(unitType(movingUnit.type).speed)
   End If

'move vertically
ElseIf Abs(movingUnit.location.y - movingUnit.target.y) >= unitType(movingUnit.type).speed Then
   If movingUnit.location.y < movingUnit.target.y Then
      movingUnit.direction = dirD
      findPath = moveDown(unitType(movingUnit.type).speed)
   Else
      movingUnit.direction = dirU
      findPath = moveUp(unitType(movingUnit.type).speed)
   End If
End If



End Function

Public Function findPathA(ByRef movingUnit As typUnit) As typCoords
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
