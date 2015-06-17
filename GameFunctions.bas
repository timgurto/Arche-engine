Attribute VB_Name = "mdlGameFunctions"
Option Explicit

Public Sub swapUnits(a, b)
Dim temp As typUnit
temp = unit(a)
unit(a) = unit(b)
unit(b) = temp
End Sub

Public Sub deleteUnit(n As Integer)
swapUnits n, activeUnits - 1

End Sub

Public Sub deleteUnits()
Dim i As Integer
i = 0
While i < activeUnits
'For i = 0 To activeUnits - 1
   If unit(i).selected Then
      deleteUnit (i)
      activeUnits = activeUnits - 1
      i = i - 1
   End If
'Next i
i = i + 1
Wend
End Sub

Public Function findPath(n As Integer) As typCoords
Dim i As Integer

'move horizontally
If Abs(unit(n).location.X - unit(n).target.X) >= unitType(unit(n).type).speed Then
   If unit(n).location.X < unit(n).target.X Then
      unit(n).direction = dirR
      findPath = moveRight(unitType(unit(n).type).speed)
   Else
      unit(n).direction = dirL
      findPath = moveLeft(unitType(unit(n).type).speed)
   End If

'move vertically
ElseIf Abs(unit(n).location.Y - unit(n).target.Y) >= unitType(unit(n).type).speed Then
   If unit(n).location.Y < unit(n).target.Y Then
      unit(n).direction = dirD
      findPath = moveDown(unitType(unit(n).type).speed)
   Else
      unit(n).direction = dirU
      findPath = moveUp(unitType(unit(n).type).speed)
   End If
End If

'***COLLISION CHECKS***

'Map edges
If Not collision(addCoords(unit(n).location, findPath), makeCoords(1, 1), makeCoords(0, 0), muxCoords(gameMap.dimensions, TERRAIN_TILE_SIZE)) Then
   If Not KEEP_WALKING_ON_COLLISION Then unit(n).freezeFrame = True 'unit(n).frame = unit(n).frame - 1
   findPath = makeCoords(0, 0)
   Exit Function
End If

'Units
For i = 0 To activeUnits - 1
   If i <> n Then
      If collision(addCoords(screenCoords(unit(n)), findPath), unitType(unit(n).type).dimensions, screenCoords(unit(i)), unitType(unit(i).type).dimensions) Then
         'unit(n).frame = 0
         If Not KEEP_WALKING_ON_COLLISION Then unit(n).freezeFrame = True 'unit(n).frame = unit(n).frame - 1
         findPath = makeCoords(0, 0)
         Exit Function
      End If
   End If
Next i

If Not (findPath.X = 0 And findPath.Y = 0) Then 'if a path is found, and thus if the unit will move
   needReExplore = True
End If

'**********************
End Function

Public Function exploreMap()
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim mid As typCoords 'the middle of each tile
Dim disX As Long 'the sub-squared X half of the distance equation
Dim loc As typCoords
Dim uT As typUnitType

For i = 1 To gameMap.dimensions.X 'set fog initially
   For j = 1 To gameMap.dimensions.Y
      gameMap.fog(i, j) = True
   Next j
Next i

'distance = Sqr((a.X - b.X) ^ 2 + (a.Y - b.Y) ^ 2)
For n = 0 To activeUnits - 1
   loc = unit(n).location
   uT = unitType(unit(n).type)
   For i = 1 To gameMap.dimensions.X
      mid.X = (i - 0.5) * TERRAIN_TILE_SIZE
      disX = (mid.X - loc.X) ^ 2
      For j = 1 To gameMap.dimensions.Y
         mid.Y = (j - 0.5) * TERRAIN_TILE_SIZE
         If Sqr(disX + (mid.Y - loc.Y) ^ 2) <= uT.lineOfSight Then
            gameMap.explored(i, j) = True
            gameMap.fog(i, j) = False
         End If
      Next j
   Next i
Next n
End Function











