Attribute VB_Name = "mdlGameFunctions"
Option Explicit

Public Function findUnit(target As typcoords) As Integer
Dim i As Integer
findUnit = -1
For i = activeUnits - 1 To 0 Step -1
   If collision(screenCoords(unit(i)), unitType(unit(i).type).dimensions, target, makeCoords(1, 1)) Then
      findUnit = i
      i = -1
   End If
Next i
End Function

Public Sub swapUnits(a As Integer, b As Integer)
Dim temp As typUnit
Dim entA As Integer, entB As Integer
Dim i As Integer

'entA = unit(a).entity
'entB = unit(b).entity

temp = unit(a)
unit(a) = unit(b)
unit(b) = temp

For i = 0 To activeUnits - 1
   If unit(i).targetUnit = a Then
      unit(i).targetUnit = b
   ElseIf unit(i).targetUnit = b Then
      unit(i).targetUnit = a
   End If
Next i

'entity(entA).index=
'entity(entB).index=
'unit(a).entity = entA
'unit(b).entity = entB

If victoryType = REGICIDE Then
   For i = 0 To activePlayers - 1
      If regicideTarget(i) = a Then
         regicideTarget(i) = b
      ElseIf regicideTarget(i) = b Then
         regicideTarget(i) = a
      End If
   Next i
End If
End Sub

Public Sub victory(p As Byte)
Dim message As String
message = player(p).name & " won!"
MsgBox (message)
Call ChangeRes(screenResolution.x, screenResolution.y)
End
End Sub

Public Sub deleteCorpse(n As Integer)
Dim temp As typCorpse
temp = corpse(n)
corpse(n) = corpse(activeCorpses - 1)
corpse(activeCorpses - 1) = temp
activeCorpses = activeCorpses - 1
End Sub

Public Sub killUnit(n As Integer)
deleteUnit (n)
End Sub

Public Sub deleteUnit(n As Integer)
Dim i As Integer

unit(n).selected = False

'Play a death sound
If unitType(unit(n).type).deathSound > -1 Then sound (unitType(unit(n).type).deathSound)

'Assess victories
If victoryType = REGICIDE Then
   For i = 0 To activePlayers - 1
      If regicideTarget(i) = n Then
         victory (i)
         i = activePlayers 'civs
      End If
   Next i
End If

'Fix units targetting it
Dim j As Integer
For j = 0 To activeUnits - 1
   If unit(j).targetUnit = n Then
      unit(j).targetUnit = -1
      unit(j).combatMode = False
      unit(j).target = unit(j).location
   End If
Next j

'Add a corpse
If activeCorpses = MAX_CORPSES Then deleteCorpse (0)
corpse(activeCorpses).dimensions = unitType(unit(n).type).dimensions
corpse(activeCorpses).location = unit(n).location
corpse(activeCorpses).type = unitType(unit(n).type).corpse
corpse(activeCorpses).timer = corpseType(corpse(activeCorpses).type).timer
increment activeCorpses

'printEntityList "Before deletion:"

'Remove unit
'For j = unit(n).entity To activeEntities - 2
'   entity(j) = entity(j + 1)
'   unit(entity(j).index).entity = j 'unit?
'Next j
'activeEntities = activeEntities - 1

swapUnits n, activeUnits - 1
For j = 0 To activeUnits - 1
   If unit(j).targetUnit = n Then
      unit(j).targetUnit = activeUnits - 1
      unit(j).combatMode = False
   Else
      If unit(j).targetUnit = activeUnits - 1 Then
         unit(j).targetUnit = n
         unit(j).combatMode = False
      End If
   End If
Next j
activeUnits = activeUnits - 1

'Might be impractical:
sortUnits

frmGame.updateStats

'printEntityList "After deletion:"

End Sub

Public Sub deleteUnits()
Dim i As Integer
i = 0
While i < activeUnits
'For i = 0 To activeUnits - 1
   If unit(i).selected And (unit(i).player = you Or DEBUG_MODE) Then
      deleteUnit (i)
      i = i - 1
   End If
'Next i
i = i + 1
Wend
End Sub

Public Function findPath(n As Integer) As typcoords
Dim i As Integer
Dim x As Integer, y As Integer
Dim c As typcoords

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

'***COLLISION CHECKS***

If Not validLocation(addCoords(collisionLoc(n), findPath), unitType(unit(n).type).collisionDim, n) Then
   If Not KEEP_WALKING_ON_COLLISION Then unit(n).freezeFrame = True 'unit(n).frame = unit(n).frame - 1
   findPath = makeCoords(0, 0)
   Exit Function
Else 'if a path is found, and thus if the unit will move
   unit(n).exploring = True
   needReExplore = True
End If

'**********************



End Function

Public Function exploreMap()
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim mid As typcoords 'the middle of each tile
Dim disX As Long 'the sub-squared X half of the distance equation
Dim loc As typcoords
Dim uT As typUnitType

'If FOG_OF_WAR Then

   For i = 0 To gameMap.dimensions.x - 1 'set fog initially
      For j = 0 To gameMap.dimensions.y - 1
         gameMap.fog(i, j) = True
      Next j
   Next i
   
   'distance = Sqr((a.X - b.X) ^ 2 + (a.Y - b.Y) ^ 2)
   For n = 0 To activeUnits - 1
      If unit(n).player = you And unit(n).exploring Then
         loc = unit(n).location
         uT = unitType(unit(n).type)
         For i = 0 To gameMap.dimensions.x - 1
            mid.x = (i - 0.5) * TERRAIN_TILE_SIZE
            disX = (mid.x - loc.x) ^ 2
            For j = 0 To gameMap.dimensions.y - 1
               mid.y = (j - 0.5) * TERRAIN_TILE_SIZE
               If Sqr(disX + (mid.y - loc.y) ^ 2) <= uT.lineOfSight Then
                  gameMap.explored(i, j) = True
                  gameMap.fog(i, j) = False
               End If
            Next j
         Next i
         unit(n).exploring = False
      End If
   Next n
   
'Else

   For n = 0 To activeUnits - 1
      If unit(n).player = you And unit(n).exploring Then
         For i = 0 To gameMap.dimensions.x - 1
            For j = 0 To gameMap.dimensions.y - 1
               If Not gameMap.explored(i, j) Then
                  If distance(unit(n).location, makeCoords((i - 0.5) * TERRAIN_TILE_SIZE, (j - 0.5) * TERRAIN_TILE_SIZE)) <= unitType(unit(n).type).lineOfSight Then
                     gameMap.explored(i, j) = True
                  End If
               End If
            Next j
         Next i
         unit(n).exploring = False
      End If
   Next n
         
'End If

needReExplore = False
End Function


Public Function getSelected() As Integer
getSelected = -1
Dim i As Integer
Dim found As Boolean 'whether a unit has been found yet
found = False
For i = 0 To activeUnits - 1
   If unit(i).selected Then
      getSelected = i
      i = activeUnits
   End If
Next i
End Function

Public Function getTile(c As typcoords) As typcoords
Dim x As Integer, y As Integer
x = Int(c.x / TERRAIN_TILE_SIZE) + 1
y = Int(c.y / TERRAIN_TILE_SIZE) + 1
getTile = makeCoords(x, y)
End Function

Public Function getUnitTile(n As Integer) As typcoords
Dim x As Integer, y As Integer
Dim u As typUnit
u = unit(n)
x = Int(u.location.x / TERRAIN_TILE_SIZE) + 1
y = Int(u.location.y / TERRAIN_TILE_SIZE) + 1
getUnitTile = makeCoords(x, y)
End Function

'Public Sub sortUnits()
''printEntityList "Before sort"
''insertion
'Dim value As typUnit
'Dim valueIndex As Integer
'Dim i As Integer
'Dim j As Integer
'For i = 1 To activeUnits - 1
'   value = unit(i)
'   valueIndex = i
'   j = i - 1
'   Do While j >= 0
'      If unit(j).location.y > value.location.y Then
'         unit(j + 1) = unit(j)
'         j = j - 1
'      Else
'         Exit Do
'      End If
'   Loop
'   unit(j + 1) = value
'Next i
'
'End Sub

Public Sub sortUnits()
'selection
Dim i As Integer
Dim j As Integer
Dim min As Integer

For i = 0 To activeUnits - 1
   min = i
   For j = i + 1 To activeUnits - 1
      If unit(j).location.y < unit(min).location.y Then min = j
   Next j
   If i <> min Then Call swapUnits(i, min)
Next i

'If DEBUG_MODE Then
'   For i = 0 To activeUnits - 1
'      Debug.Print unit(i).location.y
'   Next i
'   Debug.Print ""
'End If

End Sub

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

Public Sub reSortUnits(n As Integer, displacement As Integer, val As Integer)
Dim temp As typUnit
Select Case displacement
Case Is > 0
   'Debug.Print "Displacement > 0, n=" & n
   While n < activeUnits - 1
      If val > unit(n + 1).location.y Then
         swapUnits n, n + 1
         n = n + 1
      Else
         Exit Sub
      End If
   Wend
Case Is < 0
   'Debug.Print "Displacement < 0, n=" & n
   While n > 0
      If (val < unit(n - 1).location.y) Then
         swapUnits n, n - 1
         n = n - 1
      Else
         Exit Sub
      End If
   Wend
End Select

End Sub

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
'
Public Function collisionLoc(n As Integer) As typcoords
Dim u As typUnit
Dim t As typUnitType

u = unit(n)
t = unitType(u.type)

collisionLoc = addCoords(screenCoords(u), t.collisionLoc)
End Function

Public Function validLocation(location As typcoords, dimensions As typcoords, Optional unitIndex As Integer = -1) As Boolean
Dim c As typcoords
Dim i As Integer

validLocation = True

'Map edges
If location.x < 0 Then
   validLocation = False
   Exit Function
End If
If location.y < 0 Then
   validLocation = False
   Exit Function
End If
If location.x + dimensions.x > (gameMap.dimensions.x - 1) * TERRAIN_TILE_SIZE Then
   validLocation = False
   Exit Function
End If
If location.y + dimensions.y > (gameMap.dimensions.y - 1) * TERRAIN_TILE_SIZE Then
   validLocation = False
   Exit Function
End If

c = getTile(location)
If terrain(gameMap.terrain(c.x, c.y)).impassable Then
   validLocation = False
   Exit Function
End If
c = getTile(addCoords(location, dimensions))
If terrain(gameMap.terrain(c.x, c.y)).impassable Then
   validLocation = False
   Exit Function
End If

'Units
For i = 0 To activeUnits - 1
   If i <> unitIndex Then
      If collision( _
      location, _
      dimensions, _
      collisionLoc(i), _
      unitType(unit(i).type).collisionDim _
      ) Then
         validLocation = False
         Exit Function
      End If
   End If
Next i

End Function

