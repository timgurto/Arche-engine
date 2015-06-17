Attribute VB_Name = "mdlGameFunctions"
Option Explicit

Public Function findUnit(target As typcoords) As Integer
Dim i As Integer
findUnit = -1
For i = 0 To activeUnits - 1
   If collision(screenCoords(unit(i)), unitType(unit(i).type).dimensions, target, makeCoords(1, 1)) Then
      findUnit = i
      i = activeUnits
   End If
Next i
End Function

Public Sub swapUnits(a, b)
Dim temp As typUnit
Dim i As Integer
temp = unit(a)
unit(a) = unit(b)
unit(b) = temp
If victoryType = REGICIDE Then
   For i = 1 To 2 'civs
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
message = "Player " & p & " wins!"
MsgBox (message)
Call ChangeRes(1680, 1050)
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
If unitType(unit(n).type).deathSound > -1 Then sound (unitType(unit(n).type).deathSound)
unit(n).selected = False
If victoryType = REGICIDE Then
   For i = 1 To 2 'civs
      If regicideTarget(i) = n Then
         victory (i)
         i = 2 'civs
      End If
   Next i
End If
Dim j As Integer
      For j = 0 To activeUnits - 1
         If unit(j).targetUnit = n Then
            unit(j).targetUnit = -1
            unit(j).combatMode = False
            unit(j).target = unit(j).location
         End If
      Next j
If activeCorpses = MAX_CORPSES Then deleteCorpse (0)
corpse(activeCorpses).dimensions = unitType(unit(n).type).dimensions
corpse(activeCorpses).location = unit(n).location
corpse(activeCorpses).type = unitType(unit(n).type).corpse
corpse(activeCorpses).timer = corpseType(corpse(activeCorpses).type).timer
increment activeCorpses
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

frmGame.updateStats
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

'Map edges
If Not collision(addCoords(unit(n).location, findPath), makeCoords(1, 1), makeCoords(1, 1), subCoords(muxCoords(gameMap.dimensions, TERRAIN_TILE_SIZE), makeCoords(2, 2))) Then
   If Not KEEP_WALKING_ON_COLLISION Then unit(n).freezeFrame = True 'unit(n).frame = unit(n).frame - 1
   findPath = makeCoords(0, 0)
   Exit Function
End If

c = getTile(addCoords(unit(n).location, findPath))
If terrain(gameMap.terrain(c.x, c.y)).impassable Then
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

If Not (findPath.x = 0 And findPath.y = 0) Then 'if a path is found, and thus if the unit will move
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

If FOG_OF_WAR Then

   For i = 1 To gameMap.dimensions.x 'set fog initially
      For j = 1 To gameMap.dimensions.y
         gameMap.fog(i, j) = True
      Next j
   Next i
   
   'distance = Sqr((a.X - b.X) ^ 2 + (a.Y - b.Y) ^ 2)
   For n = 0 To activeUnits - 1
      If unit(n).player = you And unit(n).exploring Then
         loc = unit(n).location
         uT = unitType(unit(n).type)
         For i = 1 To gameMap.dimensions.x
            mid.x = (i - 0.5) * TERRAIN_TILE_SIZE
            disX = (mid.x - loc.x) ^ 2
            For j = 1 To gameMap.dimensions.y
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
   
Else

   For n = 0 To activeUnits - 1
      If unit(n).player = you And unit(n).exploring Then
         For i = 1 To gameMap.dimensions.x
            For j = 1 To gameMap.dimensions.x
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
         
End If

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

Public Sub sound(n As Integer)
Dim x As Long
x = sndPlaySound(App.Path & "\Sounds\s" & n & ".wav", sndAsync)

End Sub
