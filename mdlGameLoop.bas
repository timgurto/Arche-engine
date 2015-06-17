Attribute VB_Name = "mdlGameLoop"
Option Explicit

Public Sub runGameLoop()
Dim i As Integer
Dim j As Integer
Dim d As Long
Dim minDistance As Integer
Dim tar As typUnit
Dim c As typcoords

refreshCount = refreshCount + 1
If refreshCount = REFRESHES_PER_FRAME Then
   For i = 0 To activeUnits - 1
      refreshCount = 0
      If Not unit(i).freezeFrame Then
         unit(i).frame = unit(i).frame + 1
         If unit(i).frame >= unitType(unit(i).type).frames Then unit(i).frame = 0
      Else
         unit(i).freezeFrame = False
      End If
   Next i
End If

terrainFrameTimer = terrainFrameTimer - 20
If terrainFrameTimer <= 0 Then
   terrainFrameTimer = TERRAIN_FRAME_LENGTH
   For j = 0 To activeTerrains - 1
      If terrain(j).frames > 0 Then
         increment terrain(j).frame
         If terrain(j).frame = terrain(j).frames Then terrain(j).frame = 0
      End If
   Next j
End If

For i = 0 To activeUnits - 1
   If unit(i).targetUnit > -1 Then
      unit(i).target = unit(unit(i).targetUnit).location
   'ElseIf unit(i).targetBuilding > -1 Then
      'unit(i).target = building(unit(i).targetBuilding).location
   End If
   
   'Update unit/building coords
   If unit(i).moving Then
      unit(i).location = addCoords(unit(i).location, findPath(i))
      c = getTile(unit(i).location)
      If unitType(unit(i).type).heavy And terrain(gameMap.terrain(c.x, c.y)).collapsesTo >= 0 Then
         deleteUnit i, False
         gameMap.terrain(c.x, c.y) = terrain(gameMap.terrain(c.x, c.y)).collapsesTo
      End If
   Else
      unit(i).frame = 0
   End If
   
   unit(i).moving = distance(unit(i).location, unit(i).target) > unitType(unit(i).type).speed
   If unit(i).targetUnit > -1 Then
      unit(i).moving = IIf( _
         unitType(unit(i).type).range > 0, _
         distance(unit(i).location, unit(i).target) > unitType(unit(i).type).range * RANGED_UNIT, _
         distance(unit(i).location, unit(i).target) > unitType(unit(i).type).speed _
      )
   End If
   

   unit(i).attackTimer = unit(i).attackTimer + 20 'see gameLoop()

   'Autoattacking
   If AUTO_ATTACKING Then
      If Not unit(i).moving Then
         If unitType(unit(i).type).healing = 0 And (unit(i).targetUnit = -1 Or (Not unit(i).moving)) Then 'if not targetting anyone
            minDistance = AUTO_ATTACK_RANGE + 5 'safe margin
            For j = 0 To activeUnits - 1
               If unit(i).player <> unit(j).player Then
                  d = distance(unit(i).location, unit(j).location)
                  If d < AUTO_ATTACK_RANGE Then
                     If d < minDistance Then
                        unit(i).targetUnit = j
                        unit(i).combatMode = False
                        If DEBUG_MODE Then frmGame.lblTargetUnit = j
                        minDistance = d
                     End If
                  End If
               End If
            Next j
         End If
         End If
   End If

   'Attacking
   If unit(i).attackTimer >= unitType(unit(i).type).attackSpeed Then
      unit(i).attackTimer = 0
      
      'attack if in range
      If unit(i).targetUnit > -1 Then
         tar = unit(unit(i).targetUnit)
         If tar.player <> unit(i).player Then
            If unitType(unit(i).type).attack > 0 Then
               If (unitType(unit(i).type).range > 0 Or (findPath(i).x = 0 And findPath(i).y = 0)) And distance(unit(i).target, unit(i).location) < max(unitSize(i, unit(i).targetUnit), unitType(unit(i).type).range * RANGED_UNIT) Then
                  'If (unit(unit(i).targetUnit).targetUnit = -1) Or unitType(unit(i).type).taunting Then unit(unit(i).targetUnit).targetUnit = i
                  If unitType(unit(i).type).attackSound > -1 Then sound (unitType(unit(i).type).attackSound)
                  unit(i).combatMode = True
                  unit(unit(i).targetUnit).health = tar.health - max(unitType(unit(i).type).attack - unitType(tar.type).armor, 0)
                  If unit(unit(i).targetUnit).health <= 0 Then killUnit (unit(i).targetUnit)
                  frmGame.updateStats
               End If
            End If
         Else
            If unitType(unit(i).type).healing > 0 Then
               If (unitType(unit(i).type).range > 0 Or (findPath(i).x = 0 And findPath(i).y = 0)) And distance(unit(i).target, unit(i).location) < max(unitSize(i, unit(i).targetUnit), unitType(unit(i).type).range * RANGED_UNIT) Then
                  unit(i).combatMode = True
                  unit(unit(i).targetUnit).health = min(unit(unit(i).targetUnit).health + unitType(unit(i).type).healing, unitType(unit(unit(i).targetUnit).type).health)
                  frmGame.updateStats
               End If
            End If
         End If
      End If
   
   ElseIf unit(i).attackTimer >= unitType(unit(i).type).attackSpeed / 2 Then
      unit(i).combatMode = False
   End If

Next i

If DEBUG_MODE Then frmGame.shpExplore.BackColor = vbGreen
If needReExplore Then
   exploreMap
   If DEBUG_MODE Then frmGame.shpExplore.BackColor = vbRed
End If

drawEverything

End Sub

Public Sub drawEverything()
Dim i As Integer

frmGame.picGame.Cls

'terrain
drawMap gameMap

'corpses
For i = 0 To activeCorpses - 1
   drawCorpse corpse(i)
Next i

'player marks, units
For i = 0 To activeUnits - 1
   'if visible
   If gameMap.explored(getUnitTile(i).x, getUnitTile(i).y) Then
      drawPlayerMark unit(i)
      If unit(i).selected Then drawSelection unit(i)
      drawUnit unit(i)
      drawUnit unit(i)
   End If
Next i

For i = 0 To activeUnits - 1
   If unit(i).selected And (unit(i).moving Or unit(i).targetUnit <> -1) Then drawTarget unit(i)
Next i

For i = 0 To activeUnits - 1
   If unit(i).selected Then
      If gameMap.explored(getUnitTile(i).x, getUnitTile(i).y) Then
         drawHealthBar unit(i)
      End If
   End If
Next i
         
drawSelectionRectangle

frmGame.picGame.Refresh
End Sub

Public Sub gameLoop()

Const tickDifference As Long = 20
Dim lastTick As Long
Dim currentTick As Long

Do
   currentTick = GetTickCount
   If currentTick - lastTick > tickDifference Then
      
      runGameLoop
         
      lastTick = currentTick
   End If
   DoEvents
   
Loop



End Sub
