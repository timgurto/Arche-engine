Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()
Dim i As Integer
Dim j As Integer
Dim x As Integer, y As Integer
Dim s As String
Dim temp1 As String
Dim temp2 As String
Dim temp3 As String
Dim temp4 As String

Open gamePath & " Data\Data\Game.txt" For Input As #1
Input #1, temp1, s
HEALTH_BAR_CIV_COLOR = str2Bool(temp1)
Input #1, temp1, s
SELECTION_RECTANGLE_SHADOW = str2Bool(temp1)
Input #1, temp1, s
KEEP_WALKING_ON_COLLISION = str2Bool(temp1)
Input #1, temp1, s
SHOW_SELECTED_TARGETS = str2Bool(temp1)
Input #1, temp1, s
FOG_OF_WAR = str2Bool(temp1)
Input #1, temp1, s
ENEMIES_SELECTABLE = str2Bool(temp1)
Input #1, temp1, s
ENEMIES_HAVE_ELLIPSES = str2Bool(temp1)
Input #1, temp1, s
YOU_HAVE_ELLIPSES = str2Bool(temp1)
Input #1, temp1, s
SHOW_UNUSED_STATS = str2Bool(temp1)
Input #1, temp1, s
SPECIAL_PERCENT = str2Bool(temp1)
Input #1, temp1, s
AUTO_ATTACKING = str2Bool(temp1)
Input #1, TERRAIN_TILE_SIZE, s, _
SELECTION_ELLIPSE_WIDTH, s, _
HEALTH_BAR_WIDTH, s, _
HEALTH_BAR_COLOR, s, _
SPECIAL_NAME, s, _
PORTRAIT_WIDTH, s, _
PORTRAIT_HEIGHT, s, _
AUTO_ATTACK_RANGE, s, _
RANGED_UNIT, s, _
TERRAIN_FRAME_LENGTH
Close #1

Open gamePath & " Data\Data\Data.txt" For Input As #1

Debug.Print "Loading factions"
Input #1, activeCivs, s, s
For i = 0 To activeCivs - 1
   With civ(i)
      Input #1, .name, .color
   End With
Next i
Input #1, s

Debug.Print "Loading players"
Input #1, activePlayers, s, s
For i = 0 To activePlayers - 1
   With player(i)
      Input #1, .name, .civ, .population
   End With
Next i
Input #1, s

Debug.Print "Loading terrain"
Input #1, activeTerrains, s, s
For i = 0 To activeTerrains - 1
   With terrain(i)
      Input #1, .name, temp1, .frames, .frame
      .impassable = str2Bool(temp1)
      .dc = makeDC("t" & i & ".bmp")
   End With
Next i

Input #1, s, s

Debug.Print "Loading map"
Input #1, gameMap.displacement.x, s
Input #1, gameMap.displacement.y, s
Input #1, s

Input #1, s
Input #1, gameMap.dimensions.x, s
Input #1, gameMap.dimensions.y, s
For y = 0 To gameMap.dimensions.y - 1
   For x = 0 To gameMap.dimensions.x - 1
      Input #1, gameMap.terrain(x, y)
   Next x
Next y
Input #1, s

Input #1, s
For y = 0 To gameMap.dimensions.y - 1
   For x = 0 To gameMap.dimensions.x - 1
      Input #1, temp1
      gameMap.explored(x, y) = str2Bool(temp1)
   Next x
Next y
Input #1, s

Debug.Print "Loading units"
Input #1, activeUnitTypes, s, s
For i = 0 To activeUnitTypes - 1
   With unitType(i)
      Input #1, .name, .health, .armor, .attack, .healing, .range, _
      .background, .portraitBackground, .dimensions.x, .dimensions.y, _
      .collisionLoc.x, .collisionLoc.y, .collisionDim.x, .collisionDim.y, _
      .corpse, .selectSound, .attackSound, .deathSound, .speed, _
      .attackSpeed, .frames, .lineOfSight, .taunting
      .dc = makeDC("u" & i & ".bmp")
      .portrait = makeDC("p" & i & ".bmp")
   End With
Next i
Input #1, s

Input #1, activeUnits, s, s
For i = 0 To activeUnits - 1
   With unit(i)
      Input #1, .type, .health, .location.x, .location.y, .targetUnit, _
      .target.x, .target.y, .player, temp1, .frame, .attackTimer, _
      .direction, temp2, temp3, temp4
      .exploring = True
      .moving = str2Bool(temp1)
      .selected = str2Bool(temp2)
      .freezeFrame = str2Bool(temp3)
      .combatMode = str2Bool(temp4)
   End With
Next i
Input #1, s

Debug.Print "Loading victory conditions"
Input #1, victoryType, s
If victoryType = REGICIDE Then
   For i = 0 To activePlayers - 1
      Input #1, regicideTarget(i)
   Next i
End If
If victoryType = CONQUEST Then Input #1, s
Input #1, s

Debug.Print "Loading corpses"
Input #1, activeCorpseTypes, s, s
For i = 0 To activeCorpseTypes - 1
   With corpseType(i)
      Input #1, .timer, .dimensions.x, .dimensions.y, .background
      .dc = makeDC("c" & i & ".bmp")
   End With
Next i
Input #1, s

Input #1, activeCorpses, s, s
For i = 0 To activeCorpses - 1
   With corpse(i)
      Input #1, .type, .location.x, .location.y, .dimensions.x, .dimensions.y, .timer
   End With
Next i
Input #1, s

If SHOW_SELECTED_TARGETS Then
   Input #1, s, s
   With target
      Input #1, .dimensions.x, .dimensions.y, .background
      .dc = makeDC("_target.bmp")
   End With
End If

Close #1

If FOG_OF_WAR Then fogDC = makeDC("_fog.bmp")

sortUnits

terrainFrameTimer = TERRAIN_FRAME_LENGTH
needReExplore = True
selectionRectangle = False
refreshCount = 0
ctrlDown = False
scrollDir = dirN

screenResolution.x = Screen.Width \ Screen.TwipsPerPixelX
screenResolution.y = Screen.Height \ Screen.TwipsPerPixelY

End Sub
