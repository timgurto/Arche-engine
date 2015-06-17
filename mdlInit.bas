Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()
Dim i As Integer
Dim j As Integer

refreshCount = 0

activeUnits = 3
activeCorpses = 1

ctrlDown = False

scrollDir = dirN

target.dimensions.x = 16
target.dimensions.y = 16
target.dc = CreateCompatibleDC(0)
target.dc = LoadGraphicDC(App.Path & "\Images\target.bmp")
target.background = vbRed

gameMap.displacement = makeCoords(-200, -200)

fogDC = CreateCompatibleDC(0)
fogDC = LoadGraphicDC(App.Path & "\Images\fog.bmp")

terrain(1).dc = CreateCompatibleDC(0)
terrain(1).dc = LoadGraphicDC(App.Path & "\Images\grass.bmp")

terrain(2).dc = CreateCompatibleDC(0)
terrain(2).dc = LoadGraphicDC(App.Path & "\Images\dirt.bmp")

gameMap.dimensions = makeCoords(15, 15)
For i = 1 To 15
   For j = 1 To 15
      gameMap.terrain(i, j) = Int(Rnd * (2) + 1)
      gameMap.explored(i, j) = False
   Next j
Next i

unitType(1).name = "Standard"
unitType(1).health = 100
unitType(1).armor = 2
unitType(1).attack = 30
unitType(1).dc = CreateCompatibleDC(0)
unitType(1).dc = LoadGraphicDC(App.Path & "\Images\Standard.bmp")
unitType(1).portrait = CreateCompatibleDC(0)
unitType(1).portrait = LoadGraphicDC(App.Path & "\Images\StandardPortrait.bmp")
unitType(1).background = vbWhite
unitType(1).portraitBackground = vbWhite
unitType(1).dimensions.x = 48
unitType(1).dimensions.y = 48
unitType(1).speed = 3
unitType(1).attackSpeed = 2500
unitType(1).frames = 4
unitType(1).lineOfSight = 150

unitType(2).name = "Hero"
unitType(2).health = 75
unitType(2).armor = 0
unitType(2).attack = 5
unitType(2).dc = CreateCompatibleDC(0)
unitType(2).dc = LoadGraphicDC(App.Path & "\Images\Hero.bmp")
unitType(2).portrait = CreateCompatibleDC(0)
unitType(2).portrait = LoadGraphicDC(App.Path & "\Images\HeroPortrait.bmp")
unitType(2).background = vbGreen
unitType(2).portraitBackground = vbGreen
unitType(2).dimensions.x = 24
unitType(2).dimensions.y = 26
unitType(2).speed = 5
unitType(2).attackSpeed = 1000
unitType(2).frames = 3
unitType(2).lineOfSight = 80


unit(0).location.x = 30
unit(0).location.y = 20

unit(1).location.x = 150
unit(1).location.y = 80

unit(2).location.x = 100
unit(2).location.y = 130

For i = 0 To 2
   With unit(i)
      .type = 1
      .moving = False
      .direction = dirD
      .frame = 1
      .attackTimer = Int(Rnd * (unitType(unit(i).type).attackSpeed / 20) + 1) * 20 'This formula will be used in the actual engine
      .combatMode = False
      .selected = False
      .target = unit(i).location
      .targetUnit = -1
      .targetBuilding = -1
      .freezeFrame = False
      .exploring = True
      .player = 1
      .health = Int(Rnd * (unitType(unit(i).type).health)) + 1
   End With
Next i

unit(1).type = 2
unit(2).player = 2

corpseType(0).timer = 60
corpseType(0).dc = CreateCompatibleDC(0)
corpseType(0).dc = LoadGraphicDC(App.Path & "\Images\Guts.bmp")
corpseType(0).dimensions.x = 144
corpseType(0).dimensions.y = 144
corpseType(0).background = vbGreen

corpse(0).type = 0
corpse(0).location.x = 100
corpse(0).location.y = 100
corpse(0).dimensions.x = 40
corpse(0).dimensions.y = 60
corpse(0).timer = 10

civ(1).name = "Rebel Alliance"
civ(1).color = vbBlue

civ(2).name = "Matak"
civ(2).color = vbRed

player(1).name = "Anansi Fighters"
player(1).population = 2
player(1).civ = 1

player(2).name = "Roger Waters"
player(2).population = 1
player(2).civ = 2

needReExplore = True

End Sub
