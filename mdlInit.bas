Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()
Dim i As Integer
Dim j As Integer

refreshCount = 0

activeUnits = 7
activeCorpses = 1

ctrlDown = False

scrollDir = dirN

target.dimensions.x = 16
target.dimensions.y = 16
target.dc = CreateCompatibleDC(0)
target.dc = LoadGraphicDC(App.Path & "\Images\target.bmp")
target.background = vbRed

gameMap.displacement = makeCoords(-100, -100)

fogDC = CreateCompatibleDC(0)
fogDC = LoadGraphicDC(App.Path & "\Images\fog.bmp")

terrain(1).dc = CreateCompatibleDC(0)
terrain(1).dc = LoadGraphicDC(App.Path & "\Images\grass.bmp")

terrain(2).dc = CreateCompatibleDC(0)
terrain(2).dc = LoadGraphicDC(App.Path & "\Images\dirt.bmp")

gameMap.dimensions = makeCoords(15, 10)
For i = 1 To 15
   For j = 1 To 15
      gameMap.terrain(i, j) = Int(Rnd * (2) + 1)
      gameMap.explored(i, j) = False
   Next j
Next i

For i = 1 To 5
unitType(i).corpse = 0
Next i

unitType(1).name = "Standard"
unitType(1).health = 100
unitType(1).armor = 2
unitType(1).attack = 8
unitType(1).dc = CreateCompatibleDC(0)
unitType(1).dc = LoadGraphicDC(App.Path & "\Images\Standard.bmp")
unitType(1).portrait = CreateCompatibleDC(0)
unitType(1).portrait = LoadGraphicDC(App.Path & "\Images\StandardPortrait.bmp")
unitType(1).background = vbWhite
unitType(1).portraitBackground = vbWhite
unitType(1).dimensions.x = 48
unitType(1).dimensions.y = 48
unitType(1).speed = 2
unitType(1).attackSpeed = 2500
unitType(1).frames = 4
unitType(1).lineOfSight = 150
unitType(1).taunting = False

unitType(3).name = "Shield"
unitType(3).health = 1000
unitType(3).armor = 50
unitType(3).attack = 2
unitType(3).dc = CreateCompatibleDC(0)
unitType(3).dc = LoadGraphicDC(App.Path & "\Images\Shield.bmp")
unitType(3).portrait = CreateCompatibleDC(0)
unitType(3).portrait = LoadGraphicDC(App.Path & "\Images\HeroPortrait.bmp")
unitType(3).background = vbGreen
unitType(3).portraitBackground = vbGreen
unitType(3).dimensions.x = 53
unitType(3).dimensions.y = 56
unitType(3).speed = 2
unitType(3).attackSpeed = 1000
unitType(3).frames = 1
unitType(3).lineOfSight = 100
unitType(3).taunting = True

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
unitType(2).speed = 4
unitType(2).attackSpeed = 1000
unitType(2).frames = 3
unitType(2).lineOfSight = 80
unitType(2).taunting = False

unitType(4).name = "Cleric"
unitType(4).health = 40
unitType(4).armor = 0
unitType(4).attack = 1
unitType(4).healing = 20
unitType(4).range = 3
unitType(4).dc = CreateCompatibleDC(0)
unitType(4).dc = LoadGraphicDC(App.Path & "\Images\cleric.bmp")
unitType(4).portrait = CreateCompatibleDC(0)
unitType(4).portrait = LoadGraphicDC(App.Path & "\Images\clericPortrait.bmp")
unitType(4).background = vbGreen
unitType(4).portraitBackground = vbGreen
unitType(4).dimensions.x = 27
unitType(4).dimensions.y = 42
unitType(4).speed = 3
unitType(4).attackSpeed = 400
unitType(4).frames = 1
unitType(4).lineOfSight = 80
unitType(4).taunting = False

unitType(5).name = "King Snowman"
unitType(5).health = 250
unitType(5).armor = 5
unitType(5).attack = 0
unitType(5).healing = 0
unitType(5).range = 0
unitType(5).dc = CreateCompatibleDC(0)
unitType(5).dc = LoadGraphicDC(App.Path & "\Images\Snowman.bmp")
unitType(5).portrait = CreateCompatibleDC(0)
unitType(5).portrait = LoadGraphicDC(App.Path & "\Images\SnowmanPortrait.bmp")
unitType(5).background = vbGreen
unitType(5).portraitBackground = vbGreen
unitType(5).dimensions.x = 48
unitType(5).dimensions.y = 48
unitType(5).speed = 0
unitType(5).attackSpeed = 0
unitType(5).frames = 1
unitType(5).lineOfSight = 10
unitType(5).taunting = False

unit(0).location.x = 230
unit(0).location.y = 20

unit(1).location.x = 150
unit(1).location.y = 80

unit(2).location.x = 100
unit(2).location.y = 130

unit(3).location.x = 50
unit(3).location.y = 150

unit(4).location.x = 200
unit(4).location.y = 100

unit(5).location.x = 300
unit(5).location.y = 200

unit(6).location.x = 40
unit(6).location.y = 40

For i = 0 To 6
   With unit(i)
      .type = 1
   unit(3).type = 3
   unit(1).type = 2
   unit(4).type = 4
   unit(5).type = 5
   unit(6).type = 5
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
      .player = 2
      .health = Int(Rnd * (unitType(unit(i).type).health)) + 1
   End With
Next i

unit(2).player = 1
unit(3).player = 1
unit(4).player = 1
unit(6).player = 1
unit(6).health = unitType(5).health

unit(5).health = unitType(5).health

victoryType = REGICIDE
regicideTarget(1) = 5
regicideTarget(2) = 6

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

civ(1).name = "Egyptians"
civ(1).color = vbBlue

civ(2).name = "Greeks"
civ(2).color = vbRed

player(1).name = "Pharaoh Rameses"
player(1).population = 2
player(1).civ = 1

player(2).name = "Alexander the Great"
player(2).population = 1
player(2).civ = 2

needReExplore = True

End Sub
