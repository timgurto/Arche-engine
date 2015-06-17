Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()
Dim i As Integer
Dim j As Integer

refreshCount = 0

activeUnits = 3

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
unitType(1).dc = CreateCompatibleDC(0)
unitType(1).dc = LoadGraphicDC(App.Path & "\Images\Standard.bmp")
unitType(1).background = vbWhite
unitType(1).dimensions.x = 48
unitType(1).dimensions.y = 48
unitType(1).speed = 3
unitType(1).frames = 4
unitType(1).lineOfSight = 150

unitType(2).name = "Hero"
unitType(2).health = 75
unitType(2).dc = CreateCompatibleDC(0)
unitType(2).dc = LoadGraphicDC(App.Path & "\Images\Hero.bmp")
unitType(2).background = vbGreen
unitType(2).dimensions.x = 24
unitType(2).dimensions.y = 26
unitType(2).speed = 5
unitType(2).frames = 3
unitType(2).lineOfSight = 80


unit(0).location.x = 30
unit(0).location.y = 20

unit(1).location.x = 150
unit(1).location.y = 80

unit(2).location.x = 100
unit(2).location.y = 130

For i = 0 To 2
   unit(i).type = 1
   unit(i).moving = False
   unit(i).direction = dirD
   unit(i).frame = 1
   unit(i).selected = False
   unit(i).target = unit(i).location
   unit(i).freezeFrame = False
   unit(i).exploring = True
   unit(i).player = 1
   unit(i).health = Int(Rnd * (unitType(unit(i).type).health)) + 1
Next i

unit(1).type = 2
unit(2).player = 2

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
