Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()
Dim i As Integer

refreshCount = 0

activeUnits = 3

ctrlDown = False

target.dimensions.x = 16
target.dimensions.y = 16
target.dc = CreateCompatibleDC(0)
target.dc = LoadGraphicDC(App.Path & "\Images\target.bmp")
target.background = vbRed

terrain(1).dc = CreateCompatibleDC(0)
terrain(1).dc = LoadGraphicDC(App.Path & "\Images\grass.bmp")

terrain(2).dc = CreateCompatibleDC(0)
terrain(2).dc = LoadGraphicDC(App.Path & "\Images\dirt.bmp")

testMap.dimensions = makeCoords(4, 3)
testMap.terrain(1, 1) = 2
testMap.terrain(1, 2) = 2
testMap.terrain(1, 3) = 1 '    1 2 3 4
testMap.terrain(2, 1) = 2 '  +--------
testMap.terrain(2, 2) = 1 ' 1| 2 2 1 1
testMap.terrain(2, 3) = 1 ' 2| 2 1 1 1
testMap.terrain(3, 1) = 1 ' 3| 1 1 1 1
testMap.terrain(3, 2) = 1
testMap.terrain(3, 3) = 1
testMap.terrain(4, 1) = 1
testMap.terrain(4, 2) = 1
testMap.terrain(4, 3) = 1

unitType(1).name = "Standard"
unitType(1).dc = CreateCompatibleDC(0)
unitType(1).dc = LoadGraphicDC(App.Path & "\Images\Standard.bmp")
unitType(1).background = vbGreen
unitType(1).dimensions.x = 48
unitType(1).dimensions.y = 48
unitType(1).speed = 3
unitType(1).frames = 4

unitType(2).name = "Hero"
unitType(2).dc = CreateCompatibleDC(0)
unitType(2).dc = LoadGraphicDC(App.Path & "\Images\Hero.bmp")
unitType(2).background = vbGreen
unitType(2).dimensions.x = 24
unitType(2).dimensions.y = 26
unitType(2).speed = 5
unitType(2).frames = 3


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
Next i

unit(1).type = 2

End Sub
