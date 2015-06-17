Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()

refreshCount = 0

activeUnits = 3

ctrlDown = False

'***GAME OPTIONS - USE THESE TO CUSTOMIZE YOUR GAME***
selectionRectangleShadow = False
keepWalkingOnCollision = False
showSelectedTargets = True
'*****************************************************

target.dimensions.x = 16
target.dimensions.y = 16
target.dc = CreateCompatibleDC(0)
target.dc = LoadGraphicDC(App.Path & "\Images\target.bmp")

unitType(1).dc = CreateCompatibleDC(0)
unitType(1).dc = LoadGraphicDC(App.Path & "\Images\Standard.bmp")
unitType(1).dimensions.x = 48
unitType(1).dimensions.y = 48
unitType(1).speed = 3
unitType(1).frames = 4

unitType(2).dc = CreateCompatibleDC(0)
unitType(2).dc = LoadGraphicDC(App.Path & "\Images\Hero.bmp")
unitType(2).dimensions.x = 24
unitType(2).dimensions.y = 26
unitType(2).speed = 3
unitType(2).frames = 3


unit(0).location.x = 120
unit(0).location.y = 65

unit(1).location.x = 200
unit(1).location.y = 100

unit(2).location.x = 250
unit(2).location.y = 300

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
