Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()

refreshCount = 0

unitCount = 3

ctrlDown = False
mouseDown = False: frmGame.mouseDownIndicator = mouseDown

selectionRectangleShadow = False

unitType(1).dc = CreateCompatibleDC(0)
unitType(1).dc = LoadGraphicDC(App.Path & "\Images\Standard.bmp")
unitType(1).dimensions.X = 48
unitType(1).dimensions.Y = 48
unitType(1).speed = 3
unitType(1).frames = 4

unitType(2).dc = CreateCompatibleDC(0)
unitType(2).dc = LoadGraphicDC(App.Path & "\Images\Hero.bmp")
unitType(2).dimensions.X = 24
unitType(2).dimensions.Y = 32
unitType(2).speed = 3
unitType(2).frames = 3


unit(0).location.X = 120
unit(0).location.Y = 65

unit(1).location.X = 200
unit(1).location.Y = 100

unit(2).location.X = 250
unit(2).location.Y = 300

For i = 0 To 2
   unit(i).type = 1
   unit(i).moving = False
   unit(i).direction = dirD
   unit(i).frame = 1
   unit(i).selected = False
   unit(i).target = unit(i).location
Next i

unit(1).type = 2

End Sub
