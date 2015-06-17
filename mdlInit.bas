Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()

refreshCount = 0

unitType(1).dc = CreateCompatibleDC(0)
unitType(1).dc = LoadGraphicDC(App.Path & "\Images\Standard.bmp")
unitType(1).dimensions.x = 48
unitType(1).dimensions.y = 48
unitType(1).speed = 3
unitType(1).frames = 4

unit(1).location.x = 120
unit(1).location.y = 65
unit(1).target.x = 120
unit(1).target.y = 65
unit(1).type = 1
unit(1).moving = False
unit(1).direction = dirD
unit(1).frame = 1

End Sub
