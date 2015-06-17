Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()

unitType(1).dc = CreateCompatibleDC(0)
unitType(1).dc = LoadGraphicDC(App.Path & "\Images\u001.bmp")
unitType(1).dimensions.x = 54
unitType(1).dimensions.y = 56
unitType(1).speed = 3

unit(1).location.x = 120
unit(1).location.y = 65
unit(1).target.x = 120
unit(1).target.y = 65
unit(1).type = 1
unit(1).moving = False

End Sub
