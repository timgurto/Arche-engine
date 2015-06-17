Attribute VB_Name = "mdlInit"
Option Explicit


Public Sub init()

unitType(1).dc = CreateCompatibleDC(0)

unitType(1).dc = LoadGraphicDC(App.Path & "\Images\u001.bmp")

unitType(1).speed = 3

unit(1).location.x = 20
unit(1).location.y = 0
unit(1).target.x = 120
unit(1).target.y = 65
unit(1).type = 1

End Sub
