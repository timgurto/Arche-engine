Attribute VB_Name = "UnitFunctions"
Option Explicit

Public Sub drawUnit(u As typUnit)
Dim t As typUnitType
t = unitType(u.type)
Dim x As Long
x = TransparentBlt(frmGame.picGame.hdc, u.location.x, u.location.y, t.dimensions.x, t.dimensions.y, t.dc, u.direction * t.dimensions.x * t.frames + t.dimensions.x * (u.frame), 0, t.dimensions.x, t.dimensions.y, White)
End Sub
