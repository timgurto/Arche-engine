Attribute VB_Name = "UnitFunctions"
Option Explicit

Public Sub drawUnit(u As typUnit)
Dim t As typUnitType
t = unitType(u.type)
Dim x As Long
x = TransparentBlt(frmGame.picGame.hDC, u.location.x - t.dimensions.x / 2, u.location.y - t.dimensions.y * (7 / 8), t.dimensions.x, t.dimensions.y, t.dc, u.direction * t.dimensions.x * t.frames + t.dimensions.x * (u.frame), 0, t.dimensions.x, t.dimensions.y, vbWhite) 'IIf(u.selected, Blue, White))
End Sub

Public Sub drawselection(u As typUnit)
Dim t As typUnitType
t = unitType(u.type)
Dim x As Long
frmGame.picGame.ForeColor = &HFFFFFF
x = Ellipse(frmGame.picGame.hDC, u.location.x - t.dimensions.x / 2, u.location.y - t.dimensions.y / 8, u.location.x + t.dimensions.x / 2, u.location.y + t.dimensions.y / 8)
End Sub

