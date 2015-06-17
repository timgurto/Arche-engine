Attribute VB_Name = "UnitFunctions"
Option Explicit

Public Sub drawUnit(unitToDraw As typUnit)
Dim x As Long
x = BitBlt(frmGame.picGame.hdc, unitToDraw.location.x, unitToDraw.location.y, unitType(unitToDraw.type).dimensions.x, unitType(unitToDraw.type).dimensions.y, unitType(unitToDraw.type).dc, 0, 0, vbSrcCopy)
End Sub
