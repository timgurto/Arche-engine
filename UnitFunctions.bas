Attribute VB_Name = "UnitFunctions"
Option Explicit

Public Sub drawUnit(unit As typUnit)
Dim x As Long
   x = BitBlt(Game.picGame.hdc, unit.target.x, unit.target.y, 54, 56, caveman.dc, 0, 0, vbSrcCopy)
End Sub
