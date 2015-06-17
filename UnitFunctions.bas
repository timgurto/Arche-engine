Attribute VB_Name = "mdlDrawingFunctions"
Option Explicit

Public Sub drawMap(m As typMap)
Dim X As Integer
Dim Y As Integer
Dim z As Long
For X = 1 To m.dimensions.X
   For Y = 1 To m.dimensions.Y
      z = BitBlt(frmGame.picGame.hdc, (X - 1) * TERRAIN_TILE_SIZE - gameMap.displacement.X, (Y - 1) * TERRAIN_TILE_SIZE - gameMap.displacement.Y, TERRAIN_TILE_SIZE, TERRAIN_TILE_SIZE, terrain(m.terrain(X, Y)).dc, 0, 0, vbSrcCopy)
   Next Y
Next X
End Sub
Public Sub drawUnit(u As typUnit)
Dim t As typUnitType
t = unitType(u.type)
Dim X As Long
X = TransparentBlt(frmGame.picGame.hdc, u.location.X - t.dimensions.X / 2 - gameMap.displacement.X, u.location.Y - t.dimensions.Y * (7 / 8) - gameMap.displacement.Y, t.dimensions.X, t.dimensions.Y, t.dc, u.direction * t.dimensions.X * t.frames + t.dimensions.X * (u.frame), 0, t.dimensions.X, t.dimensions.Y, unitType(u.type).background)
If DEBUG_MODE Then X = Rectangle(frmGame.picGame.hdc, u.location.X - t.dimensions.X / 2 - gameMap.displacement.X, u.location.Y - t.dimensions.Y * (7 / 8) - gameMap.displacement.Y, u.location.X + t.dimensions.X / 2 - gameMap.displacement.X, u.location.Y + t.dimensions.Y * (1 / 8) - gameMap.displacement.Y)
End Sub
Public Sub drawTarget(u As typUnit)
Dim X As Long
X = TransparentBlt(frmGame.picGame.hdc, u.target.X - target.dimensions.X / 2 - gameMap.displacement.X, u.target.Y - target.dimensions.Y / 2 - gameMap.displacement.Y, target.dimensions.X, target.dimensions.Y, target.dc, 0, 0, target.dimensions.X, target.dimensions.Y, target.background)
End Sub
Public Sub drawSelection(u As typUnit)
Dim t As typUnitType
t = unitType(u.type)
Dim X As Long
frmGame.picGame.ForeColor = White
X = Ellipse(frmGame.picGame.hdc, u.location.X - t.dimensions.X / 2 - gameMap.displacement.X, u.location.Y - t.dimensions.Y / 8 - gameMap.displacement.Y, u.location.X + t.dimensions.X / 2 - gameMap.displacement.X, u.location.Y + t.dimensions.Y / 8 - gameMap.displacement.Y)
End Sub

Public Sub drawSelectionRectangle()
Dim X As Long
If SELECTION_RECTANGLE_SHADOW Then
   frmGame.picGame.ForeColor = Black
   X = Rectangle(frmGame.picGame.hdc, selectionRectangleLoc1.X + 1, selectionRectangleLoc1.Y + 1, selectionRectangleLoc2.X + 1, selectionRectangleLoc2.Y + 1)
End If
frmGame.picGame.ForeColor = White
X = Rectangle(frmGame.picGame.hdc, selectionRectangleLoc1.X, selectionRectangleLoc1.Y, selectionRectangleLoc2.X, selectionRectangleLoc2.Y)
End Sub
