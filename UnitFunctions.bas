Attribute VB_Name = "mdlDrawingFunctions"
Option Explicit

Public Sub drawMap(m As typMap)
Dim x As Integer
Dim y As Integer
Dim z As Long
For x = 1 To m.dimensions.x
   For y = 1 To m.dimensions.y
      'z = TransparentBlt(frmGame.picGame.hdc, (x - 1) * TERRAIN_TILE_SIZE, (y - 1) * TERRAIN_TILE_SIZE, TERRAIN_TILE_SIZE, TERRAIN_TILE_SIZE, terrain(m.terrain(x, y)).dc, 0, 0, TERRAIN_TILE_SIZE, TERRAIN_TILE_SIZE, a)
      z = BitBlt(frmGame.picGame.hdc, (x - 1) * TERRAIN_TILE_SIZE, (y - 1) * TERRAIN_TILE_SIZE, TERRAIN_TILE_SIZE, TERRAIN_TILE_SIZE, terrain(m.terrain(x, y)).dc, 0, 0, vbSrcCopy)
   Next y
Next x
End Sub
Public Sub drawUnit(u As typUnit)
Dim t As typUnitType
t = unitType(u.type)
Dim x As Long
x = TransparentBlt(frmGame.picGame.hdc, u.location.x - t.dimensions.x / 2, u.location.y - t.dimensions.y * (7 / 8), t.dimensions.x, t.dimensions.y, t.dc, u.direction * t.dimensions.x * t.frames + t.dimensions.x * (u.frame), 0, t.dimensions.x, t.dimensions.y, unitType(u.type).background)
If DEBUG_MODE Then x = Rectangle(frmGame.picGame.hdc, u.location.x - t.dimensions.x / 2, u.location.y - t.dimensions.y * (7 / 8), u.location.x + t.dimensions.x / 2, u.location.y + t.dimensions.y * (1 / 8))
End Sub
Public Sub drawTarget(u As typUnit)
Dim x As Long
x = TransparentBlt(frmGame.picGame.hdc, u.target.x - target.dimensions.x / 2, u.target.y - target.dimensions.y / 2, target.dimensions.x, target.dimensions.y, target.dc, 0, 0, target.dimensions.x, target.dimensions.y, target.background)
End Sub
Public Sub drawSelection(u As typUnit)
Dim t As typUnitType
t = unitType(u.type)
Dim x As Long
frmGame.picGame.ForeColor = White
x = Ellipse(frmGame.picGame.hdc, u.location.x - t.dimensions.x / 2, u.location.y - t.dimensions.y / 8, u.location.x + t.dimensions.x / 2, u.location.y + t.dimensions.y / 8)
End Sub

Public Sub drawSelectionRectangle()
Dim x As Long
If SELECTION_RECTANGLE_SHADOW Then
   frmGame.picGame.ForeColor = Black
   x = Rectangle(frmGame.picGame.hdc, selectionRectangleLoc1.x + 1, selectionRectangleLoc1.y + 1, selectionRectangleLoc2.x + 1, selectionRectangleLoc2.y + 1)
End If
frmGame.picGame.ForeColor = White
x = Rectangle(frmGame.picGame.hdc, selectionRectangleLoc1.x, selectionRectangleLoc1.y, selectionRectangleLoc2.x, selectionRectangleLoc2.y)
End Sub
