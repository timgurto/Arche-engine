Attribute VB_Name = "DrawingFunctions"
Option Explicit

Public Sub drawUnit(u As typUnit)
Dim t As typUnitType
t = unitType(u.type)
Dim X As Long
X = TransparentBlt(frmGame.picGame.hdc, u.location.X - t.dimensions.X / 2, u.location.Y - t.dimensions.Y * (7 / 8), t.dimensions.X, t.dimensions.Y, t.dc, u.direction * t.dimensions.X * t.frames + t.dimensions.X * (u.frame), 0, t.dimensions.X, t.dimensions.Y, vbWhite) 'IIf(u.selected, Blue, White))
End Sub

Public Sub drawSelection(u As typUnit)
Dim t As typUnitType
t = unitType(u.type)
Dim X As Long
frmGame.picGame.ForeColor = White
X = Ellipse(frmGame.picGame.hdc, u.location.X - t.dimensions.X / 2, u.location.Y - t.dimensions.Y / 8, u.location.X + t.dimensions.X / 2, u.location.Y + t.dimensions.Y / 8)
End Sub

Public Sub drawSelectionRectangle()
Dim X As Long
If selectionrectangleshadow Then
   frmGame.picGame.ForeColor = Black
   X = Rectangle(frmGame.picGame.hdc, selectionRectangleLoc1.X + 1, selectionRectangleLoc1.Y + 1, selectionRectangleLoc2.X + 1, selectionRectangleLoc2.Y + 1)
End If
frmGame.picGame.ForeColor = White
X = Rectangle(frmGame.picGame.hdc, selectionRectangleLoc1.X, selectionRectangleLoc1.Y, selectionRectangleLoc2.X, selectionRectangleLoc2.Y)
End Sub
