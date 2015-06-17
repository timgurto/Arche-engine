Attribute VB_Name = "mdlGameLoop"
Option Explicit

Public Sub runGameLoop()
If unit(1).moving Then
   unit(1).location = addCoords(unit(1).location, findPath(unit(1)))
End If
   
unit(1).moving = distance(unit(1).location, unit(1).target) > unitType(unit(1).type).speed

frmGame.picGame.Cls
drawUnit unit(1)
frmGame.picGame.Refresh

End Sub

Public Sub gameLoop()

Const tickDifference As Long = 20
Dim lastTick As Long
Dim currentTick As Long

Do
   currentTick = GetTickCount
   If currentTick - lastTick > tickDifference Then
      
      runGameLoop
         
      lastTick = currentTick
   End If
   DoEvents
   
Loop



End Sub
