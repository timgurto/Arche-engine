Attribute VB_Name = "mdlGameLoop"
Option Explicit

Public Sub runGameLoop()
refreshCount = refreshCount + 1
If refreshCount = REFRESHES_PER_FRAME Then
   For i = 0 To unitCount - 1
      refreshCount = 0
      unit(i).frame = unit(i).frame + 1
      If unit(i).frame >= unitType(unit(i).type).frames Then unit(i).frame = 0
   Next i
End If

For i = 0 To unitCount - 1
   If unit(i).moving Then
      unit(i).location = addCoords(unit(i).location, findPath(unit(i)))
   Else
      unit(i).frame = 0
   End If
   
   unit(i).moving = distance(unit(i).location, unit(i).target) > unitType(unit(i).type).speed

Next i

drawEverything

End Sub

Public Sub drawEverything()
frmGame.picGame.Cls



For i = 0 To unitCount - 1
   If unit(i).selected Then drawSelection unit(i)
   drawUnit unit(i)
Next i

drawSelectionRectangle

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
