Attribute VB_Name = "mdlGameLoop"
Option Explicit

Public Sub runGameLoop()
refreshCount = refreshCount + 1
If refreshCount = REFRESHES_PER_FRAME Then
   refreshCount = 0
   unit(1).frame = unit(1).frame + 1
   If unit(1).frame >= unitType(unit(1).type).frames Then unit(1).frame = 0
End If

If unit(1).moving Then
   unit(1).location = addCoords(unit(1).location, findPath(unit(1)))
Else
   unit(1).frame = 0
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
