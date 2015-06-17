Attribute VB_Name = "mdlGameLoop"
Option Explicit

Public Sub gameLoop()

Const tickDifference As Long = 20
Dim lastTick As Long
Dim currentTick As Long

Do
   currentTick = GetTickCount
   If currentTick - lastTick > tickDifference Then
      
      '=========================================
      
      unit(1).location.x = unit(1).location.x + unitType(unit(1).type).speed
      frmGame.picGame.Cls
      drawUnit unit(1)
      frmGame.picGame.Refresh
      
      '=========================================
         
      lastTick = currentTick
   End If
   DoEvents
   
Loop



End Sub
