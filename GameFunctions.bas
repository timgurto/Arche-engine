Attribute VB_Name = "GameFunctions"
Option Explicit

Public Sub swapUnits(a, b)
Dim temp As typUnit
temp = unit(a)
unit(a) = unit(b)
unit(b) = temp
End Sub

Public Sub deleteUnit(n As Integer)
swapUnits n, activeUnits - 1

End Sub

Public Sub deleteUnits()
Dim i As Integer
i = 0
While i < activeUnits
'For i = 0 To activeUnits - 1
   If unit(i).selected Then
      deleteUnit (i)
      activeUnits = activeUnits - 1
      i = i - 1
   End If
'Next i
i = i + 1
Wend
End Sub
