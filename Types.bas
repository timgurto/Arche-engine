Attribute VB_Name = "Types"
Option Explicit

Public Type typCoords
   x As Integer
   y As Integer
End Type


Public Type typUnitType
   dc As Long
   speed As Single
End Type

Public caveman As typUnitType


Public Type typUnit
   location As typCoords
   target As typCoords
   type As Byte
End Type

