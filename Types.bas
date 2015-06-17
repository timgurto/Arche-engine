Attribute VB_Name = "Types"
Option Explicit

Public Type typCoords
   x As Integer
   y As Integer
End Type


Public Type typUnitType
   dc As Long
   speed As Integer
   dimensions As typCoords
   frames As Byte
End Type

Public Type typUnit
   location As typCoords
   target As typCoords
   type As Byte
   moving As Boolean
   frame As Byte
   direction As Byte
End Type

