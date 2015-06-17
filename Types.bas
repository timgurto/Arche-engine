Attribute VB_Name = "mdlTypes"
Option Explicit

Public Type typCoords
   x As Integer
   y As Integer
End Type


Public Type typTarget
   dc As Long
   background As Long
   dimensions As typCoords
End Type

Public Type typTerrain
   dc As Long
End Type

Public Type typMap
   dimensions As typCoords
   displacement As typCoords
   terrain(100, 100) As Integer
   explored(100, 100) As Boolean
   fog(100, 100) As Boolean
End Type

Public Type typUnitType
   name As String
   speed As Integer
   lineOfSight As Integer
   dc As Long
   background As Long
   dimensions As typCoords
   frames As Byte
End Type

Public Type typUnit
   location As typCoords
   target As typCoords
   type As Byte
   moving As Boolean
   frame As Integer
   direction As Byte
   selected As Boolean
   freezeFrame As Boolean 'whether to freeze the next frame of animation
End Type

