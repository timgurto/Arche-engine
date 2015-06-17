Attribute VB_Name = "mdlTypes"
Option Explicit

Public Type typcoords
   x As Integer
   y As Integer
End Type


Public Type typTarget
   dc As Long
   background As Long
   dimensions As typcoords
End Type

Public Type typTerrain
   dc As Long
End Type

Public Type typMap
   dimensions As typcoords
   displacement As typcoords
   terrain(100, 100) As Integer
   explored(100, 100) As Boolean
   fog(100, 100) As Boolean
End Type

Public Type typCiv
   name As String
   color As Long
End Type

Public Type typPlayer
   name As String
   population As Integer
   civ As Byte
End Type

Public Type typUnitType
   name As String
   speed As Integer
   attackSpeed As Integer 'ms.  Multiple of 20.  Attacks every x ms.
   health As Integer
   armor As Integer
   attack As Integer
   range As Integer
   isRanged As Boolean
   healing As Integer
   corpse As Byte
   lineOfSight As Integer
   dc As Long
   portrait As Long
   portraitBackground As Long
   background As Long
   dimensions As typcoords
   frames As Byte
   taunting As Boolean 'Whether this unit's attacks force its targets to attack it
End Type

Public Type typUnit
   health As Integer
   special As Integer
   location As typcoords
   targetUnit As Integer
   targetBuilding As Integer
   target As typcoords
   player As Byte
   type As Byte
   moving As Boolean
   frame As Byte
   attackTimer As Integer
   direction As Byte
   selected As Boolean
   freezeFrame As Boolean 'whether to freeze the next frame of animation
   exploring As Boolean
   combatMode As Boolean 'whether this unit is in the combat half of its attacking cycle
End Type

Public Type typCorpseType
   dimensions As typcoords
   background As Long
   dc As Long
   timer As Integer 'how long the corpse stays in the game.  -1 = forever
End Type

Public Type typCorpse
   type As Byte
   location As typcoords
   dimensions As typcoords
   timer As Integer
End Type
