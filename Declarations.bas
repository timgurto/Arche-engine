Attribute VB_Name = "Declarations"
Option Explicit

Public Const Red = &HFF
Public Const Green = &HFF00&
Public Const Blue = &HFF0000
Public Const Yellow = &HFFFF&
Public Const Cyan = &HFFFF00
Public Const Magenta = &HFF00FF
Public Const Black = &H0
Public Const White = &HFFFFFF

Public Const dirU = 0
Public Const dirR = 1
Public Const dirD = 2
Public Const dirL = 3

Public Const KEY_CTRL = 17

Public Const REFRESHES_PER_FRAME As Integer = 6

Public refreshCount As Integer

Public unitType(10) As typUnitType

Public unit(100) As typUnit

Public activeUnits As Integer

Public i As Integer

Public ctrlDown As Boolean
Public mouseDown As Boolean

Public selectionRectangleLoc1 As typCoords
Public selectionRectangleLoc2 As typCoords

'***GAME OPTIONS - USE THESE TO CUSTOMIZE YOUR GAME***
Public selectionRectangleShadow As Boolean 'whether the selection rectangle has a shadow
Public keepWalkingOnCollision As Boolean 'whether units continue their walking animation while waiting for an obstructin to be removed

