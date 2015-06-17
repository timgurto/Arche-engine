Attribute VB_Name = "mdlDeclarations"
Option Explicit

Public Const DEBUG_MODE = False
Public Const Red = &HFF
Public Const Green = &HFF00&
Public Const Blue = &HFF0000
Public Const Yellow = &HFFFF&
Public Const Cyan = &HFFFF00
Public Const Magenta = &HFF00FF
Public Const Black = &H0
Public Const White = &HFFFFFF

Public Const dirN = -1 'not moving
Public Const dirU = 0
Public Const dirR = 1
Public Const dirD = 2
Public Const dirL = 3
Public Const dirE = 4
Public Const dirF = 5
Public Const dirG = 6
Public Const dirH = 7


Public Const KEY_CTRL = 17
Public Const KEY_DELETE = 46

Public Const REFRESHES_PER_FRAME As Integer = 6

Public refreshCount As Integer

Public scrollDir As Integer 'the direction the map is currently scrolling

Public unitType(10) As typUnitType
Public unit(100) As typUnit
Public activeUnits As Integer

Public target As typTarget

Public terrain(10) As typTerrain
Public gameMap As typMap

Public ctrlDown As Boolean
Public mouseDown As Boolean

Public selectionRectangleLoc1 As typCoords
Public selectionRectangleLoc2 As typCoords

'***GAME OPTIONS - USE THESE TO CUSTOMIZE YOUR GAME***
Public Const SELECTION_RECTANGLE_SHADOW As Boolean = True 'whether the selection rectangle has a shadow
Public Const KEEP_WALKING_ON_COLLISION As Boolean = False 'whether units continue their walking animation while waiting for an obstructin to be removed
Public Const SHOW_SELECTED_TARGETS As Boolean = True 'whether selected units' targets are displayed
Public Const TERRAIN_TILE_SIZE As Integer = 48 'the x and y dimensions of each terrain tile
'*****************************************************
