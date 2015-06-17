Attribute VB_Name = "mdlDeclarations"
Option Explicit

Public Const DEBUG_MODE = False
Public Const dirN = 99 'not moving
Public Const dirU = 0
Public Const dirR = 1
Public Const dirD = 2
Public Const dirL = 3
Public Const dirE = 4
Public Const dirF = 5
Public Const dirG = 6
Public Const dirH = 7

Public Const you = 1

Public Const KEY_CTRL = 17
Public Const KEY_DELETE = 46

Public Const REFRESHES_PER_FRAME As Integer = 6

Public refreshCount As Integer

Public scrollDir As Byte 'the direction the map is currently scrolling

Public player(2) As typPlayer
Public civ(2) As typCiv

Public unitType(10) As typUnitType
Public unit(100) As typUnit
Public activeUnits As Integer

Public target As typTarget

Public terrain(10) As typTerrain

Public gameMap As typMap
Public fogDC As Long

Public ctrlDown As Boolean
Public mouseDown As Boolean

Public selectionRectangleLoc1 As typCoords
Public selectionRectangleLoc2 As typCoords

Public needReExplore As Boolean

'***GAME OPTIONS - USE THESE TO CUSTOMIZE YOUR GAME***
Public Const SELECTION_RECTANGLE_SHADOW As Boolean = True 'whether the selection rectangle has a shadow
Public Const KEEP_WALKING_ON_COLLISION As Boolean = False 'whether units continue their walking animation while waiting for an obstructin to be removed
Public Const SHOW_SELECTED_TARGETS As Boolean = True 'whether selected units' targets are displayed
Public Const TERRAIN_TILE_SIZE As Integer = 48 'the x and y dimensions of each terrain tile
Public Const FOG_OF_WAR As Boolean = False
Public Const ENEMIES_SELECTABLE As Boolean = True
Public Const ENEMIES_HAVE_ELLIPSES As Boolean = False
Public Const SELECTION_ELLIPSE_WIDTH As Integer = 1
Public Const HEALTH_BAR_WIDTH As Integer = 4
Public Const HEALTH_BAR_COLOR As Single = vbGreen
Public Const SHOW_UNUSED_STATS As Boolean = True 'Whether stats are displayed if a unit doesn't have them, eg. 0 armor, no special
Public Const SPECIAL_PERCENT As Boolean = False 'Whether a unit's 'special' is displayed as a percentage or not
Public Const SPECIAL_NAME As String = "Influence"
'*****************************************************
