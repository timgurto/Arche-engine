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

Public Const you = 0

Public Const KEY_CTRL = 17
Public Const KEY_DELETE = 46

Public screenResolution As typcoords

Public Const REFRESHES_PER_FRAME As Integer = 6

Public Const CONQUEST = 0
Public Const REGICIDE = 1

Public Const UNIT_TYPE = 0

Public victoryType As Byte
Public regicideTarget(10) As Byte

Public Const MAX_PLAYERS = 10
Public activePlayers As Integer
Public Const MAX_CIVS = 10
Public activeCivs As Integer
Public Const MAX_UNITS = 150
Public activeUnits As Integer
Public Const MAX_UNIT_TYPES = 20
Public activeUnitTypes As Integer
Public Const MAX_TERRAINS = 10
Public activeTerrains As Integer
Public Const MAX_CORPSE_TYPES = 10
Public activeCorpseTypes As Integer
Public Const MAX_CORPSES = 1000
Public activeCorpses As Integer

Public refreshCount As Integer

Public scrollDir As Byte 'the direction the map is currently scrolling

Public player(MAX_PLAYERS) As typPlayer
Public civ(MAX_CIVS) As typCiv

Public unitType(MAX_UNIT_TYPES) As typUnitType
Public unit(MAX_UNITS) As typUnit

Public target As typTarget

Public terrain(MAX_TERRAINS) As typTerrain
Public terrainFrameTimer As Integer

Public corpseType(MAX_CORPSE_TYPES) As typCorpseType
Public corpse(MAX_CORPSES) As typCorpse

Public gameMap As typMap
Public fogDC As Long

Public ctrlDown As Boolean
Public mouseDown As Boolean

Public selectionRectangleLoc1 As typcoords
Public selectionRectangleLoc2 As typcoords

Public needReExplore As Boolean

Public selectionRectangle As Boolean

'***GAME OPTIONS - USE THESE TO CUSTOMIZE YOUR GAME***
Public Const SELECTION_RECTANGLE_SHADOW As Boolean = True 'whether the selection rectangle has a shadow
Public Const KEEP_WALKING_ON_COLLISION As Boolean = False 'whether units continue their walking animation while waiting for an obstructin to be removed
Public Const SHOW_SELECTED_TARGETS As Boolean = True 'whether selected units' targets are displayed
Public Const TERRAIN_TILE_SIZE As Integer = 48 'the x and y dimensions of each terrain tile
Public Const FOG_OF_WAR As Boolean = True
Public Const ENEMIES_SELECTABLE As Boolean = True
Public Const ENEMIES_HAVE_ELLIPSES As Boolean = False
Public Const SELECTION_ELLIPSE_WIDTH As Integer = 1
Public Const HEALTH_BAR_WIDTH As Integer = 4
Public Const HEALTH_BAR_COLOR As Single = vbGreen
Public Const SHOW_UNUSED_STATS As Boolean = False 'Whether stats are displayed if a unit doesn't have them, eg. 0 armor, no special
Public Const SPECIAL_PERCENT As Boolean = False 'Whether a unit's 'special' is displayed as a percentage or not
Public Const SPECIAL_NAME As String = "Mana"
Public Const PORTRAIT_WIDTH As Integer = 15
Public Const PORTRAIT_HEIGHT As Integer = 15
Public Const AUTO_ATTACKING As Boolean = True
Public Const AUTO_ATTACK_RANGE As Integer = 150
Public Const RANGED_UNIT As Integer = 50
Public Const TERRAIN_FRAME_LENGTH = 200 'ms
'*****************************************************
