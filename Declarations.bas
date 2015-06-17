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

Public gamePath As String

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
Public HEALTH_BAR_CIV_COLOR As Boolean
Public SELECTION_RECTANGLE_SHADOW As Boolean
Public KEEP_WALKING_ON_COLLISION As Boolean
Public SHOW_SELECTED_TARGETS As Boolean
Public TERRAIN_TILE_SIZE As Integer
Public FOG_OF_WAR As Boolean
Public ENEMIES_SELECTABLE As Boolean
Public ENEMIES_HAVE_ELLIPSES As Boolean
Public YOU_HAVE_ELLIPSES As Boolean
Public SELECTION_ELLIPSE_WIDTH As Integer
Public HEALTH_BAR_WIDTH As Integer
Public HEALTH_BAR_COLOR As Single
Public SHOW_UNUSED_STATS As Boolean
Public SPECIAL_PERCENT As Boolean
Public SPECIAL_NAME As String
Public PORTRAIT_WIDTH As Integer
Public PORTRAIT_HEIGHT As Integer
Public AUTO_ATTACKING As Boolean
Public AUTO_ATTACK_RANGE As Integer
Public RANGED_UNIT As Integer
Public TERRAIN_FRAME_LENGTH As Integer
'*****************************************************
