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

Public Const REFRESHES_PER_FRAME As Integer = 6

Public refreshCount As Integer

Public unitType(10) As typUnitType

Public unit(100) As typUnit

Public unitCount As Integer

Public i As Integer

