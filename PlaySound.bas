Attribute VB_Name = "X__external_PlaySound"
'***EXTERNAL MODULE***
'PlaySound
'Clint V Franklin
'*********************


Option Explicit

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long

  

Public Const sndAsync = &H1
Public Const sndLoop = &H8
Public Const sndNoStop = &H10

