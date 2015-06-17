VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   795
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCorpses 
      Interval        =   1000
      Left            =   14160
      Top             =   10200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   9480
      Width           =   2775
      Begin VB.PictureBox picPortrait 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00400000&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   60
         ScaleHeight     =   645
         ScaleWidth      =   645
         TabIndex        =   30
         Top             =   60
         Width           =   675
      End
      Begin VB.Label lblCiv 
         BackStyle       =   0  'Transparent
         Caption         =   "Egyptians"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   780
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1860
         TabIndex        =   26
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblHealing 
         BackStyle       =   0  'Transparent
         Caption         =   "3400/5000"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1860
         TabIndex        =   25
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblRange 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1860
         TabIndex        =   24
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblArmor 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   1560
         Width           =   855
      End
      Begin VB.Image imgRange 
         Height          =   360
         Left            =   1440
         Picture         =   "Game.frx":000C
         Top             =   1140
         Width           =   360
      End
      Begin VB.Image imgHealing 
         Height          =   360
         Left            =   1440
         Picture         =   "Game.frx":108E
         Top             =   780
         Width           =   360
      End
      Begin VB.Image imgSPeed 
         Height          =   360
         Left            =   1440
         Picture         =   "Game.frx":2110
         Top             =   1500
         Width           =   360
      End
      Begin VB.Image imgAttack 
         Height          =   360
         Left            =   60
         Picture         =   "Game.frx":3192
         Top             =   1140
         Width           =   360
      End
      Begin VB.Image imgHealth 
         Height          =   360
         Left            =   60
         Picture         =   "Game.frx":4214
         Top             =   780
         Width           =   360
      End
      Begin VB.Label lblAttack 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblHealth 
         BackStyle       =   0  'Transparent
         Caption         =   "3400/5000"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Image imgArmor 
         Height          =   360
         Left            =   60
         Picture         =   "Game.frx":5296
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label lblPlayer 
         BackStyle       =   0  'Transparent
         Caption         =   "Pharaoh Rameses"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   780
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblType 
         BackStyle       =   0  'Transparent
         Caption         =   "Axeman"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   780
         TabIndex        =   19
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Timer tmrScroll 
      Interval        =   8
      Left            =   14520
      Top             =   11040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Unit(s)"
      Height          =   375
      Left            =   10200
      TabIndex        =   2
      Top             =   10920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create new unit"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   10920
      Width           =   1695
   End
   Begin VB.PictureBox picGame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   9240
      Left            =   45
      ScaleHeight     =   9210
      ScaleWidth      =   15240
      TabIndex        =   0
      Top             =   45
      Width           =   15270
   End
   Begin VB.Label lblTargetUnit 
      Caption         =   "Label1"
      Height          =   255
      Left            =   9840
      TabIndex        =   29
      Top             =   10320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSelected 
      Height          =   255
      Left            =   9000
      TabIndex        =   27
      Top             =   10320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblFrame 
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   10800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape shpExplore 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   8160
      Top             =   10920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUnits 
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   10800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblMapCoords 
      Height          =   255
      Left            =   12120
      TabIndex        =   15
      Top             =   11280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblDisplacement 
      Height          =   255
      Left            =   12120
      TabIndex        =   14
      Top             =   11040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblCoords 
      Height          =   255
      Left            =   12120
      TabIndex        =   13
      Top             =   10800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblBottomLeft 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   0
      TabIndex        =   12
      Top             =   11475
      Width           =   45
   End
   Begin VB.Label lblTopLeft 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblBottomRight 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   15315
      TabIndex        =   10
      Top             =   11475
      Width           =   45
   End
   Begin VB.Label lblTopRight 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   15315
      TabIndex        =   9
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblbottom 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   45
      TabIndex        =   8
      Top             =   11475
      Width           =   15270
   End
   Begin VB.Label lblTop 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   15270
   End
   Begin VB.Label lblRight 
      BackStyle       =   0  'Transparent
      Height          =   11430
      Left            =   15315
      TabIndex        =   5
      Top             =   45
      Width           =   45
   End
   Begin VB.Label lblLeft 
      BackStyle       =   0  'Transparent
      Height          =   11430
      Left            =   0
      TabIndex        =   4
      Top             =   45
      Width           =   45
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   11040
      Width           =   255
   End
   Begin VB.Label lblContextHelp 
      BackColor       =   &H00000060&
      Caption         =   "Test "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7800
      TabIndex        =   3
      Top             =   9600
      Width           =   4095
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Not DEBUG_MODE Then Call ChangeRes(1680, 1050)
End
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirN
tmrScroll.Enabled = False
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim n As Integer
Dim collides As Boolean

n = activeUnits
If n < MAX_UNITS Then

   activeUnits = activeUnits + 1
   
   
   If DEBUG_MODE Then lblUnits.Caption = "Units: " & activeUnits
   
   
   unit(n).type = Int(Rnd * (2) + 1)
   unit(n).player = Int(Rnd * (2) + 1)
   unit(n).health = Int(Rnd * (unitType(unit(n).type).health) + 1)
   unit(n).moving = False
   unit(n).direction = Int(Rnd * (4))
   unit(n).frame = 1
   unit(n).selected = False
   unit(n).freezeFrame = False
   unit(n).targetUnit = -1
   
   For i = 0 To activeUnits - 1
      If unit(i).targetUnit = activeUnits - 1 Then
         unit(i).targetUnit = -1
         unit(i).combatMode = False
      End If
   Next i
   
   increment player(unit(n).player).population
   
   Do
      unit(n).location.x = Int(Rnd * (gameMap.dimensions.x * 48) + 1)
      unit(n).location.y = Int(Rnd * (gameMap.dimensions.y * 48) + 1)
      collides = False
      
      For i = 0 To activeUnits - 1
         If i <> n Then
            If collision(screenCoords(unit(n)), unitType(unit(n).type).dimensions, screenCoords(unit(i)), unitType(unit(i).type).dimensions) Then
               collides = True
            End If
         End If
      Next i
   Loop Until Not collides
   
   unit(n).target = unit(n).location
   unit(n).exploring = True
   needReExplore = True
   
End If

End Sub

Private Sub Command3_Click()
deleteUnits
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirN
tmrScroll.Enabled = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirN
tmrScroll.Enabled = False
End Sub

Private Sub lblContextHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirN
tmrScroll.Enabled = False
End Sub

Private Sub lblTop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirU
tmrScroll.Enabled = True
End Sub
Private Sub lblBottom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirD
tmrScroll.Enabled = True
End Sub
Private Sub lblLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirL
tmrScroll.Enabled = True
End Sub
Private Sub lblRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirR
tmrScroll.Enabled = True
End Sub

Private Sub lblTopRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirE
tmrScroll.Enabled = True
End Sub
Private Sub lblBottomRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirF
tmrScroll.Enabled = True
End Sub
Private Sub lblBottomLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirG
tmrScroll.Enabled = True
End Sub
Private Sub lblTopLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirH
tmrScroll.Enabled = True
End Sub

Private Sub picGame_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_CTRL Then ctrlDown = True
If KeyCode = KEY_DELETE Then deleteUnits
End Sub

Private Sub picGame_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_CTRL Then ctrlDown = False
End Sub

Private Sub Form_Load()
Dim x As Integer
Randomize

Debug.Print "==================================="
Debug.Print "          == Commands: =="
Debug.Print "printEntityList"
Debug.Print
init

If Not DEBUG_MODE Then Call ChangeRes(1024, 768)
Me.Show
picGame.SetFocus

If DEBUG_MODE Then
   lblCoords.Visible = True
   lblDisplacement.Visible = True
   lblMapCoords.Visible = True
   lblUnits.Visible = True
   shpExplore.Visible = True
   lblSelected.Visible = True
   lblTargetUnit.Visible = True
End If

gameLoop
updateStats
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DEBUG_MODE Then Call ChangeRes(1680, 1050)
End
End Sub

Private Sub picGame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

If Button = 2 Then 'RMB
   For i = 0 To activeUnits - 1
      If unit(i).selected Then
         If unit(i).player = you Then ' You can't move enemy units
            unit(i).target.x = x / Screen.TwipsPerPixelX + gameMap.displacement.x
            unit(i).target.y = y / Screen.TwipsPerPixelY + gameMap.displacement.y
            unit(i).targetUnit = findUnit(unit(i).target)
            unit(i).combatMode = False
            'If unit(i).targetUnit = i Then unit(i).targetUnit = -1
            If DEBUG_MODE Then lblTargetUnit = unit(i).targetUnit
            'If unit(i).targetUnit = -1 Then
               'unit(i).targetBuilding = findBuilding(unit(i).target)
            'End If
            unit(i).moving = True
         End If
      End If
   Next i

ElseIf Button = 1 Then 'LMB
   mouseDown = True
End If
End Sub

Private Sub picGame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim u As Integer

If DEBUG_MODE Then
   lblCoords.Caption = "Co-ords: (" & x / Screen.TwipsPerPixelX & ", " & y / Screen.TwipsPerPixelY & ")"
   lblMapCoords.Caption = "Map Co-ords: (" & x / Screen.TwipsPerPixelX + gameMap.displacement.x & ", " & y / Screen.TwipsPerPixelY + gameMap.displacement.y & ")"
End If

scrollDir = dirN

If Not mouseDown Then
   selectionRectangleLoc1.x = x / Screen.TwipsPerPixelX
   selectionRectangleLoc1.y = y / Screen.TwipsPerPixelY
End If
selectionRectangleLoc2.x = x / Screen.TwipsPerPixelX
selectionRectangleLoc2.y = y / Screen.TwipsPerPixelY

selectionRectangle = distance(selectionRectangleLoc1, selectionRectangleLoc2) > 2

drawEverything

writeContext ("")
If pointCollidesWithUnit(addCoords(makeCoords(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY), gameMap.displacement), u) Then
   writeContext IIf(DEBUG_MODE, "Unit " & u & ": ", "") & unitType(unit(u).type).name
   If DEBUG_MODE Then lblFrame.Caption = unit(u).frame
ElseIf DEBUG_MODE Then
   lblFrame.Caption = ""
End If

End Sub

Private Sub picGame_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

If mouseDown Then
   'For i = 0 To activeUnits - 1
   For i = activeUnits - 1 To 0 Step -1
      If ENEMIES_SELECTABLE Or unit(i).player = you Then
         If Not ctrlDown Then unit(i).selected = False 'unselect, unless CTRL is being pressed
         If (selectionRectangleLoc2.x + gameMap.displacement.x >= unit(i).location.x - unitType(unit(i).type).dimensions.x / 2 And _
             selectionRectangleLoc1.x + gameMap.displacement.x <= unit(i).location.x + unitType(unit(i).type).dimensions.x / 2) Or _
            (selectionRectangleLoc2.x + gameMap.displacement.x <= unit(i).location.x + unitType(unit(i).type).dimensions.x / 2 And _
             selectionRectangleLoc1.x + gameMap.displacement.x >= unit(i).location.x - unitType(unit(i).type).dimensions.x / 2) Then
            If (selectionRectangleLoc2.y + gameMap.displacement.y >= unit(i).location.y - unitType(unit(i).type).dimensions.y * (7 / 8) And _
                selectionRectangleLoc1.y + gameMap.displacement.y <= unit(i).location.y + unitType(unit(i).type).dimensions.y * (1 / 8)) Or _
               (selectionRectangleLoc2.y + gameMap.displacement.y <= unit(i).location.y + unitType(unit(i).type).dimensions.y * (1 / 8) And _
                selectionRectangleLoc1.y + gameMap.displacement.y >= unit(i).location.y - unitType(unit(i).type).dimensions.y * (7 / 8)) Then
               unit(i).selected = Not (unit(i).selected = True And ctrlDown)
               If Not selectionRectangle Then i = -1
            End If
         End If
      End If
   Next i
End If
mouseDown = False

i = getSelected
If i > -1 Then If unitType(unit(i).type).selectSound > -1 Then sound (unitType(unit(i).type).selectSound)

updateStats
End Sub

Private Sub writeContext(text As String)
lblContextHelp.Caption = text
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrCorpses_Timer()
Dim i As Integer
For i = 0 To activeCorpses - 1
   If corpse(i).timer <> -1 Then
      corpse(i).timer = corpse(i).timer - 1
      If corpse(i).timer = 0 Then deleteCorpse (i)
   End If
Next i
End Sub

Private Sub tmrScroll_Timer()
Select Case scrollDir
   Case dirU
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(0, -2))
   Case dirD
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(0, 2))
   Case dirL
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(-2, 0))
   Case dirR
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(2, 0))
   Case dirE
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(2, -2))
   Case dirF
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(2, 2))
   Case dirG
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(-2, 2))
   Case dirH
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(-2, -2))
End Select

If DEBUG_MODE Then lblDisplacement.Caption = "Displacement: (" & gameMap.displacement.x & ", " & gameMap.displacement.y & ")"
End Sub

Public Sub updateStats()
Dim u As typUnit
Dim t As typUnitType
Dim sel As Integer
sel = getSelected

If DEBUG_MODE Then lblSelected = sel

If sel >= 0 Then
   u = unit(sel)
   t = unitType(u.type)
   
   If DEBUG_MODE Then lblTargetUnit = unit(sel).targetUnit

   picPortrait.Visible = True
   lblType = t.name
   lblPlayer = player(u.player).name
   lblCiv = civ(player(u.player).civ).name
   lblHealth = u.health & "/" & t.health
   lblAttack = t.attack
   lblArmor = t.armor
   lblRange = t.range
   lblSpeed = t.speed
   lblHealing = t.healing
   
   If SHOW_UNUSED_STATS Then
      imgAttack.Visible = lblAttack.Visible = True
      imgArmor.Visible = lblArmor.Visible = True
      imgRange.Visible = lblRange.Visible = True
      imgHealing.Visible = lblHealing.Visible = True
      imgHealth.Visible = lblHealth.Visible = True
      imgSPeed.Visible = lblSpeed.Visible = True
   Else
      imgAttack.Visible = (t.attack > 0)
      lblAttack.Visible = (t.attack > 0)
      imgArmor.Visible = (t.armor > 0)
      lblArmor.Visible = (t.armor > 0)
      imgRange.Visible = (t.range > 0)
      lblRange.Visible = (t.range > 0)
      imgHealing.Visible = (t.healing > 0)
      lblHealing.Visible = (t.healing > 0)
      imgHealth.Visible = (t.health > 0)
      lblHealth.Visible = (t.health > 0)
      imgSPeed.Visible = (t.speed > 0)
      lblSpeed.Visible = (t.speed > 0)
   End If
Else
   imgAttack.Visible = False
   imgArmor.Visible = False
   imgRange.Visible = False
   imgHealing.Visible = False
   imgHealth.Visible = False
   imgSPeed.Visible = False
   picPortrait.Visible = False
   lblAttack = ""
   lblArmor = ""
   lblRange = ""
   lblHealing = ""
   lblHealth = ""
   lblSpeed = ""
   lblPlayer = ""
   lblCiv = ""
   lblType = ""
End If

frmGame.picPortrait.Cls
If sel >= 0 Then Call drawPortrait(unitType(unit(sel).type).portrait, unitType(unit(sel).type).portraitBackground)
frmGame.picPortrait.Refresh

End Sub
