VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   15750
   ClientLeft      =   795
   ClientTop       =   0
   ClientWidth     =   25200
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   45
      TabIndex        =   17
      Top             =   13725
      Width           =   11175
      Begin VB.PictureBox picPortrait 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1875
         Left            =   2760
         ScaleHeight     =   1845
         ScaleWidth      =   1845
         TabIndex        =   29
         Top             =   30
         Width           =   1875
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   6960
         X2              =   8880
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblTarget 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Height          =   255
         Left            =   4680
         TabIndex        =   31
         Top             =   1560
         Width           =   6495
      End
      Begin VB.Label lblCiv 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Egyptians"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4680
         TabIndex        =   27
         Top             =   1080
         Width           =   6495
      End
      Begin VB.Label lblSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblHealing 
         BackStyle       =   0  'Transparent
         Caption         =   "3400/5000"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label lblRange 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblArmor 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgRange 
         Height          =   360
         Left            =   1440
         Picture         =   "Game.frx":000C
         Top             =   900
         Width           =   360
      End
      Begin VB.Image imgHealing 
         Height          =   360
         Left            =   60
         Picture         =   "Game.frx":108E
         Top             =   1440
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
         Top             =   900
         Width           =   360
      End
      Begin VB.Image imgHealth 
         Height          =   360
         Left            =   60
         Picture         =   "Game.frx":4214
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblAttack 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblHealth 
         BackStyle       =   0  'Transparent
         Caption         =   "3400/5000"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgArmor 
         Height          =   360
         Left            =   1440
         Picture         =   "Game.frx":5296
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pharaoh Rameses"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Top             =   840
         Width           =   6495
      End
      Begin VB.Label lblType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Axeman"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   4800
         TabIndex        =   18
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.Timer tmrScroll 
      Interval        =   8
      Left            =   18600
      Top             =   15000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Unit(s)"
      Height          =   375
      Left            =   14280
      TabIndex        =   2
      Top             =   14880
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create new unit"
      Height          =   375
      Left            =   12600
      TabIndex        =   1
      Top             =   14880
      Width           =   1695
   End
   Begin VB.PictureBox picGame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   13560
      Left            =   45
      ScaleHeight     =   13530
      ScaleWidth      =   25080
      TabIndex        =   0
      Top             =   90
      Width           =   25110
   End
   Begin VB.Timer tmrCorpses 
      Interval        =   1000
      Left            =   18120
      Top             =   14640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Resign"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   21240
      TabIndex        =   33
      Top             =   15240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   23040
      TabIndex        =   32
      Top             =   15240
      Width           =   1695
   End
   Begin VB.Label lblContextHelp 
      BackColor       =   &H00000000&
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
      Height          =   1455
      Left            =   21240
      TabIndex        =   30
      Top             =   13725
      Width           =   3735
   End
   Begin VB.Label lblTargetUnit 
      Caption         =   "Label1"
      Height          =   255
      Left            =   13920
      TabIndex        =   28
      Top             =   14280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSelected 
      Height          =   255
      Left            =   13080
      TabIndex        =   26
      Top             =   14280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblFrame 
      Height          =   255
      Left            =   17280
      TabIndex        =   16
      Top             =   14880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape shpExplore 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   12240
      Top             =   14880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUnits 
      Height          =   255
      Left            =   15120
      TabIndex        =   15
      Top             =   14880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblMapCoords 
      Height          =   255
      Left            =   16200
      TabIndex        =   14
      Top             =   15240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblDisplacement 
      Height          =   255
      Left            =   16200
      TabIndex        =   13
      Top             =   15000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblCoords 
      Height          =   255
      Left            =   16200
      TabIndex        =   12
      Top             =   14760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblBottomLeft 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   0
      TabIndex        =   11
      Top             =   15705
      Width           =   45
   End
   Begin VB.Label lblTopLeft 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblBottomRight 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   25155
      TabIndex        =   9
      Top             =   15705
      Width           =   45
   End
   Begin VB.Label lblTopRight 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   25155
      TabIndex        =   8
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblbottom 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   45
      TabIndex        =   7
      Top             =   15705
      Width           =   25110
   End
   Begin VB.Label lblTop 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   25110
   End
   Begin VB.Label lblRight 
      BackStyle       =   0  'Transparent
      Height          =   15660
      Left            =   25155
      TabIndex        =   4
      Top             =   45
      Width           =   45
   End
   Begin VB.Label lblLeft 
      BackStyle       =   0  'Transparent
      Height          =   15660
      Left            =   0
      TabIndex        =   3
      Top             =   45
      Width           =   45
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   24840
      TabIndex        =   5
      Top             =   15360
      Width           =   255
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call ChangeRes(1680, 1050)
End
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirN
tmrScroll.Enabled = False
End Sub

Private Sub Command2_Click()
frmGame.SetFocus

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
frmGame.SetFocus

deleteUnits
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirN
tmrScroll.Enabled = False
End Sub

Private Sub Label1_Click()
init
updateStats
End Sub

Private Sub lblContextHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
scrollDir = dirN
tmrScroll.Enabled = False
End Sub

Private Sub lblExit_Click()
End
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
If KeyCode = KEY_UP Then scrollDir = dirU
If KeyCode = KEY_DOWN Then scrollDir = dirD
If KeyCode = KEY_LEFT Then scrollDir = dirL
If KeyCode = KEY_RIGHT Then scrollDir = dirR

End Sub

Private Sub picGame_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_CTRL Then ctrlDown = False
If KeyCode >= KEY_LEFT And KeyCode <= KEY_DOWN Then scrollDir = dirN
End Sub

Private Sub Form_Load()
Dim x As Integer
Randomize
init

If Not DEBUG_MODE Then Call ChangeRes(1680, 1050)
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
   For i = 0 To activeUnits - 1
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
Dim distance As Integer
distance = 5
Select Case scrollDir
   Case dirU
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(0, -1 * distance))
   Case dirD
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(0, distance))
   Case dirL
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(-1 * distance, 0))
   Case dirR
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(distance, 0))
   Case dirE
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(distance, -1 * distance))
   Case dirF
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(distance, distance))
   Case dirG
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(-1 * distance, distance))
   Case dirH
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(-1 * distance, -1 * 2))
End Select

If DEBUG_MODE Then lblDisplacement.Caption = "Displacement: (" & gameMap.displacement.x & ", " & gameMap.displacement.y & ")"
End Sub

Public Sub updateStats()
Dim u As typUnit
Dim t As typUnitType
Dim sel As Integer
sel = getSelected

lblTarget = ""
Line1.Visible = False

If DEBUG_MODE Then lblSelected = sel

If sel >= 0 Then
   u = unit(sel)
   t = unitType(u.type)
   
   If DEBUG_MODE Then lblTargetUnit = unit(sel).targetUnit
   
   If unit(sel).targetUnit > -1 Then
      lblTarget = "Current target: " & unitType(unit(unit(sel).targetUnit).type).name
      Line1.Visible = True
   End If

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
      imgAttack.Visible = True
      lblAttack.Visible = True
      imgArmor.Visible = True
      lblArmor.Visible = True
      imgRange.Visible = True
      lblRange.Visible = True
      imgHealing.Visible = True
      lblHealing.Visible = True
      imgHealth.Visible = True
      lblHealth.Visible = True
      imgSPeed.Visible = True
      lblSpeed.Visible = True
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
