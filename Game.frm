VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   795
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScroll 
      Interval        =   8
      Left            =   8880
      Top             =   8040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Unit(s)"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create new unit"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   8040
      Width           =   1695
   End
   Begin VB.PictureBox picGame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   7080
      Left            =   45
      ScaleHeight     =   7050
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   255
      Width           =   11910
   End
   Begin VB.Label lblUnits 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblMapCoords 
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblDisplacement 
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   7680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblCoords 
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   7440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblBottomLeft 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   0
      TabIndex        =   12
      Top             =   8955
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
      Left            =   11955
      TabIndex        =   10
      Top             =   8955
      Width           =   45
   End
   Begin VB.Label lblTopRight 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   11955
      TabIndex        =   9
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblbottom 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   45
      TabIndex        =   8
      Top             =   8955
      Width           =   11910
   End
   Begin VB.Label lblTop 
      BackStyle       =   0  'Transparent
      Height          =   45
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   11910
   End
   Begin VB.Label lblRight 
      BackStyle       =   0  'Transparent
      Height          =   8910
      Left            =   11955
      TabIndex        =   5
      Top             =   45
      Width           =   45
   End
   Begin VB.Label lblLeft 
      BackStyle       =   0  'Transparent
      Height          =   8910
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
      Left            =   11760
      TabIndex        =   6
      Top             =   0
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
      Left            =   1710
      TabIndex        =   3
      Top             =   7350
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
Call ChangeRes(1680, 1050)
End
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirN
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim n As Integer
Dim collides As Boolean

n = activeUnits
activeUnits = activeUnits + 1

If DEBUG_MODE Then lblUnits.Caption = "Units: " & activeUnits


unit(n).type = Int(Rnd * (2) + 1)
unit(n).moving = False
unit(n).direction = Int(Rnd * (3))
unit(n).frame = 1
unit(n).selected = False
unit(n).freezeFrame = False

Do
   unit(n).location.X = Int(Rnd * (gameMap.dimensions.X * 48) + 1)
   unit(n).location.Y = Int(Rnd * (gameMap.dimensions.Y * 48) + 1)
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

End Sub

Private Sub Command3_Click()
deleteUnits
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirN
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirN
End Sub

Private Sub lblContextHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirN
End Sub

Private Sub lblTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirU
End Sub
Private Sub lblBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirD
End Sub
Private Sub lblLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirL
End Sub
Private Sub lblRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirR
End Sub

Private Sub lblTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirE
End Sub
Private Sub lblBottomRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirF
End Sub
Private Sub lblBottomLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirG
End Sub
Private Sub lblTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
scrollDir = dirH
End Sub

Private Sub picGame_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_CTRL Then ctrlDown = True
If KeyCode = KEY_DELETE Then deleteUnits
End Sub

Private Sub picGame_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_CTRL Then ctrlDown = False
End Sub

Private Sub Form_Load()
init

If Not DEBUG_MODE Then Call ChangeRes(800, 600)
Me.Show
picGame.SetFocus

If DEBUG_MODE Then
   lblCoords.Visible = True
   lblDisplacement.Visible = True
   lblMapCoords.Visible = True
   lblUnits.Visible = True
End If

gameLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DEBUG_MODE Then Call ChangeRes(1680, 1050)
End
End Sub

Private Sub picGame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

If Button = 2 Then 'RMB
   For i = 0 To activeUnits - 1
      If unit(i).selected Then
         unit(i).target.X = X / Screen.TwipsPerPixelX + gameMap.displacement.X
         unit(i).target.Y = Y / Screen.TwipsPerPixelY + gameMap.displacement.Y
         unit(i).moving = True
      End If
   Next i

ElseIf Button = 1 Then 'LMB
   mouseDown = True
End If
End Sub

Private Sub picGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim u As Integer

If DEBUG_MODE Then
   lblCoords.Caption = "Co-ords: (" & X / Screen.TwipsPerPixelX & ", " & Y / Screen.TwipsPerPixelY & ")"
   lblMapCoords.Caption = "Map Co-ords: (" & X / Screen.TwipsPerPixelX + gameMap.displacement.X & ", " & Y / Screen.TwipsPerPixelY + gameMap.displacement.Y & ")"
End If

scrollDir = dirN

If Not mouseDown Then
   selectionRectangleLoc1.X = X / Screen.TwipsPerPixelX
   selectionRectangleLoc1.Y = Y / Screen.TwipsPerPixelY
End If
selectionRectangleLoc2.X = X / Screen.TwipsPerPixelX
selectionRectangleLoc2.Y = Y / Screen.TwipsPerPixelY

drawEverything

writeContext ("")
If pointCollidesWithUnit(addCoords(makeCoords(X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY), gameMap.displacement), u) Then
   writeContext IIf(DEBUG_MODE, "Unit " & u & ": ", "") & unitType(unit(u).type).name
End If

End Sub

Private Sub picGame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

If mouseDown Then
   For i = 0 To activeUnits - 1
      If Not ctrlDown Then unit(i).selected = False 'unselect, unless CTRL is being pressed
      If (selectionRectangleLoc2.X + gameMap.displacement.X >= unit(i).location.X - unitType(unit(i).type).dimensions.X / 2 And _
          selectionRectangleLoc1.X + gameMap.displacement.X <= unit(i).location.X + unitType(unit(i).type).dimensions.X / 2) Or _
         (selectionRectangleLoc2.X + gameMap.displacement.X <= unit(i).location.X + unitType(unit(i).type).dimensions.X / 2 And _
          selectionRectangleLoc1.X + gameMap.displacement.X >= unit(i).location.X - unitType(unit(i).type).dimensions.X / 2) Then
         If (selectionRectangleLoc2.Y + gameMap.displacement.Y >= unit(i).location.Y - unitType(unit(i).type).dimensions.Y * (7 / 8) And _
             selectionRectangleLoc1.Y + gameMap.displacement.Y <= unit(i).location.Y + unitType(unit(i).type).dimensions.Y * (1 / 8)) Or _
            (selectionRectangleLoc2.Y + gameMap.displacement.Y <= unit(i).location.Y + unitType(unit(i).type).dimensions.Y * (1 / 8) And _
             selectionRectangleLoc1.Y + gameMap.displacement.Y >= unit(i).location.Y - unitType(unit(i).type).dimensions.Y * (7 / 8)) Then
            unit(i).selected = Not (unit(i).selected = True And ctrlDown)
         End If
      End If
   Next i
End If
mouseDown = False
End Sub

Private Sub writeContext(text As String)
lblContextHelp.Caption = text
End Sub

Private Sub tmrScroll_Timer()
Select Case scrollDir
   Case dirU
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(0, -1))
   Case dirD
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(0, 1))
   Case dirL
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(-1, 0))
   Case dirR
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(1, 0))
   Case dirE
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(1, -1))
   Case dirF
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(1, 1))
   Case dirG
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(-1, 1))
   Case dirH
      gameMap.displacement = addCoords(gameMap.displacement, makeCoords(-1, -1))
End Select

If DEBUG_MODE Then lblDisplacement.Caption = "Displacement: (" & gameMap.displacement.X & ", " & gameMap.displacement.Y & ")"
End Sub
