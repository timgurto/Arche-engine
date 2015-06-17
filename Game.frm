VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Create new unit"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   7320
      Width           =   1695
   End
   Begin VB.PictureBox picGame 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7275
      ScaleWidth      =   11940
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   255
         Left            =   11760
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
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

Private Sub Command2_Click()
Dim n As Integer
n = activeUnits
activeUnits = activeUnits + 1

unit(n).location.x = Int(Rnd * (800) + 1)
unit(n).location.y = Int(Rnd * (481) + 1)

Dim collides As Boolean
Do
collides = False

For i = 0 To activeUnits - 1
   If i <> n Then
      If collision(screenCoords(unit(n)), unitType(unit(n).type).dimensions, screenCoords(unit(i)), unitType(unit(i).type).dimensions) Then
         collides = True
      End If
   End If
Next i
Loop Until Not collides

unit(n).type = 2
unit(n).moving = False
unit(n).direction = Int(Rnd * (3))
unit(n).frame = 1
unit(n).selected = False
unit(n).target = unit(n).location
unit(n).freezeFrame = False

End Sub

Private Sub picGame_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_CTRL Then ctrlDown = True
End Sub

Private Sub picGame_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_CTRL Then ctrlDown = False
End Sub

Private Sub Form_Load()
init

Call ChangeRes(800, 600)
Me.Show
picGame.SetFocus

gameLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ChangeRes(1680, 1050)
End
End Sub

Private Sub picGame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then 'RMB
   For i = 0 To activeUnits - 1
      If unit(i).selected Then
         unit(i).target.x = x / Screen.TwipsPerPixelX
         unit(i).target.y = y / Screen.TwipsPerPixelY
         unit(i).moving = True
      End If
   Next i

ElseIf Button = 1 Then 'LMB
   mouseDown = True
End If
End Sub

Private Sub picGame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not mouseDown Then
   selectionRectangleLoc1.x = x / Screen.TwipsPerPixelX
   selectionRectangleLoc1.y = y / Screen.TwipsPerPixelY
End If
selectionRectangleLoc2.x = x / Screen.TwipsPerPixelX
selectionRectangleLoc2.y = y / Screen.TwipsPerPixelY

drawEverything

End Sub

Private Sub picGame_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If mouseDown Then
   For i = 0 To activeUnits - 1
      If Not ctrlDown Then unit(i).selected = False 'unselect, unless CTRL is being pressed
      If (selectionRectangleLoc2.x >= unit(i).location.x - unitType(unit(i).type).dimensions.x / 2 And _
          selectionRectangleLoc1.x <= unit(i).location.x + unitType(unit(i).type).dimensions.x / 2) Or _
         (selectionRectangleLoc2.x <= unit(i).location.x + unitType(unit(i).type).dimensions.x / 2 And _
          selectionRectangleLoc1.x >= unit(i).location.x - unitType(unit(i).type).dimensions.x / 2) Then
         If (selectionRectangleLoc2.y >= unit(i).location.y - unitType(unit(i).type).dimensions.y * (7 / 8) And _
             selectionRectangleLoc1.y <= unit(i).location.y + unitType(unit(i).type).dimensions.y * (1 / 8)) Or _
            (selectionRectangleLoc2.y <= unit(i).location.y + unitType(unit(i).type).dimensions.y * (7 / 8) And _
             selectionRectangleLoc1.y >= unit(i).location.y - unitType(unit(i).type).dimensions.y * (1 / 8)) Then
            unit(i).selected = Not (unit(i).selected = True And ctrlDown)
         End If
      End If
   Next i
End If
mouseDown = False
End Sub
