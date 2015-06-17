VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGame 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C000&
      Height          =   5172
      Left            =   720
      ScaleHeight     =   5115
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   480
      Width           =   5892
   End
   Begin VB.Label selectiony2 
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label selectionx2 
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label selectiony1 
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label selectionx1 
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label mouseDownIndicator 
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub picGame_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_CTRL Then ctrlDown = True
End Sub

Private Sub picGame_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_CTRL Then ctrlDown = False
End Sub

Private Sub Form_Load()
init
Me.Show
gameLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub picGame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then 'RMB
   For i = 0 To unitCount - 1
      If unit(i).selected Then
         unit(i).target.X = X / Screen.TwipsPerPixelX
         unit(i).target.Y = Y / Screen.TwipsPerPixelY
         unit(i).moving = True
      End If
   Next i

ElseIf Button = 1 Then 'LMB

   mouseDown = True: frmGame.mouseDownIndicator = mouseDown

End If
End Sub

Private Sub picGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not mouseDown Then
   selectionRectangleLoc1.X = X / Screen.TwipsPerPixelX
   selectionRectangleLoc1.Y = Y / Screen.TwipsPerPixelY
End If
selectionRectangleLoc2.X = X / Screen.TwipsPerPixelX
selectionRectangleLoc2.Y = Y / Screen.TwipsPerPixelY

frmGame.selectionx1 = selectionRectangleLoc1.X
frmGame.selectionx2 = selectionRectangleLoc2.X
frmGame.selectiony1 = selectionRectangleLoc1.Y
frmGame.selectiony2 = selectionRectangleLoc2.Y

drawEverything

End Sub

Private Sub picGame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mouseDown Then
   For i = 0 To unitCount - 1
      If Not ctrlDown Then unit(i).selected = False 'unselect, unless CTRL is being pressed
      If (selectionRectangleLoc2.X >= unit(i).location.X - unitType(unit(i).type).dimensions.X / 2 And _
          selectionRectangleLoc1.X <= unit(i).location.X + unitType(unit(i).type).dimensions.X / 2) Or _
         (selectionRectangleLoc2.X <= unit(i).location.X + unitType(unit(i).type).dimensions.X / 2 And _
          selectionRectangleLoc1.X >= unit(i).location.X - unitType(unit(i).type).dimensions.X / 2) Then
         If (selectionRectangleLoc2.Y >= unit(i).location.Y - unitType(unit(i).type).dimensions.Y * (7 / 8) And _
             selectionRectangleLoc1.Y <= unit(i).location.Y + unitType(unit(i).type).dimensions.Y * (1 / 8)) Or _
            (selectionRectangleLoc2.Y <= unit(i).location.Y + unitType(unit(i).type).dimensions.Y * (7 / 8) And _
             selectionRectangleLoc1.Y >= unit(i).location.Y - unitType(unit(i).type).dimensions.Y * (1 / 8)) Then
            unit(i).selected = Not (unit(i).selected = True And ctrlDown)
         End If
      End If
   Next i
End If

mouseDown = False: frmGame.mouseDownIndicator = mouseDown
End Sub
