VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   492
      Left            =   2880
      TabIndex        =   1
      Top             =   6000
      Width           =   612
   End
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
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
init
Me.Show
gameLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub picGame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then 'RMB
   For i = 0 To unitCount - 1
      If unit(i).selected Then
         unit(i).target.x = x / Screen.TwipsPerPixelX
         unit(i).target.y = y / Screen.TwipsPerPixelY
         unit(i).moving = True
      End If
   Next i

ElseIf Button = 1 Then 'LMB

   For i = 0 To unitCount - 1
      unit(i).selected = False
      If x / Screen.TwipsPerPixelX >= unit(i).location.x - unitType(unit(i).type).dimensions.x / 2 Then
         If x / Screen.TwipsPerPixelX <= unit(i).location.x + unitType(unit(i).type).dimensions.x / 2 Then
            If y / Screen.TwipsPerPixelY >= unit(i).location.y - unitType(unit(i).type).dimensions.y * (7 / 8) Then
               If y / Screen.TwipsPerPixelY <= unit(i).location.y - unitType(unit(i).type).dimensions.y * (1 / 8) Then
                  unit(i).selected = True
               End If
            End If
         End If
      End If
   Next i
      
End If
End Sub
