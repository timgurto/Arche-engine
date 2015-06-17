VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   8016
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   668
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   612
      Left            =   3840
      TabIndex        =   1
      Top             =   6000
      Width           =   732
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6840
      Top             =   2760
   End
   Begin VB.PictureBox picGame 
      AutoRedraw      =   -1  'True
      Height          =   5172
      Left            =   720
      ScaleHeight     =   5124
      ScaleWidth      =   5844
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
   unit(1).target.x = x / TWIPS_PER_PIXEL - unitType(unit(1).type).dimensions.x / 2
   unit(1).target.y = y / TWIPS_PER_PIXEL - unitType(unit(1).type).dimensions.y / 2
   unit(1).moving = True
End If
End Sub
