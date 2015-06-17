VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   8016
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   8016
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

Private Sub Command1_Click()
gameLoop
End Sub

Private Sub Form_Load()

'Game = CreateCompatibleDC(0)
unitType(1).dc = CreateCompatibleDC(0)

caveman.speed = 1
unitType(1).dc = LoadGraphicDC(App.Path & "\Images\u001.bmp")

Dim x As Long

unitType(1).speed = 3

unit(1).location.x = 0
unit(1).location.y = 0
unit(1).target.x = 50
unit(1).target.y = 30
unit(1).type = 1

drawUnit unit(1)
'x = BitBlt(picGame.hdc, y, 0, 54, 56, caveman.dc, 0, 0, vbSrcCopy)
picGame.Refresh



End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
unit(1).location.x = unit(1).location.x + unitType(unit(1).type).speed
Dim x As Long
picGame.Cls
drawUnit unit(1)
'x = BitBlt(picGame.hdc, y, 0, 54, 56, caveman.dc, 0, 0, vbSrcCopy)
picGame.Refresh
End Sub
