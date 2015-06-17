VERSION 5.00
Begin VB.Form Game 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   8016
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   8016
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
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
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'Game = CreateCompatibleDC(0)
caveman.dc = CreateCompatibleDC(0)

caveman.speed = 1
caveman.dc = LoadGraphicDC(App.Path & "\Images\u001.bmp")

'Dim og As typUnit
Dim x As Long
'
'og.location.x = 0
'og.location.y = 0
'og.target.x = 50
'og.target.y = 30
y = 0

x = BitBlt(picGame.hdc, y, 0, 54, 56, caveman.dc, 0, 0, vbSrcCopy)
picGame.Refresh


End Sub

Private Sub Timer1_Timer()
y = y + 3
Dim x As Long
picGame.Cls
x = BitBlt(picGame.hdc, y, 0, 54, 56, caveman.dc, 0, 0, vbSrcCopy)
picGame.Refresh
End Sub
