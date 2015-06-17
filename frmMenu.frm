VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arche Engine"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4800
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   4380
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Game"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   960
      TabIndex        =   2
      Top             =   5880
      Width           =   2775
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "GAMES"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1470
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ARCHON"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   1080
         Picture         =   "frmMenu.frx":08CA
         Stretch         =   -1  'True
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Arche Engine v.12.1.5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a game to load:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
gamePath = File1.path & "\" & File1.FileName
Me.Hide
frmGame.Show
End Sub

Private Sub File1_Click()
Dim s As String
Command1.Enabled = True
Open File1.path & "\" & File1.FileName For Input As #1
While Not EOF(1)
   Input #1, s
   Text1.text = s
Wend
Close #1
End Sub

Private Sub File1_DblClick()
Call Command1_Click
End Sub

Private Sub Form_Load()
File1.path = App.path & "\Games"
End Sub

Private Sub Frame1_Click()
MsgBox "(C) 2006-2009 Tim Gurto"
End Sub

Private Sub Image1_Click()
Call Frame1_Click
End Sub

Private Sub Label2_Click()
Call Frame1_Click
End Sub

Private Sub Label3_Click()
Call Frame1_Click
End Sub

Private Sub Label4_Click()
Call Frame1_Click
End Sub
