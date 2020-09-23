VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Acerca"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":08CA
   ScaleHeight     =   4320
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image4 
      Height          =   255
      Index           =   1
      Left            =   1680
      Picture         =   "Form4.frx":3D41
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   3
      Left            =   1200
      Picture         =   "Form4.frx":3ED4
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   2
      Left            =   840
      Picture         =   "Form4.frx":4327
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   0
      Left            =   3480
      Picture         =   "Form4.frx":44C8
      Top             =   140
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   2835
      Left            =   360
      Picture         =   "Form4.frx":465B
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2820
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "Form4.frx":49BF
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   120
      Picture         =   "Form4.frx":5289
      Stretch         =   -1  'True
      Top             =   50
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ser As Boolean

Private Sub Form_Load()
Call formulario.Rng(Me, RGB(0, 0, 0))
Call formulario.Centro(Me, "center")
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image4(0).Picture = Image4(1).Picture
End Sub

Private Sub Image4_Click(Index As Integer)
Unload Me
End Sub

Private Sub Image4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
ser = True
Image4(0).Picture = Image4(3).Picture
End Sub

Private Sub Image4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If ser = False Then Image4(0).Picture = Image4(2).Picture
End Sub

Private Sub Image4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
ser = False
Image4(0).Picture = Image4(1).Picture
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call formulario.mover(Me)
End Sub
