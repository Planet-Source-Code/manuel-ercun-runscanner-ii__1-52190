VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "Form3"
   ScaleHeight     =   3735
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2520
      Top             =   1680
   End
   Begin Proyecto1.UserControl2 UserControl21 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
   End
   Begin VB.Image Image2 
      Height          =   1245
      Left            =   960
      Picture         =   "Form3.frx":0000
      Top             =   840
      Width           =   4395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando gráficos"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   3405
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   3735
      Left            =   0
      Picture         =   "Form3.frx":0721
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
test = True
Call formulario.Centro(Me, "center")

End Sub

Private Sub Timer1_Timer()
Form1.Show
Timer1.Enabled = False
End Sub
