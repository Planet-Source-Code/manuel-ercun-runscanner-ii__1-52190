VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H00C00000&
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ScaleHeight     =   3300
   ScaleWidth      =   4830
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   2280
      Picture         =   "UserControl1.ctx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   2040
      Picture         =   "UserControl1.ctx":0CCA
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   1080
      Picture         =   "UserControl1.ctx":134B
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   240
      Picture         =   "UserControl1.ctx":19F9
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   360
      Picture         =   "UserControl1.ctx":20E8
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Event click()

Dim ipicture As StdPicture
Dim pausa As Boolean

Public Property Get Picture() As StdPicture
Set Picture = ipicture
End Property
Public Property Set Picture(ByVal new_pic As StdPicture)
Set ipicture = new_pic
Image1(4).Picture = new_pic
End Property

Private Sub Image1_Click(Index As Integer)
RaiseEvent click
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
pausa = True
Image1(0).Picture = Image1(1).Picture
sndPlaySound ByVal App.Path & "\sonidos\Moving.wav", 0

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If pausa = False Then Image1(0).Picture = Image1(3).Picture
End Sub


Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
pausa = False
Image1(0).Picture = Image1(2).Picture
End Sub

Private Sub UserControl_Initialize()
UserControl.Tag = "g"
Image1(0).Move 0, 0
Image1(4).Move 200, 200

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Image1(0).Width
UserControl.Height = Image1(0).Height

End Sub

Public Sub desfocus()
Dim contro As Control
For Each contro In Form1.Controls
If contro.Tag <> "g" Then Image1(0).Picture = Image1(2).Picture
Next contro
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Picture", ipicture, Nothing)
End Sub
