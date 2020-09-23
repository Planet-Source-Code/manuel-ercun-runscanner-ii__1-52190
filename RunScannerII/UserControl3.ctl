VERSION 5.00
Begin VB.UserControl UserControl3 
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   ScaleHeight     =   660
   ScaleWidth      =   1830
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000029&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      Height          =   615
      Left            =   1200
      ScaleHeight     =   615
      ScaleWidth      =   1230
      TabIndex        =   0
      Top             =   840
      Width           =   1235
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000000&
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Command"
         ForeColor       =   &H00FFC0C0&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   1220
         X2              =   1220
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   1200
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   10
         X2              =   10
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   1200
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   4
         Visible         =   0   'False
         X1              =   0
         X2              =   1200
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   5
         Visible         =   0   'False
         X1              =   1180
         X2              =   1180
         Y1              =   0
         Y2              =   600
      End
   End
End
Attribute VB_Name = "UserControl3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Event click()

Dim iforecolor As OLE_COLOR
Dim icaption As String


Public Property Get ForeColor() As OLE_COLOR
ForeColor = iforecolor
End Property
Public Property Let ForeColor(ByVal new_fore As OLE_COLOR)
iforecolor = new_fore
Label1.ForeColor = new_fore
PropertyChanged "ForeColor"
End Property


Public Property Get Caption() As String
Caption = icaption
End Property
Public Property Let Caption(ByVal new_caption As String)
icaption = new_caption
Label1.Caption = new_caption
PropertyChanged "Caption"
End Property

Private Sub Label1_Click()
Call Picture1_Click
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 picture1_MouseDown 0, 0, 0, 0
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 picture1_MouseUp 0, 0, 0, 0
End Sub

Private Sub Picture1_Click()
RaiseEvent click
End Sub

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 1

sndPlaySound ByVal App.Path & "\sonidos\Moving.wav", 0

End Sub

Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = Line1.LBound To Line1.UBound
Line1(i).Visible = True
Next i
Shape2.Visible = True

End Sub

Private Sub picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
End Sub

Private Sub UserControl_Initialize()
Dim i As Integer
Label1.ForeColor = vbBlue
Picture1.BackColor = RGB(50, 50, 50)
Picture1.Move 0, 0
For i = Line1.LBound To Line1.UBound
Line1(i).Visible = False
Next i
Shape2.Visible = False
UserControl.Tag = "0"
icaption = "Command"
End Sub

Public Sub desfocus(Form As Form)
Dim contro As Control
For Each contro In Form.Controls
If contro.Tag <> "0" Then
For i = Line1.LBound To Line1.UBound
Line1(i).Visible = False
Next i
Shape2.Visible = False
End If
Next contro
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Caption = PropBag.ReadProperty("Caption", "Command")
ForeColor = PropBag.ReadProperty("ForeColor", vbBlue)
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Picture1.Width
UserControl.Height = Picture1.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", icaption, "Command")
Call PropBag.WriteProperty("ForeColor", iforecolor, vbBlue)
End Sub
