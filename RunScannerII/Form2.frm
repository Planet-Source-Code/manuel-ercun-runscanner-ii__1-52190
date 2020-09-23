VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "RunScannersII"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":08CA
   ScaleHeight     =   6390
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   4560
      Top             =   1920
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Important port"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "It´s a Trojan port"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8493
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483644
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Proyecto1.UserControl3 UserControl31 
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   3960
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "REMOVE PORT"
   End
   Begin Proyecto1.UserControl3 UserControl32 
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "ADD PORT"
   End
   Begin Proyecto1.UserControl3 UserControl33 
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   4800
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "SAVE CHANGE"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":6C83
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":6FD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":7327
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   4720
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TCP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   3930
      TabIndex        =   12
      Top             =   360
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UDP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   4800
      TabIndex        =   11
      Top             =   360
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3120
      Picture         =   "Form2.frx":7679
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port Description:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port Number:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   930
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   2
      Left            =   840
      Picture         =   "Form2.frx":7983
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   3
      Left            =   1200
      Picture         =   "Form2.frx":7B24
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   1
      Left            =   1560
      Picture         =   "Form2.frx":7F77
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   0
      Left            =   6000
      Picture         =   "Form2.frx":810A
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   0
      Left            =   120
      Picture         =   "Form2.frx":829D
      Top             =   195
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ser As Boolean
Dim r As Variant
Dim cas As Boolean

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Check2.Value = vbUnchecked
End Sub



Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Check1.Value = vbUnchecked
End Sub

Private Sub Form_Load()
On Error Resume Next
Call formulario.Rng(Me, RGB(0, 0, 0))
Call formulario.Centro(Me, "center")
Check1.BackColor = RGB(50, 50, 50)

ser = False
Set ListView1.SmallIcons = ImageList1
ListView1.View = lvwReport
ListView1.ColumnHeaders.Add 1, , "PORT", 1200
ListView1.ColumnHeaders.Add 2, , "DESCRIPTIONS", 5000
Puertos App.Path & "\Data\porttcp.dat", ListView1
cas = False
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
des
End Sub

Private Sub Image4_Click(Index As Integer)
Unload Me
End Sub

Private Sub Image4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ser = True
Image4(0).Picture = Image4(3).Picture
End Sub

Private Sub Image4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ser = False Then Image4(0).Picture = Image4(2).Picture

End Sub

Private Sub Image4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ser = False
 Image4(0).Picture = Image4(1).Picture
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
formulario.mover Me
End Sub



Private Sub des()
UserControl31.desfocus Me
UserControl32.desfocus Me
UserControl33.desfocus Me
End Sub

Private Sub Label4_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
ListView1.ListItems.Clear
Puertos App.Path & "\Data\porttcp.dat", ListView1
cas = False
Case 1
ListView1.ListItems.Clear
Puertos App.Path & "\Data\portudp.dat", ListView1
cas = True
End Select
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4(Index).ForeColor = vbRed
End Sub

Private Sub Label4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4(Index).ForeColor = vbBlue
End Sub



Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
des
End Sub




Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub
If IsNumeric(Chr(KeyAscii)) = False Then
KeyAscii = 0
Else
DoEvents
End If
End Sub

Private Sub UserControl31_click()
ListView1.ListItems.Remove ListView1.SelectedItem.Index
End Sub

Private Sub UserControl32_click()
On Error Resume Next
If Text1 <> "" Then
If Check1.Value = vbChecked Then Set r = ListView1.ListItems.Add(, , Trim(Text1), , 2)
If Check2.Value = vbChecked Then Set r = ListView1.ListItems.Add(, , Trim(Text1), , 3)
If Check1.Value = vbUnchecked And Check2.Value = vbUnchecked Then Set r = ListView1.ListItems.Add(, , Trim(Text1), , 1)
r.SubItems(1) = Text2
r.Checked = True
End If
End Sub

Private Sub UserControl33_click()
If cas = False Then Saveport App.Path & "\data\porttcp.dat", ListView1
If cas = True Then Saveport App.Path & "\data\portudp.dat", ListView1
End Sub
