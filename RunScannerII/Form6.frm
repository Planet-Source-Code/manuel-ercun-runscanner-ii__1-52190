VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Runscanner"
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":08CA
   ScaleHeight     =   6405
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   4200
      Top             =   2520
   End
   Begin Proyecto1.UserControl3 UserControl32 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "ADD"
   End
   Begin Proyecto1.UserControl3 UserControl31 
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "REMOVE"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form6.frx":6C83
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8493
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin Proyecto1.UserControl3 UserControl33 
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   4440
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "SAVE CHANGE"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Host:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   75
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3120
      Picture         =   "Form6.frx":6FD5
      Top             =   240
      Width           =   480
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
      Picture         =   "Form6.frx":72DF
      Top             =   195
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   0
      Left            =   6000
      Picture         =   "Form6.frx":749C
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   2
      Left            =   960
      Picture         =   "Form6.frx":762F
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   3
      Left            =   1320
      Picture         =   "Form6.frx":77D0
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   1
      Left            =   1680
      Picture         =   "Form6.frx":7C23
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ser As Boolean

Private Sub Form_Load()
On Error GoTo ema
Call formulario.Rng(Me, RGB(0, 0, 0))
Call formulario.Centro(Me, "center")
ListView1.ColumnHeaders.Add 1, , "HostName", 1500
ListView1.ColumnHeaders.Add 2, , "Comment", 3500
Set ListView1.SmallIcons = ImageList1
Text2 = Form1.Text1
Text1 = "IP Scanner"
load App.Path & "\" & "Data\favorites.dat"
ema:
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



Private Sub ListView1_Click()
Text2 = ListView1.SelectedItem
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
des
End Sub

Private Sub Timer1_Timer()
rotar Label2, "RUNSCANNER II"
End Sub

Private Sub des()
UserControl31.desfocus Me
UserControl32.desfocus Me
UserControl33.desfocus Me
End Sub

Private Sub UserControl31_click()
ListView1.ListItems.Remove ListView1.SelectedItem.Index
End Sub

Private Sub UserControl32_click()
Dim b As Variant
Set b = ListView1.ListItems.Add(, , Text2, , 1)
b.SubItems(1) = Text1
End Sub

Private Sub UserControl33_click()
Saveadd App.Path & "\" & "Data\favorites.dat"
End Sub

Private Sub Saveadd(ruta As String)
On Error Resume Next
Open ruta For Output As #2
For i = 1 To ListView1.ListItems.Count
Print #2, "e" & ListView1.ListItems.Item(i).Text & ";" & ListView1.ListItems.Item(i).ListSubItems(1).Text
Next i
Close #2
End Sub

Private Sub load(ruta As String)
On Error Resume Next
Dim g As String, j As Variant
Open ruta For Input As #1
Do
Line Input #1, g
v = Split(g, ";")
For i = LBound(v) To UBound(v)
If left(v(i), 1) = "e" Then Set j = ListView1.ListItems.Add(, , Mid(v(i), 2), , 1)
If left(v(i), 1) <> "e" Then j.SubItems(1) = v(i)

Next i


Loop Until EOF(1)


Close #1
End Sub
