VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form8 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "RunScanner"
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":08CA
   ScaleHeight     =   6225
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Options Ping"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   5895
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3480
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Text            =   "32"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Text            =   "7"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Text            =   "4"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3240
         TabIndex        =   10
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packet:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Resquest:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1110
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   1080
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
            Picture         =   "Form8.frx":6C83
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Proyecto1.UserControl3 UserControl32 
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "PING"
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
      Picture         =   "Form8.frx":6F9D
      Top             =   195
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3120
      Picture         =   "Form8.frx":715A
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   0
      Left            =   6000
      Picture         =   "Form8.frx":7464
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   1
      Left            =   2040
      Picture         =   "Form8.frx":75F7
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   3
      Left            =   1680
      Picture         =   "Form8.frx":778A
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   2
      Left            =   1320
      Picture         =   "Form8.frx":7BDD
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ser As Boolean

Private Sub Form_Load()
On Error Resume Next
Call formulario.Rng(Me, RGB(0, 0, 0))
Call formulario.Centro(Me, "center")
Frame1.BackColor = RGB(50, 50, 50)
Set ListView1.SmallIcons = ImageList1
ListView1.View = lvwReport
ListView1.ColumnHeaders.Add 1, , "HostName", 2500
ListView1.ColumnHeaders.Add 2, , "Bytes", 800
ListView1.ColumnHeaders.Add 3, , "Time", 800
ListView1.ColumnHeaders.Add 4, , "TTL", 800
ListView1.ColumnHeaders.Add 5, , "Resquest", 800

Text4 = Form1.Text1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl32.desfocus Me
End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl32.desfocus Me
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

Private Sub UserControl32_click()
On Error Resume Next
Dim s As Variant
ListView1.ListItems.Clear
v = Split(ping(Trim(Text4), Text1, Text2), ";")
For i = LBound(v) To UBound(v)
If left(v(i), 1) = "i" Then Set s = ListView1.ListItems.Add(, , Mid(v(i), 2), , 1)
If left(v(i), 1) = "t" Then s.SubItems(2) = Mid(v(i), 2)
If left(v(i), 1) = "l" Then s.SubItems(3) = Mid(v(i), 2)
s.SubItems(1) = Text3
s.SubItems(4) = Text1
Next i
End Sub
