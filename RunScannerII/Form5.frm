VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "RunScanners II"
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":08CA
   ScaleHeight     =   6195
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000007&
      Caption         =   "Scanner"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1800
      TabIndex        =   21
      Top             =   4680
      Width           =   4455
      Begin VB.CheckBox Check1 
         Caption         =   "UDP"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   23
         Top             =   360
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TCP"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000007&
      Caption         =   "Sound"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1800
      TabIndex        =   17
      Top             =   3720
      Width           =   4455
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3840
         TabIndex        =   26
         Text            =   "135"
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Value"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000007&
      Caption         =   "Types Options"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1800
      TabIndex        =   11
      Top             =   2760
      Width           =   4455
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   15
         Text            =   "4"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sec"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Wait for port:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "Types"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1800
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "Only port"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1 to 65526"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Port List"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Options"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      Begin VB.CheckBox Check1 
         Caption         =   "Banners"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "O.S."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "DNS"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin Proyecto1.UserControl3 UserControl32 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "ACCEPT"
   End
   Begin Proyecto1.UserControl3 UserControl31 
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "CANCEL"
   End
   Begin Proyecto1.UserControl3 UserControl33 
      Height          =   615
      Left            =   360
      TabIndex        =   24
      Top             =   3960
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "SOUND"
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3120
      Picture         =   "Form5.frx":6C83
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
      Picture         =   "Form5.frx":6F8D
      Top             =   195
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   0
      Left            =   6000
      Picture         =   "Form5.frx":714A
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   1
      Left            =   1680
      Picture         =   "Form5.frx":72DD
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   3
      Left            =   1320
      Picture         =   "Form5.frx":7470
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   2
      Left            =   960
      Picture         =   "Form5.frx":78C3
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ser As Boolean



Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 2
If Check1(2).Value = Unchecked Then Check1(5).Value = vbUnchecked
End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
des
End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
des
End Sub



Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
des
End Sub

Private Sub Image4_Click(Index As Integer)
Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
formulario.mover Me
End Sub
Private Sub Form_Load()
On Error Resume Next
Call formulario.Rng(Me, RGB(0, 0, 0))
Call formulario.Centro(Me, "center")
Frame1.BackColor = RGB(50, 50, 50)
Frame2.BackColor = RGB(50, 50, 50)
Frame3.BackColor = RGB(50, 50, 50)
Frame4.BackColor = RGB(50, 50, 50)
Frame5.BackColor = RGB(50, 50, 50)
For i = Check1.lbound To Check1.ubound
Check1(i).BackColor = RGB(50, 50, 50)
Next i
For i = Option1.lbound To Option1.ubound
Option1(i).BackColor = RGB(50, 50, 50)
Next i
Text3 = App.Path & "\" & "Sonidos\Blip.wav"
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

Private Sub UserControl31_click()
Unload Me
End Sub

Private Sub des()
UserControl31.desfocus Me
UserControl32.desfocus Me
UserControl33.desfocus Me
End Sub

Private Sub UserControl32_click()
Me.Hide
End Sub

Private Sub UserControl33_click()

On Error GoTo ema
With CommonDialog1
 .DefaultExt = "WAV"
 .DialogTitle = "Open sound..."
 .CancelError = True
 .Filter = "All Wave(*.wav)|*.wav"
 .ShowOpen
 If Len(.FileName) = 0 Then Exit Sub
 Text3 = .FileName
End With



ema:
End Sub
