VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "RunScannersII"
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   6825
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock3 
      Index           =   0
      Left            =   3240
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   3360
      Top             =   2640
   End
   Begin VB.TextBox Text2 
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
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5370
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   2880
      Top             =   3120
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   3000
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   480
      Picture         =   "Form1.frx":7083
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   6120
      Width           =   480
   End
   Begin Proyecto1.UserControl3 UserControl34 
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   3000
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "CONFIG"
   End
   Begin Proyecto1.UserControl3 UserControl33 
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "OPTIONS"
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   7
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   3600
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Proyecto1.UserControl3 UserControl32 
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "STOP"
   End
   Begin Proyecto1.UserControl3 UserControl31 
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "START"
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4455
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7858
      _Version        =   393217
      HideSelection   =   0   'False
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin Proyecto1.UserControl1 UserControl11 
      Height          =   735
      Index           =   2
      Left            =   3840
      TabIndex        =   3
      Top             =   5880
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Picture         =   "Form1.frx":794D
   End
   Begin Proyecto1.UserControl1 UserControl11 
      Height          =   735
      Index           =   0
      Left            =   5520
      TabIndex        =   1
      Top             =   5880
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Picture         =   "Form1.frx":8227
   End
   Begin Proyecto1.UserControl1 UserControl11 
      Height          =   735
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   5880
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Picture         =   "Form1.frx":8F01
   End
   Begin Proyecto1.UserControl3 UserControl35 
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   3720
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "FAVORITES"
   End
   Begin Proyecto1.UserControl3 UserControl36 
      Height          =   615
      Left            =   480
      TabIndex        =   11
      Top             =   4440
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1085
      Caption         =   "PING"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":97DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9B2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9E7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A1D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A523
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A875
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ABC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AF19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B2DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B631
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B9E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BD35
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C073
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C3C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C717
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CA69
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CD83
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D0D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D3EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D709
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3120
      Picture         =   "Form1.frx":DCA3
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
      Picture         =   "Form1.frx":DFAD
      Top             =   195
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents shell_tray As Tray
Attribute shell_tray.VB_VarHelpID = -1

Dim systemtray As New Tray

Dim j As Integer, h As Integer

Private Sub Form_Load()
On Error Resume Next

Call formulario.Rng(Me, RGB(0, 0, 0))
Call formulario.Centro(Me, "center")
Unload Form3
systemtray.ToolTipText = "RunScannerII For ErcUn"
Set systemtray.Pictures = Picture1
Set TreeView1.ImageList = ImageList1
salir = False
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
desenfocar UserControl11.Count - 1
des
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then systemtray.ShellAdd Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
For i = 1 To cuenta
Unload Winsock1(i)
Next i
For i = 1 To cuenta1
Unload Winsock2(i)
Next i
For i = 1 To cuenta2
Unload Winsock3(i)
Next i
End
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then formulario.mover Me
End Sub












Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If j = 3 Then
Call shell_tray_DblClk
j = 0
End If
If Button = 1 Then
j = j + 1
End If

End Sub

Private Sub shell_tray_DblClk()
systemtray.ShellDel Me
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
Text2 = "Wait...."
Puertosbaners
Definidoudp App.Path & "\Data\portudp.dat", Trim(Text1)
Text2 = "Finish."
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ras = False
Timer2.Enabled = False
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
Form7.Move Form1.left + X + 1000, Form1.top + Y + 800
Form7.Show

End If
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
des
End Sub

Private Sub UserControl11_click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
Call Form_Unload(0)
Case 1
Me.WindowState = vbMinimized
Case 2
Form4.Show vbModal
End Select
End Sub


Private Sub des()
UserControl31.desfocus Me
UserControl32.desfocus Me
UserControl33.desfocus Me
UserControl34.desfocus Me
UserControl35.desfocus Me
UserControl36.desfocus Me
End Sub

Private Sub UserControl31_click()
On Error Resume Next
TreeView1.Nodes.Clear
Text1.Enabled = False

Text2 = "Scanning....."
For i = 1 To cuenta
Unload Winsock1(i)
Next i
For i = 1 To cuenta1
Unload Winsock2(i)
Next i
For i = 1 To cuenta2
Unload Winsock3(i)
Next i
Erase arg()
salir = False
Set b = TreeView1.Nodes.Add(, , , Trim(Text1), 8)
Set k = TreeView1.Nodes.Add(b, tvwChild, , "TCP", 20)
If Form5.Check1(1).Value = vbChecked Then sendicm TreeView1, Trim(Text1)
If Form5.Check1(0).Value = vbChecked Then TreeView1.Nodes.Add k, tvwChild, , "DNS:" & DnsName(Trim(Text1)), 16
If Form5.Option1(0).Value = True Then Definido App.Path & "\Data\porttcp.dat", Trim(Text1)
If Form5.Option1(1).Value = True Then Timer2.Interval = Val(Form5.Text2) * 1000: Range Text1
If Form5.Option1(2).Value = True And Form5.Text1 <> "" Then only CLng(Trim(Form5.Text1))
If Form5.Check1(5).Value = vbChecked Then Set r = TreeView1.Nodes.Add(, tvwChild, , "UDP", 20)

Text2 = "Finish."

If Form5.Check1(2).Value = vbChecked And Form5.Option1(1).Value = False Then Text2 = "Wait....": Timer1.Enabled = True

Text1.Enabled = True
End Sub

Private Sub UserControl32_click()
salir = True
End Sub

Private Sub UserControl33_click()
Form5.Show vbModal
End Sub

Private Sub UserControl34_click()
Form2.Show vbModal
End Sub

Private Sub UserControl35_click()
Form6.Show vbModal
End Sub

Private Sub UserControl36_click()
Form8.Show vbModal
End Sub

Private Sub Winsock1_Connect(Index As Integer)
On Error Resume Next
Dim can As Integer
If Winsock1(Index).State = sckConnected Then
can = Winsock1(Index).RemotePort

If Form5.Option1(0).Value = True Then Iconos Winsock1(Index).RemotePort, arg(Index)
If Form5.Option1(1).Value = True Then Iconos Winsock1(Index).RemotePort, "OPEN;"
If Form5.Option1(2).Value = True Then Iconos Winsock1(Index).RemotePort, "OPEN;"
If Form5.Check1(3).Value = vbChecked And can = CInt(Trim(Form5.Text4)) Then sonido Form5.Text3

End If
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1(Index).Close
End Sub

Private Sub Winsock2_Connect(Index As Integer)
On Error Resume Next
If Winsock2(Index).State = sckConnected Then
If Winsock2(Index).RemotePort = 80 Then Winsock2(Index).SendData "Http://" & Winsock2(Index).RemoteHostIP & " HTTP/1.1" & vbCrLf & vbCrLf

End If
End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim g As String
On Error Resume Next
Winsock2(Index).GetData g
If left(g, 4) = "HTTP" Then
v = Split(g, vbCrLf)
For i = LBound(v) To UBound(v)
If left(v(i), 6) = "Server" Then TreeView1.Nodes.Add Index, tvwChild, , v(i), 19
Next i
Else
TreeView1.Nodes.Add Index, tvwChild, , g, 18
End If
End Sub

Private Sub Winsock2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock2(Index).Close
End Sub

Private Sub Puertosbaners()
On Error Resume Next
Dim can As Long
For i = 1 To TreeView1.Nodes.Count

If IsNumeric(Mid(TreeView1.Nodes.Item(i).Text, 1, InStr(TreeView1.Nodes.Item(i).Text, " "))) = True And InStr(Mid(TreeView1.Nodes.Item(i).Text, 1, InStr(TreeView1.Nodes.Item(i).Text, " ")), ".") = 0 Then
cuenta1 = TreeView1.Nodes.Count
load Winsock2(i)
can = CLng(Mid(TreeView1.Nodes.Item(i).Text, 1, InStr(TreeView1.Nodes.Item(i).Text, " ")))
Winsock2(i).Close
Winsock2(i).Connect Trim(Text1), Trim(can)

End If

Next i
End Sub

Private Sub Winsock3_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo ema
Dim g As String
Dim t As Variant
Winsock3(Index).GetData g
If g = "22" Then
Set t = TreeView1.Nodes.Add(r, tvwChild, , Winsock3(Index).RemotePort, 1)
t.Expanded = True
End If
ema:
End Sub

