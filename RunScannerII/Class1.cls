VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203

Event DblClk()

Private WithEvents ipictures As PictureBox
Attribute ipictures.VB_VarHelpID = -1
Dim sn As NOTIFYICONDATA
Dim iToolTipText As String


Public Property Get Pictures() As PictureBox
Set Pictures = ipictures
End Property
Public Property Set Pictures(ByVal new_pictures As PictureBox)
Set ipictures = new_pictures

sn.hwnd = new_pictures.hwnd
sn.hIcon = new_pictures.picture
End Property

Private Sub ipictures_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X = WM_LBUTTONDBLCLK Then RaiseEvent DblClk
End Sub



Public Property Get ToolTipText() As String
ToolTipText = iToolTipText
End Property
Public Property Let ToolTipText(ByVal new_tool As String)
iToolTipText = new_tool
sn.szTip = new_tool
End Property



Private Sub Class_Initialize()

sn.cbSize = Len(sn)
sn.hIcon = 0&
sn.hwnd = 0&
sn.szTip = vbNullString
sn.uID = 2&
sn.uCallbackMessage = WM_MOUSEMOVE
sn.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
End Sub
Public Function ShellAdd(Form As Form)
Form.Hide
Shell_NotifyIcon NIM_ADD, sn
End Function

Public Function ShellDel(Form As Form)
Shell_NotifyIcon NIM_DELETE, sn
Form.WindowState = vbNormal
Form.Show
End Function





