VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const RGN_DIFF = 4


Private i As Long, z As Long

Public Sub Rng(Form As Form, Color As Long)

Dim res As Long, com As Long, pixel As Long


Form.BackColor = Color

com = CreateRectRgn(0, 0, Form.Width / 15, Form.Height / 15)

For z = 0 To Form.Height / 15 'twip

For i = 0 To Form.Width / 15 'twip
    
pixel = GetPixel(Form.hDC, i, z)
        
If pixel = Color Then
res = CreateRectRgn(i, z, i + 1, z + 1)
If res <> 0 Then
Call CombineRgn(com, com, res, RGN_DIFF)
Call DeleteObject(res)
End If
End If
'process i,z,(Form.Height / 15) + (Form.Width / 15)
Next i
If test = True Then process z, (Form.Height / 15) + (Form.Width / 15)

Next z
       
If test = True Then process (Form.Height / 15) + (Form.Width / 15), (Form.Height / 15) + (Form.Width / 15)

Call SetWindowRgn(Form.hwnd, com, 1)
Call DeleteObject(com)
test = False
End Sub

Private Sub process(x As Long, Max As Long)
Form3.UserControl21.Max = Max
Form3.UserControl21.min = 0
'form3.usercontrol21.value = y+x
Form3.UserControl21.Value = x
End Sub

Public Sub mover(Form As Form)
ReleaseCapture
SendMessage Form.hwnd, &HA1, 2, 0&
End Sub

Public Sub Centro(Form As Form, position As String, Optional x As Long, Optional y As Long)
Select Case position
Case "center"
Form.Move (Screen.Width / 2) - (Form.Width / 2), (Screen.Height / 2) - (Form.Height / 2)
Case "optional"
Form.Move x, y
Case "secundario"
Form.Move (Form1.Width / 2) - (Form.Width / 2), (Form1.Height / 2) - (Form.Height / 2)
End Select
End Sub
