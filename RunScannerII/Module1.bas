Attribute VB_Name = "Module1"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public i As Long, test As Boolean
Public formulario As New Contorno
Public v() As String, arg() As String
Public k As Variant, r As Variant
Public cuenta As Integer, cuenta1 As Integer, cuenta2 As Integer
Public salir As Boolean, ras As Boolean

Public Sub desenfocar(Index As Long)

For i = 0 To Index
Form1.UserControl11.Item(i).desfocus
Next i
End Sub


Public Sub rotar(objeto As Label, texto As String)
For i = 1 To Len(texto)
objeto.Caption = objeto.Caption & Mid(texto, i, 1)
Sleep 50
If i = Len(texto) Then objeto.Caption = ""
Next i
End Sub


Public Function Puertos(ruta As String, list As ListView) As Boolean
Dim g As String
Dim r As Variant
On Error Resume Next
Open ruta For Input As #1
Do
Line Input #1, g
v = Split(g, ";")
For i = LBound(v) To UBound(v)
 If left(v(i), 1) = "v" Then
 Set r = list.ListItems.Add(, , Mid(v(i), 2, InStr(v(i), "*") - 2), , 1)
 r.SubItems(1) = Mid(v(i), InStr(v(i), "*") + 1)
 End If
 If left(v(i), 1) = "r" Then
 Set r = list.ListItems.Add(, , Mid(v(i), 2, InStr(v(i), "*") - 2), , 2)
 r.SubItems(1) = Mid(v(i), InStr(v(i), "*") + 1)
 End If
 If left(v(i), 1) = "a" Then
 Set r = list.ListItems.Add(, , Mid(v(i), 2, InStr(v(i), "*") - 2), , 3)
 r.SubItems(1) = Mid(v(i), InStr(v(i), "*") + 1)
 End If
 r.Checked = True
 Next i
Loop Until EOF(1)
Close #1
Puertos = True
End Function

Public Sub Saveport(ruta As String, list As ListView)
On Error Resume Next
Open ruta For Output As #1
For i = 1 To list.ListItems.Count
If list.ListItems.Item(i).SmallIcon = 1 Then Print #1, "v" & list.ListItems.Item(i).Text & "*" & list.ListItems.Item(i).ListSubItems(1).Text & ";"
If list.ListItems.Item(i).SmallIcon = 2 Then Print #1, "r" & list.ListItems.Item(i).Text & "*" & list.ListItems.Item(i).ListSubItems(1).Text & ";"
If list.ListItems.Item(i).SmallIcon = 3 Then Print #1, "a" & list.ListItems.Item(i).Text & "*" & list.ListItems.Item(i).ListSubItems(1).Text & ";"
Next i
Close #1
End Sub

'----------------------------------------------------------------
Public Sub Definido(ruta As String, ip As String)
Dim g As String, Index As Integer
On Error Resume Next
Index = 0
Open ruta For Input As #1
Do
Line Input #1, g

Index = Index + 1
load Form1.Winsock1(Index)
Form1.Winsock1(Index).Protocol = sckTCPProtocol
Form1.Winsock1(Index).Close
Form1.Winsock1(Index).Connect Trim(ip), Val(Trim(Mid(g, 2, InStr(g, "*") - 2)))
ReDim Preserve arg(Index)
arg(Index) = CStr(Mid(g, InStr(1, g, "*") + 1))
Loop Until EOF(1)
Close #1
cuenta = Index

End Sub
Public Sub Definidoudp(ruta As String, ip As String)
Dim g As String, Index As Integer, res As Long
On Error Resume Next
res = 6000
Index = 0
Open ruta For Input As #1
Do
Line Input #1, g

Index = Index + 1
res = res + 1
load Form1.Winsock3(Index)
Form1.Winsock3(Index).Protocol = sckUDPProtocol
Form1.Winsock3(Index).Close
Form1.Winsock3(Index).LocalPort = res
Form1.Winsock3(Index).RemoteHost = ip
Form1.Winsock3(Index).RemotePort = Val(Trim(Mid(g, 2, InStr(g, "*") - 2)))
Form1.Winsock3(Index).Bind res, Form1.Winsock3(Index).LocalIP
Form1.Winsock3(Index).SendData "22"

Loop Until EOF(1)
Close #1
cuenta2 = Index
End Sub

Public Sub Iconos(sed As Long, argumento As String)
On Error Resume Next
Dim t As Variant, sid As Long

If sed = 21 Then
sid = 9
ElseIf sed = 25 Then
sid = 11
ElseIf sed = 23 Then
sid = 7
ElseIf sed = 53 Then
sid = 9
ElseIf sed = 79 Then
sid = 12
ElseIf sed = 80 Then
sid = 15
ElseIf sed = 81 Then
sid = 9
ElseIf sed = 88 Then
sid = 13
ElseIf sed = 110 Then
sid = 9
ElseIf sed = 135 Then
sid = 9
ElseIf sed = 139 Then
sid = 10
Else
sid = 1
End If

Set t = Form1.TreeView1.Nodes.Add(k, tvwChild, , sed & "    " & Mid(argumento, 1, Len(argumento) - 1), sid)

t.Expanded = True
t.EnsureVisible
End Sub


Public Sub Range(ip As String)
On Error Resume Next
Dim port As Long
Dim m As Integer
port = 0
m = 0
Do
port = port + 1
m = m + 1
load Form1.Winsock1(m)
Form1.Winsock1(m).Close
Form1.Winsock1(m).Connect Trim(ip), port
Form1.Text2 = "Scanned Port:" & port
If m >= 255 Then
Form1.Timer2.Enabled = True
ras = True
If salir = True Then Exit Do
Do
DoEvents
Loop Until ras = False
Form1.Timer2.Enabled = False
For i = 1 To m
Unload Form1.Winsock1(i)
Next i
m = 0
End If
Loop Until port >= 65526
salir = False
End Sub
Public Sub only(port As Long)
On Error Resume Next
Form1.Winsock1(0).Close
Form1.Winsock1(0).Connect Trim(Form1.Text1), port
End Sub


Sub sonido(src As String)
sndPlaySound src, 0&
End Sub
