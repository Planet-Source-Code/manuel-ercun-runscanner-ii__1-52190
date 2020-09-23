Attribute VB_Name = "Module2"
Option Explicit

Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type


Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type



Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSAdata) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Byte, ByVal Length As Long, ByVal haddrtype As Long) As Long
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal name As String) As Long


Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Public Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean


Dim wsa As WSAdata
Public b As Variant
Dim res1 As Long
Private ipe As IP_ECHO_REPLY
Private ipo As IP_OPTION_INFORMATION
Public Sub Startup()
WSAStartup &H101, wsa
End Sub
Public Function Remote(ip As String) As Long
On Error Resume Next
Startup
Dim res As Long, res2 As Long, s As String
Dim host As Hostent
CopyMemory host, ByVal gethostbyname(ip), Len(host)
CopyMemory res, ByVal host.h_addr_list, 4
CopyMemory res1, ByVal res, 4
Remote = res1
WSACleanup
End Function


Public Function DnsName(ip As String) As String
On Error Resume Next
Dim lenip(3) As Byte
Dim lehost As Long, res As Long
Dim s As String, host As Hostent
Startup
res = inet_addr(ip)
CopyMemory lenip(0), res, Len(res)
lehost = gethostbyaddr(lenip(0), 4, 2)
CopyMemory host, ByVal lehost, Len(host)

 s = String(255, vbNullChar)
 
lstrcpy s, host.h_name
DnsName = left(s, InStr(s, vbNullChar) - 1)
WSACleanup
End Function

Public Sub sendicm(tree As TreeView, ip As String)
Dim res  As Long, se As Long, icm As Long
On Error Resume Next
ipo.TTL = 255

DoEvents
icm = IcmpCreateFile
res = IcmpSendEcho(icm, Remote(ip), 0, 0, ipo, ipe, Len(ipe), 3000)
se = CLng(ipe.Options.TTL)

If se >= 90 And se <= 142 Then
Set b = tree.Nodes.Add(k, tvwChild, , "O.S:(Windows)", 3)
ElseIf se >= 225 Or (se >= 50 And se <= 70) Then
Set b = tree.Nodes.Add(k, tvwChild, , "O.S:(Problably Unix)", 5)
ElseIf se = 0 Then
Set b = tree.Nodes.Add(k, tvwChild, , "O.S:(Undetermined)", 17)
Else
Set b = tree.Nodes.Add(k, tvwChild, , "O.S:(Other)", 14)
End If
IcmpCloseHandle icm


End Sub
