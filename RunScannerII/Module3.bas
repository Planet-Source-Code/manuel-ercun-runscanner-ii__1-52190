Attribute VB_Name = "Module3"
Option Explicit

'ping

Private pinge As IP_ECHO_REPLY
Private pingo As IP_OPTION_INFORMATION

Public Function ping(ip As String, time As Long, pack As Integer) As String
Dim retorno As String, icm As Long, res As Long
On Error Resume Next
time = Val(time) * 1000


icm = IcmpCreateFile
For i = 1 To pack
DoEvents
res = IcmpSendEcho(icm, Remote(ip), 0, 0, pingo, pinge, Len(pinge), time)
retorno = retorno & "i" & ip & ";" & "t" & CStr(Trim(pinge.RoundTripTime)) & ";" & "l" & CStr(Trim(pinge.Options.TTL)) & ";"
Next i
IcmpCloseHandle icm
If retorno <> "" Then ping = retorno
End Function

