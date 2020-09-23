VERSION 5.00
Begin VB.UserControl UserControl2 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   ScaleHeight     =   3465
   ScaleWidth      =   5310
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Rect
     left As Long
     top As Long
     right As Long
     Button As Long
End Type


Dim angulo As Rect
Dim imax, imin, ivalue As Long

Public Property Get Max() As Long
Max = imax
End Property
Public Property Let Max(ByVal new_max As Long)
imax = new_max
If imax < imin Then imax = imin
If imax < ivalue Then imax = ivalue
PropertyChanged "Max"
End Property

Public Property Get min() As Long
min = imin
End Property
Public Property Let min(ByVal new_min As Long)
imin = new_min
If imin > imax Then imin = imax
If imin < ivalue Then imin = ivalue
PropertyChanged "Min"
End Property

Public Property Get Value() As Long
Value = ivalue
End Property
Public Property Let Value(ByVal new_value As Long)
ivalue = new_value
If ivalue < imin Then ivalue = imin
If ivalue > imax Then ivalue = imax
Call process(ivalue)
PropertyChanged "Value"
End Property
Private Sub process(ByRef porcentaje As Long)
Dim res As Long
res = Screen.TwipsPerPixelX
angulo.left = res
angulo.top = res
angulo.right = UserControl.ScaleWidth - res
angulo.Button = UserControl.ScaleHeight - res
UserControl.DrawMode = 13

UserControl.Line (angulo.left, angulo.top)-(angulo.right, angulo.Button), UserControl.BackColor, BF

If porcentaje > 0 Then
UserControl.DrawMode = 7
UserControl.Line (angulo.left, angulo.top)-((angulo.right / imax) * porcentaje, angulo.Button), UserControl.FillColor, BF
UserControl.Line (angulo.left, angulo.top)-((angulo.right / imax) * porcentaje, angulo.Button), UserControl.BackColor, BF
End If
End Sub


Private Sub usercontrolControl_ReadProperties(PropBag As PropertyBag)
Max = PropBag.ReadProperty("Max", 100)
min = PropBag.ReadProperty("Min", 0)
Value = PropBag.ReadProperty("Value", 0)
End Sub

Private Sub usercontrolControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Max", imax, 100)
Call PropBag.WriteProperty("Min", imin, 0)
Call PropBag.WriteProperty("Value", ivalue, 0)
End Sub

Private Sub UserControl_Initialize()
imax = 100
imin = 0
ivalue = 0
End Sub
