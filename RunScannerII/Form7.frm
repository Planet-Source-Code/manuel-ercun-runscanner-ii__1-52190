VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   LinkTopic       =   "Form7"
   ScaleHeight     =   2175
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":27B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483644
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   9596
      EndProperty
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error Resume Next
ListView1.Move 0, 0
ListView1.View = lvwReport
Set ListView1.SmallIcons = ImageList1
ListView1.ListItems.Add 1, , "Copiar IP", , 2
ListView1.ListItems.Add 2, , "Delete all", , 1
End Sub

Private Sub Form_Resize()
Me.Width = ListView1.Width
Me.Height = ListView1.Height
End Sub

Private Sub ListView1_Click()
On Error Resume Next
If ListView1.SelectedItem.SmallIcon = 1 Then Form1.TreeView1.Nodes.Clear
If ListView1.SelectedItem.SmallIcon = 2 Then Clipboard.SetText Form1.Text1
Unload Me
End Sub

Private Sub ListView1_LostFocus()
On Error Resume Next
Unload Me
End Sub
