VERSION 5.00
Begin VB.Form Frmbugz 
   Caption         =   "Updates And improvements For TextPad V 4.120 Beta 20"
   ClientHeight    =   4335
   ClientLeft      =   2235
   ClientTop       =   2985
   ClientWidth     =   7455
   Icon            =   "Frmbugz.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   4335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Frmbugz.frx":27A2
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Frmbugz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Resize()
On Error GoTo Resizeerr:
Text1.Height = Frmbugz.ScaleHeight
Text1.Width = Frmbugz.ScaleWidth
Resizeerr:
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Frmbugz = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Beep
End Sub
