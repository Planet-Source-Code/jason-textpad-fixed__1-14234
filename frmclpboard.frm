VERSION 5.00
Begin VB.Form frmclpboard 
   Caption         =   "Clipboard text"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "frmclpboard.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Clipboard Text: "
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmclpboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
If Clipboard.GetText <> "" Then
Text1.Text = Clipboard.GetText
Else
frmclpboard.Hide
MsgBox "There Is no Clipboard text Too display", vbExclamation, "Clipboard Text Error"
Unload Me
End If
End Sub

Private Sub Form_Resize()
On Error GoTo Resizeerror:
Text1.Height = Me.ScaleHeight - 241
Text1.Width = Me.ScaleWidth
Resizeerror:
Exit Sub
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
Beep
End Sub
