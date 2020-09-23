VERSION 5.00
Begin VB.Form frmeditclipboard 
   Caption         =   "Edit ClipBoard Text"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   Icon            =   "Frmeditclipboard.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4620
      TabIndex        =   4
      Top             =   2655
      Width           =   4620
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Updates Window From Clipboard"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   4575
         Begin VB.CommandButton Command3 
            Caption         =   "&Save Too Clipboard"
            Height          =   375
            Left            =   1320
            TabIndex        =   2
            ToolTipText     =   "Writes Currently edited text too clipboard"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Clear Clipboard"
            Height          =   375
            Left            =   3120
            TabIndex        =   3
            ToolTipText     =   "Clears Clipboard Text"
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmeditclipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo clipboarderror:
Text1.Text = Clipboard.GetText
clipboarderror:

If Err.Number <> 0 Then
MsgBox "An error Has been encountered while accessing the Clipboard" _
& Chr(13) & Err.Description, vbCritical, "Error , Clipboard"
Exit Sub
End If

End Sub


Private Sub Command2_Click()
Clipboard.SetText ("")
End Sub

Private Sub Command3_Click()
Clipboard.SetText Text1.Text()
End Sub

Private Sub Form_Load()
On Error GoTo clipboarderror:
frmeditclipboard.Text1.Text = Clipboard.GetText

clipboarderror:
If Err.Number <> 0 Then
MsgBox "An error Has been encountered while accessing the Clipboard" _
& Chr(13) & Err.Description, vbCritical, "Error , Clipboard"
Exit Sub
End If
End Sub

Private Sub Form_Resize()
'Text1.Height = frmeditclipboard.ScaleHeight + -495
' Text1.Width = frmeditclipboard.ScaleWidth
  On Error GoTo Errresize:
  Text1.Width = Me.ScaleWidth - (Text1.Left * 2)
  Text1.Height = Me.Height - 1100
  Frame1.Width = Me.Width - 150
Errresize:
Exit Sub
End Sub

