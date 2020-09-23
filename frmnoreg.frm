VERSION 5.00
Begin VB.Form Frmnoreg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TextPad Setup"
   ClientHeight    =   5025
   ClientLeft      =   2550
   ClientTop       =   3690
   ClientWidth     =   5820
   Icon            =   "frmnoreg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsetdefaultoptions 
      Caption         =   "Set Default Options (Recommended)"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CheckBox ChckExternalEditor 
      Caption         =   "Use External Editor Too open Files Too large For TextPad Too open."
      Height          =   255
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   $"frmnoreg.frx":000C
      Top             =   3600
      Width           =   5415
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      Picture         =   "frmnoreg.frx":00B6
      ScaleHeight     =   675
      ScaleWidth      =   5760
      TabIndex        =   7
      Top             =   0
      Width           =   5820
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      ToolTipText     =   "Save settings ,Close window."
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox Chckassociations 
      Caption         =   "Textpad &should check whether it is the default text viewer"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "This enables textpad too notify you when textpad is not the default text viewer."
      Top             =   3240
      Width           =   4455
   End
   Begin VB.CheckBox Chckwordwrap 
      Caption         =   "Use &Word - Wrap"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "This Enables textpad too wrap text too the text box."
      Top             =   2880
      Width           =   4455
   End
   Begin VB.CheckBox Chcktoolbar 
      Caption         =   "Always show &Toolbar"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "This enables TextPad too Show the toolbar every time you start textpad."
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   5655
      Begin VB.CheckBox chckallowassociation 
         Caption         =   "Allow TextPad too be the default Text Viewer (*.TXT) on this Computer"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   5415
      End
   End
   Begin VB.Label Label2 
      Caption         =   "When Youre Done Choosing Options Press  The ""Done"" Button.  ."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmnoreg.frx":77F0
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "Frmnoreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

If Chcktoolbar.Value = vbChecked Then
SaveRegistryString "Toolbar", "Visible", 1
' Toolbar Will be Visible At startup
Else
If Chcktoolbar.Value = vbUnchecked Then
SaveRegistryString "Toolbar", "Visible", 0
' Toolbar  Will Not Be visible At startup
End If
End If

If Chckwordwrap.Value = vbChecked Then
SaveRegistryString "Wordwrap", "Wordwrap", 1
' this will svae the settings too the registry
Else
If Chckwordwrap.Value = vbUnchecked Then
SaveRegistryString "Wordwrap", "Wordwrap", 0
End If
End If

If Chckassociations.Value = vbChecked Then
SaveRegistryString "chckassociations", "show", 1
Else
If Chckassociations.Value = vbUnchecked Then
SaveRegistryString "chckassociations", "show", 0
End If
End If

If ChckExternalEditor.Value = vbChecked Then
DetectExternalEditor
      SaveRegistryString "UseExternalEditor", "use", 1
         SaveRegistryString "UseExternalEditor", "path", ExternalEditorPath

    Else
If ChckExternalEditor.Value = vbUnchecked Then
      SaveRegistryString "UseExternalEditor", "use", 0
End If
End If

If Chckassociations = vbChecked Then
savestring HKEY_CLASSES_ROOT, _
"Txtfile\shell\open\command", _
"", App.Path & "\" & App.EXEName & ".EXE" & " %1"
End If

Dim msg, style, response, title

msg = "Settings Have been successfully saved" & _
vbCrLf & vbCrLf & "Would you like too Start TextPad now ?"
 
Beep
response = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "TextPad")
Beep

Select Case response 'Begin select case clause
Case vbYes ' if user selects The yes button then
ShellNewTextPad (vbNormalFocus)
Unload Me ' unload this form frmnoreg
End ' Stop code so the hidden form1 can be terminated
' and the new one can be shelled with the new settings
Case vbNo
End
End Select
End Sub

Private Sub cmdsetdefaultoptions_Click()
Chcktoolbar.Value = vbChecked
ChckExternalEditor.Value = vbChecked
Chckwordwrap.Value = vbChecked
chckallowassociation.Value = vbChecked
End Sub

Private Sub Form_Load()
Beep ' beep too grab the users attention

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
' if the user closes this window we dont
'want textpad too continually stay open
'in the backround so well stop
'the code and exit immediately.
End Sub
