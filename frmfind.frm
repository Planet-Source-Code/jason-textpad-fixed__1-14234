VERSION 5.00
Begin VB.Form frmfind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmfind.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "&Match Case "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F&ind..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Txtfind 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "&Find what :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Compare Binary
Private Sub Command1_Click()
    '*****vb written find source from Sdi Note application
    '*****\\\\\\\\\\\\///////////////////////
If Check1.Value = vbChecked Then
Casesensitivesearch_frmfind
Exit Sub
End If
findit

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Txtfind_Change()
If txtfind.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
Sub Casesensitivesearch_frmfind()
    Dim Search, Where   ' Declare variables.
    ' Get search string from user.
    strfind = frmfind.txtfind.Text
    Search = frmfind.txtfind.Text
    Where = InStr(Form1.ActiveControl.Text, Search) ' Find string in text.
    If Where Then   ' If found,
     Form1.ActiveControl.SelStart = Where - 1  ' set selection start and
     Form1.ActiveControl.SelLength = Len(Search)   ' set selection length.
       Form1.SetFocus

    Else
        MsgBox "Cannot find " & Chr(34) & Search & Chr(34) _
        , vbInformation, "TextPad" ' Notify user.
    End If
End Sub


