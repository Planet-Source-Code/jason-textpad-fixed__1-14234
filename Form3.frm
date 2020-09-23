VERSION 5.00
Begin VB.Form Frmtxtfilemanager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text File manager"
   ClientHeight    =   6165
   ClientLeft      =   3000
   ClientTop       =   1695
   ClientWidth     =   9945
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.PictureBox Pictoobig 
      Height          =   5775
      Left            =   5400
      ScaleHeight     =   5715
      ScaleWidth      =   4275
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label Lblwhy 
         Caption         =   "Error, The File you selected Is too large too be displayed here. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   4095
      End
   End
   Begin VB.TextBox Txtfile 
      Height          =   5775
      Left            =   5400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete Selected File "
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      ForeColor       =   &H00000000&
      Height          =   2040
      Left            =   120
      Pattern         =   "*.TXT;*.INI;*.lOG"
      TabIndex        =   2
      Top             =   3480
      Width           =   4935
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selected File :"
      Height          =   6135
      Left            =   5280
      TabIndex        =   8
      Top             =   0
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select a file :"
      Height          =   5655
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected File :"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Frmtxtfilemanager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Dir1_Change()
Dim strdir As String
strdir = Dir1.Path
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dim msg, style, title, response
On Error GoTo driveerr:
Dir1.Path = Drive1.Drive
driveerr:

If Err.Number <> 0 Then
 
 response = MsgBox("Error , Please insert a disk into the selected drive. ", vbRetryCancel + vbExclamation + vbDefaultButton2, "Selected Drive error")

Select Case response
Case vbRetry
Call Drive1_Change
Exit Sub

Case vbCancel
Exit Sub
End Select
End If

End Sub

Private Sub File1_Click()
Reset ' reset all open disks
Pictoobig.visible = False
 
selectedfile = Frmtxtfilemanager.File1.Path & "\" & Frmtxtfilemanager.File1.filename
On Error GoTo cannotfindfileerr:
If FileLen(selectedfile) > 65000 Then Pictoobig.visible = True: Exit Sub

Close #1

On Error GoTo cannotfindfileerr:

Open selectedfile For Binary Access Read As #1
Txtfile.Text = Input(LOF(1), 1)
Close #1

cannotfindfileerr:
If Err.Number <> 0 Then

selectedfile = Frmtxtfilemanager.File1.Path & Frmtxtfilemanager.File1.filename

Close #1
Open selectedfile For Binary Access Read As #1
Txtfile.Text = Input(LOF(1), 1)
Close #1
Exit Sub
End If

Err:
Exit Sub
End Sub



Private Sub Form_Unload(Cancel As Integer)
Close #1
Set Frmtxtfilemanager = Nothing
End Sub

Private Sub Txtfile_KeyPress(KeyAscii As Integer)
MsgBox "You cannot Edit Files In the File Manager.", vbInformation, "TextPad"

End Sub
