VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About TextPad"
   ClientHeight    =   4380
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5295
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3023.155
   ScaleMode       =   0  'User
   ScaleWidth      =   4972.279
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picabout 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5235
      TabIndex        =   10
      Top             =   0
      Width           =   5295
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Written By : Jason - Simeone"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.vb-world.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Hyperlink URL : http://www.vb-world.net"
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Registry access source code from: "
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please  Email me if you have any suggestions."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Compiled on : Sunday, January 07, 2001"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lbllicense 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This software is licensed too:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4732.821
      Y1              =   2101.714
      Y2              =   2101.714
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "A Small application for writing and saving text documents.  "
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4732.821
      Y1              =   2112.067
      Y2              =   2112.067
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 4.120 Beta 20"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   4125
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Email address : Cyberarea@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   3990
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me

End Sub
Private Sub Form_Load()
Dim license As String
Dim company As String
license = GetSettingString(HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion", _
"RegisteredOwner", "")
company = GetSettingString(HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion", _
"RegisteredOrganization", "")
lbllicense.caption = license & Chr(13) & Chr(10) & company
On Error GoTo Picerr:

Picabout.Picture = Frmnoreg.Picture1.Picture
' why Bring another hefty 62 kb along
' with us Just load the picture from frmnoreg

Picerr: ' error control
If Err.Number <> 0 Then
Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAbout.Picabout = Nothing
Set frmAbout = Nothing
End Sub

Private Sub Label5_Click()
gotoweb
End Sub

