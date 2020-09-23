VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options - Toolbar "
   ClientHeight    =   3915
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5970
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picoptions 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   2
      Left            =   240
      ScaleHeight     =   2415
      ScaleWidth      =   5535
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame2 
         Caption         =   "Misc."
         Height          =   1815
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   5415
         Begin VB.Frame Frame3 
            Caption         =   "Current External Editor"
            Height          =   735
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   5175
            Begin VB.CommandButton cmdChooseexternaleditor 
               Caption         =   "Choose External Editor ......"
               Height          =   375
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   4935
            End
         End
         Begin VB.CheckBox ChckExternalEditor 
            Caption         =   "Use External Editor Too open Files Too large For TextPad Too open."
            Height          =   375
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   $"frmOptions.frx":000C
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "frmOptions.frx":00B6
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   0
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox Picoptions 
      BorderStyle     =   0  'None
      Height          =   2460
      Index           =   1
      Left            =   240
      ScaleHeight     =   2460
      ScaleMode       =   0  'User
      ScaleWidth      =   5535
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame1 
         Caption         =   "File associations"
         Height          =   1770
         Left            =   0
         TabIndex        =   16
         Top             =   600
         Width           =   5415
         Begin VB.CheckBox Chckassociations 
            Caption         =   "&Textpad should check wether it is the default text viewer"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Enable this option if you would like for text pad too check wether it is the default text viewer."
            Top             =   720
            Width           =   4335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Allow textpad too be associated with Text files (*.TXT)"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Enable this option too associate text pad With Text files , disable for It not too be associated with Text files. "
            Top             =   360
            Width           =   4215
         End
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   0
         Picture         =   "frmOptions.frx":04F8
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   1560
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.PictureBox Picoptionsd 
      BorderStyle     =   0  'None
      Height          =   3780
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox nonusable 
      BorderStyle     =   0  'None
      Height          =   3780
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   13
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox ippy 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   19
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   12
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picoptions 
      BorderStyle     =   0  'None
      Height          =   2460
      Index           =   0
      Left            =   240
      ScaleHeight     =   2460
      ScaleWidth      =   5535
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   5535
      Begin VB.Frame fraSample1 
         Caption         =   "Toolbar"
         Height          =   2010
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   5415
         Begin VB.CheckBox Check2 
            Caption         =   "&Always Show toolbar  (default)"
            Height          =   255
            Left            =   120
            TabIndex        =   1
            ToolTipText     =   "Enable this option too always show the toolbar , Disable Too hide the toolbar"
            Top             =   480
            Width           =   3735
         End
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   0
         Picture         =   "frmOptions.frx":0FB2
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "A&pply"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      ToolTipText     =   "Saves any changes you have made without closing this dialog box."
      Top             =   3495
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      ToolTipText     =   "Cancels any changes you have made and Closes this dialog box."
      Top             =   3495
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2370
      TabIndex        =   4
      ToolTipText     =   "saves changes and closes This dialog box. "
      Top             =   3495
      Width           =   1095
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   3285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5794
      TabWidthStyle   =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Toolbar "
            Key             =   "Group1"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click for toolbar options"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "File associations"
            Key             =   "Group2"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click for File associations"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "External Editor    "
            Key             =   "Group3"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click For External Editor Options"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOptions.frx":12BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOptions.frx":15D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOptions.frx":18F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOptions.frx":1C0A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
  
 Call SaveMainSettings
  RetrieveALLSettings

End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Resizenotewithtoolbar
End Sub

Private Sub cmdChooseexternaleditor_Click()
On Error GoTo CdlcCancelErr:
CDialog.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
CDialog.Filter = "Executable Files(*.EXE) |*.EXE"
CDialog.DialogTitle = "Choose An External Viewer IE ;  Wordpad.EXE "
CDialog.Cancelerror = True
CDialog.ShowOpen
If CDialog.filename <> "" Then
SaveRegistryString "Useexternaleditor", "Path", CDialog.filename
cmdChooseexternaleditor.caption = CDialog.filename
End If
CdlcCancelErr:
If Err.Number = 32755 Then
Exit Sub
End If
End Sub

Private Sub cmdOK_Click()

 
 Call SaveMainSettings
  Unload Me
RetrieveALLSettings

End Sub

Private Sub Cmdclose_Click()
Unload Me
End Sub
Private Sub SaveMainSettings()
Dim retval As String ' Declare String variable
retval = GetSettingString(HKEY_CLASSES_ROOT, _
"Txtfile\shell\open\command", _
"", App.Path & "\" & App.EXEName & ".EXE" & " %1")

If Check1.Value = vbChecked Then
savestring HKEY_CLASSES_ROOT, _
"Txtfile\shell\open\command", _
"", App.Path & "\" & App.EXEName & ".EXE" & " %1"
End If

If Check1.Value = vbUnchecked Then
On Error GoTo cdialogerr:
If retval = App.Path & "\" & App.EXEName & ".EXE" & " %1" Then ' 2

CDialog.Flags = cdlOFNHideReadOnly
CDialog.Filter = "Executables (*.EXE) |*.EXE"
CDialog.Cancelerror = True
CDialog.DialogTitle = "Please select notepad (Usually Located In your windows directory) and press open"
CDialog.ShowOpen
MsgBox CDialog.filename & _
 vbCrLf & " will now be used too view text files on this computer", vbInformation, "Text pad"
If CDialog.filename <> "" Then ' 3
savestring HKEY_CLASSES_ROOT, _
"Txtfile\shell\open\command", _
"", CDialog.filename & " %1"
End If
' 2 TODO : Else Statement
End If
' 3 TODO : ELSE STATEMENT
End If


SaveSetting_Toolbar Check2

SaveSetting_chckassociations Chckassociations

saveSetting_UseExternalEditor ChckExternalEditor

cdialogerr:
Exit Sub

End Sub


Private Sub Form_Load()
Dim retval As String
Dim strregreading As String
strregreading = GetSettingString(HKEY_CLASSES_ROOT, _
"txtfile\Shell\open\command", _
"", "")
If strregreading = App.Path & "\" & App.EXEName & ".EXE" & " %1" Then
Check1.Value = vbChecked
Else
Check1.Value = vbUnchecked
End If
If GetSetting("Textpad", "chckassociations", "Show", 1) = 1 Then
Chckassociations.Value = vbChecked
Else
Chckassociations.Value = vbUnchecked
End If
' above for text pad associations

'Below For toolbar reg options

If GetSetting("Textpad", "Toolbar", "Visible", 1) = 1 Then
Check2.Value = vbChecked
Else
If GetSetting("Textpad", "Toolbar", "Visible", 0) = 0 Then
Check2.Value = vbUnchecked
End If
End If

' Below For External Editor Options

If GetSetting("Textpad", "UseExternaleditor", "use", 1) = 1 Then
ChckExternalEditor.Value = vbChecked
Else
If GetSetting("TextPad", "UseExternalEditor", "use", 0) = 0 Then
ChckExternalEditor.Value = vbUnchecked
End If
End If

retval = GetSetting("TextPad", "UseExternalEditor", "Path")
If retval = vbNullString Then
DetectExternalEditor
cmdChooseexternaleditor.caption = ExternalEditorPath
Else
If retval <> "" Then
cmdChooseexternaleditor.caption = retval
End If
End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        ' visual basic has created the bottom code
    'for the tab's key constants
    '\\\\\\\\\\\\\\\\\\\///////////////////////////
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmOptions = Nothing
End Sub

Private Sub tbsOptions_Click()
    Select Case tbsOptions.SelectedItem.Index
    Case 1
        frmOptions.caption = "Options - Toolbar "
    Case 2
        frmOptions.caption = "Options - File Associations"
    Case 3
        frmOptions.caption = "Options - External Editor."
    End Select
    ' ABOVE ^^^^^^^ Use Select Case Statement Instead
    '       |||||||
    ' Of Barberic IF THEN Statement Too Set the forms Caption
    ' Depending on the options selected Through The Tbsoptions
    ' .selecteditem.index Property
    
    Dim i As Integer
    ' visual basic has created the bottom code
    'for the tab's
    '\\\\\\\\\\\\\\\\\\\///////////////////////////
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            Picoptions(i).Left = 210
            Picoptions(i).visible = True
            Picoptions(i).Enabled = True
        Else
            Picoptions(i).Left = -20000
            Picoptions(i).Enabled = False
        End If
    Next
    
End Sub
