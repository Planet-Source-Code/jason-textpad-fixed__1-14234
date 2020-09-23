VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Untitled - TextPad                                                         "
   ClientHeight    =   5865
   ClientLeft      =   1035
   ClientTop       =   2565
   ClientWidth     =   8760
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   767
      ButtonWidth     =   714
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "new"
            Object.ToolTipText     =   "new"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "open"
            Object.ToolTipText     =   "open"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "save"
            Object.ToolTipText     =   "save"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "copy"
            Object.ToolTipText     =   "copy"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cut"
            Object.ToolTipText     =   "cut"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "paste"
            Object.ToolTipText     =   "paste"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "find"
            Object.ToolTipText     =   "find..."
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "close"
            Object.ToolTipText     =   "close file"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "properties"
            Object.ToolTipText     =   "Properties  (File)"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "time&date"
            Object.ToolTipText     =   "Insert Time and Date"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "options"
            Object.ToolTipText     =   "Options"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "About"
            Object.ToolTipText     =   "About Text pad"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   6720
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.TextBox Text2 
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   6735
   End
   Begin MSComDlg.CommonDialog CfontDialog 
      Left            =   2160
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   120
      Top             =   4920
   End
   Begin VB.TextBox Text1 
      Height          =   4250
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label lblfilename 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   5535
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":27A2
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":28B4
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":29C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3096
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":31A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":32BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":33CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnufileitem 
      Caption         =   " &File  "
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnunewfileitem 
         Caption         =   "&New "
      End
      Begin VB.Menu mnuopenfileitem 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuclosefileitem 
         Caption         =   "&Close "
      End
      Begin VB.Menu mnusaveitem 
         Caption         =   "S&ave"
      End
      Begin VB.Menu mnusaveasfile 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu line12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuproperties 
         Caption         =   "&Properties"
         Shortcut        =   ^P
      End
      Begin VB.Menu line8 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnurecentfile1 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnurecentfile2 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnurecentfile3 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuedititem 
      Caption         =   " &Edit  "
      Begin VB.Menu mnucopyitem 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnucutitem 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnupasteitem 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnudeleteitem 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuselectallitem 
         Caption         =   "&Select All"
         Shortcut        =   ^S
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinserttimeitem 
         Caption         =   "Insert &Time"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuinsertdateitem 
         Caption         =   "Insert &Date"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuinserttimeanddateitem 
         Caption         =   "I&nsert Time\Date"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuclearclipboardtextitem 
         Caption         =   "Cl&ear Clipboard Text "
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuviewclipboardtextitem 
         Caption         =   "&View Clipboard Text"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnueditclipboard 
         Caption         =   "Edit Clipboard Te&xt "
         Shortcut        =   ^E
      End
      Begin VB.Menu line13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuwordwrap 
         Caption         =   "&Word Wrap"
      End
      Begin VB.Menu mnusetfont 
         Caption         =   "Set &Font"
      End
   End
   Begin VB.Menu mnusearchitem 
      Caption         =   " &Search  "
      Begin VB.Menu mnufinditem 
         Caption         =   "Fin&d..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnufindnextitem 
         Caption         =   "&Find Next"
      End
   End
   Begin VB.Menu mnuoptionsitem 
      Caption         =   " &View  "
      Begin VB.Menu mnulaunchnewinstanceitem 
         Caption         =   "&Launch new instance"
         Begin VB.Menu mnunormal 
            Caption         =   "&Normal "
         End
         Begin VB.Menu mnumaximized 
            Caption         =   "&Maximized "
         End
         Begin VB.Menu mnuminimized 
            Caption         =   "Minimi&zed"
         End
         Begin VB.Menu line2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuchoose 
            Caption         =   "&Choose..."
         End
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhidetoolbaritem 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnudeleteatextfileitem 
         Caption         =   "Text &File Manager "
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufullscreen 
         Caption         =   "Full &Screen"
         Shortcut        =   {F5}
      End
      Begin VB.Menu line11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuoptionsdialogitem 
         Caption         =   "&Options "
      End
   End
   Begin VB.Menu MNUHELPITEM 
      Caption         =   "  &Help    "
      Begin VB.Menu mnubugfixes 
         Caption         =   "Bug &Fixes"
      End
      Begin VB.Menu line9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnueditFileinfoitem 
      Caption         =   "&mnueditFileinfoitem"
      Visible         =   0   'False
      Begin VB.Menu mnucopyitem1 
         Caption         =   "&Copy"
      End
   End
   Begin VB.Menu Mnueditfileinfoitem2 
      Caption         =   "&Mnueditfileinfoitem2"
      Visible         =   0   'False
      Begin VB.Menu Mnucopyitem2 
         Caption         =   "&Copy"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
form1_startup ' call the startup procedure in
' modmain
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim msg, response

If fstate.dirty = True Then GoTo Unload:
If fstate.dirty = False Then End


Unload:

If Form1.lblfilename.caption <> "" Then
 msg = "The Text in " _
 & Form1.lblfilename.caption & " File has changed" _
& Chr(13) & Chr(13) & "Do you wish too save The changes ?"
Else ' * If there isnt then Display The One Below
msg = "The Text in the untitled file has changed" _
& Chr(13) & Chr(13) & "Do you wish too save the changes?"
End If

response = MsgBox(msg, vbYesNoCancel + vbExclamation + vbDefaultButton3, "TextPad ")

Select Case response
Case vbYes     ' User chose Yes.
Filequicksave ' Call Procedure in modmain
End
MyString = "Yes"    ' Perform some action.
 
Case vbCancel
Cancel = True
Exit Sub

Case vbNo
    End
    MyString = "No" ' Perform some action.
End Select



End Sub
Private Sub Form_Resize()
On Error GoTo Nxterr: ' error control
If Form1.WindowState = vbMinimized Then Exit Sub
Call Resizenotewithtoolbar

Nxterr:
If Err.Number <> 0 Then
Exit Sub
End If
End Sub
Private Sub mnuabout_Click()
Beep
Load frmAbout
frmAbout.Show (vbModal)
End Sub
Private Sub mnubugfixes_Click()
Load Frmbugz
Frmbugz.Show (vbModal)
End Sub

Private Sub mnuchoose_Click()
Load Frmnewinstance
Frmnewinstance.Show (vbModal)
End Sub

Private Sub Mnuclearclipboardtextitem_Click()
Clipboard.SetText ("")
End Sub

Private Sub mnuclosefileitem_Click()
Call mnunewfileitem_Click
On Error GoTo ErrFile:
CommonDialog1.filename = ("")
Form1.ActiveControl.Text = ""
Form1.caption = "Untitled - TextPad"
lblfilename.caption = ("")
fstate.dirty = False ' Tell Text pad That This is False now
Close #1
ErrFile:
If Err.Number <> 0 Then
Exit Sub
End If
End Sub
Private Sub mnucopyitem_Click()
Clipboard.SetText Form1.ActiveControl.SelText
' Simple Clipboard Copy Procedure
End Sub
Private Sub mnucutitem_Click()
Clipboard.SetText Form1.ActiveControl.SelText
Form1.ActiveControl.SelText = ""
End Sub
Private Sub mnudeleteatextfileitem_Click()
Load Frmtxtfilemanager
Frmtxtfilemanager.Show (vbModal)
End Sub
Private Sub mnudeleteitem_Click()
Form1.ActiveControl.SelText = ""
End Sub

Private Sub mnueditclipboard_Click()
Load frmeditclipboard
frmeditclipboard.Show (vbModal)

End Sub

Private Sub mnuexit_Click()
Call Form_QueryUnload(1, 0)
' instead of Repitiusly leaving this exit code here
' Too save Space and make textpad faster and smaller
' well just reuse Form_queryunloads code instead
' Because Cancel  = true Isnt working on module
' level :(

End Sub
Private Sub mnufileinformationitem_Click()
Load frmfileinfo
frmfileinfo.Show (vbModal)

End Sub
Private Sub mnufinditem_Click()
Load frmfind
frmfind.Show (vbModeless), Me
End Sub

Private Sub mnufindnextitem_Click()
findnexttext ' Just call the same exact code form Module1
' another waste of code Little too my knowledge last
' week i had written the same code in Module1 and
' i musty have forgotten OOPS !!!
End Sub
Private Sub mnufullscreen_Click()
If frmfullscreen.visible = False Then
mnufullscreen.Checked = True
Load frmfullscreen
frmfullscreen.Show
Form1.WindowState = vbMinimized
Else
If frmfullscreen.visible = True Then
mnufullscreen.Checked = False
Unload frmfullscreen
Unload Frmleavefullscreen
End If
End If

End Sub

Private Sub mnuhidetoolbaritem_Click()
  Dim retval As Integer
    Toolbar.visible = Not Toolbar.visible
    ' Change the check to match the current state
   mnuhidetoolbaritem.Checked = Toolbar.visible
    ' Call the resize procedure
    Resizenotewithtoolbar
    Select Case Toolbar.visible
    Case True
    retval = 1
    Case False
    retval = 0
    End Select
    SaveRegistryString "Toolbar", "Visible", retval
End Sub

Private Sub mnuinsertdateitem_Click()
On Error GoTo memoryerror:
Form1.ActiveControl.SelText = Date
memoryerror:
If Err.Number <> 0 Then
MsgBox "TextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
& Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then TextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "TextPad "
Exit Sub
End If

End Sub

Private Sub mnuinserttimeanddateitem_Click()
On Error GoTo memoryerror:
Form1.ActiveControl.SelText = Now ' Inserts Sytem Time and date
memoryerror:
If Err.Number <> 0 Then
MsgBox "TextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
& Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then TextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "TextPad "

Exit Sub
End If

End Sub

Private Sub mnuinserttimeitem_Click()
On Error GoTo outofmemory:
Form1.ActiveControl.SelText = Time
outofmemory:
If Err.Number <> 0 Then
MsgBox "TextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
& Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then TextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "TextPad "
Exit Sub
End If

End Sub


Private Sub mnumaximized_Click()
ShellNewTextPad (vbMaximizedFocus)
End Sub

Private Sub mnuminimized_Click()
ShellNewTextPad (vbMinimizedFocus)
End Sub

Private Sub mnunewfileitem_Click()
If fstate.dirty = True Then
 newfile 'call newfile procedure in module1
Else
If fstate.dirty = False Then 'else from first if then statement
Form1.ActiveControl.Text = ("")
lblfilename.caption = ""
fstate.dirty = False
Form1.caption = "Untitled - TextPad"
End If
End If
End Sub

Private Sub mnunormal_Click()
ShellNewTextPad (vbNormalFocus)
End Sub

Private Sub mnuopenfileitem_Click()
Openfile ' call openfile proc in module1
End Sub

Private Sub mnuoptionsdialogitem_Click()
On Error GoTo Objectunloadederror:
Load frmOptions
frmOptions.Show (vbModal)
Objectunloadederror:
If Err.Number <> 0 Then
MsgBox "TextPad Has encountered An Unexpected error" _
& Chr(13) & Err.Description, vbCritical, "Error"
Exit Sub
End If
End Sub

Private Sub mnupasteitem_Click()
On Error GoTo Outofmemoryerr:
Form1.ActiveControl.SelText = Clipboard.GetText()
Outofmemoryerr:
If Err.Number <> 0 Then
MsgBox "TextPad Has Encountered The Following Error(s) While Pasting Your Selection : " _
& Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then TextPad Cannot Paste Anymore Into The Text Box Because It Has Run Out of memory", vbCritical, "TextPad "
Exit Sub
End If
End Sub

Private Sub mnuproperties_Click()
Load frmfileinfo
frmfileinfo.Show (vbModal)

End Sub

Private Sub mnurecentfile1_Click()
recentfile1 ' call recent file 1 procedure in
'Module1
End Sub

Private Sub mnusaveasfile_Click()
'********************************
'This method of saving will Give the user
'a choice of how too save the file
'With a certain extension
'********************************
Close #1
CommonDialog1.Cancelerror = True
CommonDialog1.Filter = "Text documents (*.TXT) |*.TXT| INI files (*.INI) |*.INI| Log Files (*.LOG) |*.LOG| All Files (*.*) |*.* "
CommonDialog1.DialogTitle = "Save As"
CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt

On Error GoTo dialogerror
CommonDialog1.ShowSave
If CommonDialog1.filename <> "" Then
Open CommonDialog1.filename For Output As #1
Print #1, Form1.ActiveControl.Text
Close #1
fstate.dirty = False
lblfilename.caption = CommonDialog1.filename
Form1.caption = lblfilename.caption & " - TextPad"
dialogerror:
If Err.Number <> 0 Then
Exit Sub
Close #1
End If
End If

End Sub

Private Sub mnusaveitem_Click()
Filequicksave
' call The filequicksave Procedure in module1
' Because Reusing Code OVER AND OVER AND OVER
' Can Get REALLY annoying and can slowdown TextPads
' Start up
fstate.dirty = False
End Sub

Private Sub mnuselectallitem_Click()
Form1.ActiveControl.SelStart = 0
Form1.ActiveControl.SelLength = Len(Form1.ActiveControl.Text)
End Sub

Private Sub mnusetfont_Click()
On Error GoTo Cancelerror:
CfontDialog.Flags = cdlCFBoth
CfontDialog.ShowFont
Form1.ActiveControl.Font = CfontDialog.fontname
Form1.ActiveControl.FontBold = CfontDialog.FontBold
Form1.ActiveControl.fontsize = CfontDialog.fontsize
SaveRegistryString "Font", "font", Form1.ActiveControl.fontname
SaveRegistryString "Font", "Fontsize", Form1.ActiveControl.fontsize
Cancelerror:
Exit Sub
End Sub

Private Sub mnuviewclipboardtextitem_Click()
On Error GoTo Objectunloaded:
Load frmclpboard
frmclpboard.Show (vbModal)
Objectunloaded:
Exit Sub
End Sub

Private Sub mnuwordwrap_Click()
Dim retval As Boolean
Dim regval As Integer
retval = Usewordwrap
Usewordwrap = Not Usewordwrap
retval = Usewordwrap
Togglewordwrap (retval)
Select Case retval
Case True
regval = 1
Case False
regval = 0
End Select
SaveRegistryString "Wordwrap", "Wordwrap", regval
Debug.Print retval
End Sub

Private Sub Text1_Change()
TextChangecontrol
' Call the TextChangecontrol Proc in
' Module 1 too handle Repitious
' code that Just wastes Valuable Coding time and space
' on the users Hard-disk

End Sub

Private Sub Text2_Change()

TextChangecontrol
' Call the TextChangecontrol Proc in
' Module 1 too handle Repitious
' code that Just wastes Valuable Coding time and space
' on the users Hard-disk

End Sub

Private Sub Timer4_Timer()
On Error Resume Next
If Form1.ActiveControl.SelText <> "" Then
mnuselectedtextitem.Enabled = True
Else
mnuselectedtextitem.Enabled = False
End If

On Error Resume Next
If Clipboard.GetText <> "" Then
Mnuclearclipboardtextitem.Enabled = True
mnupasteitem.Enabled = True
mnuviewclipboardtextitem.Enabled = True
Else
Mnuclearclipboardtextitem.Enabled = False
mnupasteitem.Enabled = False
mnuviewclipboardtextitem.Enabled = False
End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
On Error GoTo Errorhandler:
Select Case Button.Key
Case "open"
Call mnuopenfileitem_Click
Case "save"
Call mnusaveitem_Click
Case "copy"
Call mnucopyitem_Click
Case "cut"
Call mnucutitem_Click
Case "paste"
Call mnupasteitem_Click
Case "find"
Call mnufinditem_Click
Case "close"
Call mnuclosefileitem_Click
Case "time&date"
Call mnuinserttimeanddateitem_Click
Case "options"
Call mnuoptionsdialogitem_Click
Case "About"
Call mnuabout_Click
Case "new"
Call mnunewfileitem_Click
Case "properties"
Call mnuproperties_Click
End Select
Errorhandler:
If Err.Number <> 0 Then
MsgBox "An Unexpected Error has occured while accessing the Toolbar", vbCritical, "TextPad"
Exit Sub
End If
End Sub

Private Sub Toolbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuoptionsitem, 4, , , mnuhidetoolbaritem
End If
End Sub

