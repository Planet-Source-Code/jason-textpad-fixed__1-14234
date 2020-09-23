Attribute VB_Name = "Modmain"
'*****************************************
'**New version ; Everything Was cleaned , Everything Was Fixed
'**There Are NO Known Bugz That I know Of that TextPad Has _
'**Shown after the HUGE code overHaul , It may Look the same
'**But theres a noticable difference.

'\\GOOD things
'**Some Command line Bugz Fixed (Again)
'**Some Things Work Better than Expected

'\\BAD things
'**Code may be a little shaky for now until i REHAUL EVERYTHING
'**Code still looks like a Disaster
'**I havent Reached this Programs Full potental YET.


Public strfind As String
Option Explicit
Type fileopened
dirty  As Integer
End Type
Option Compare Text
Public Const Normal_Cdlogflags = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNLongNames
'**Public Currentfilename As String
Public ExternalEditorPath As String
Type filestring ' The heart of textpads BOOLEAN Memory
dirty As Integer ' Without This Then TextPad wouldnt
End Type ' Know if a file was changed or not
Public fstate As filestring

Public Const URL = "http://www.vb-world.net"
Public Const email = "Cyberarea@hotmail.com"
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1 'Unicode nul terminated string
Public Const REG_BINARY = 3 'Free form binary
Public Const REG_DWORD = 4 '32-bit number
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal Hkey As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" _
Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey _
As String, phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" _
Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey _
As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" _
Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal _
lpValueName As String) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" _
Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey _
As String, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, lpData _
As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" _
Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal _
lpValueName As String, ByVal Reserved As Long, ByVal _
dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub savestring(Hkey As Long, strPath As String, _
strValue As String, strdata As String)
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(Hkey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, _
REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand)
End Sub


Public Sub gotoweb()
Dim success As Long

success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub

Public Sub sendemail()
Dim success As Long
success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub


Sub findnexttext()
    If strfind <> "" Then
    Dim Search, Where   ' Declare variables.
    ' Get search string from user.
    Search = strfind
    Where = InStr(Form1.ActiveControl.Text, Search)   ' Find string in text.
    If Where Then   ' If found,
        Form1.ActiveControl.SelStart = Where - 1  ' set selection start and
       Form1.ActiveControl.SelLength = Len(Search)   ' set selection length.
    Else
        MsgBox "Cannot find  " & Chr(34) & Search & Chr(34) _
        , vbInformation, "TextPad" ' Notify user.
    End If
Else
If strfind = "" Then
Load Frmfindnext
Frmfindnext.Show (0), Form1

End If
End If

End Sub


  Sub findit()
       
    strfind = frmfind.txtfind.Text
    Dim Search, Where     ' Declare variables.
    ' Get search string from user.
    Search = frmfind.txtfind.Text
    Where = InStr(Form1.ActiveControl, Search) ' Find string in text.
    
    If Where Then   ' If found,
        Form1.ActiveControl.SelStart = Where - 1  ' set selection start and
      Form1.ActiveControl.SelLength = Len(Search)   ' set selection length.
    Form1.SetFocus
    
    strfind = frmfind.txtfind.Text
    Else
        MsgBox "Cannot find " & Chr(34) & Search & Chr(34) _
        , vbInformation, "Text Pad" ' Notify user.
    End If
  End Sub
Sub OpenFilecommandline()
Dim Text_control1, Text_Control2
Set Text_control1 = Form1.Text1
Set Text_Control2 = Form1.Text2

Open Command$ For Binary Access Read As #1
        If FileLen(Command$) >= 65000 Then GoTo toobig:
On Error GoTo toobig:
Text_control1.Text = Input(LOF(1), 1)
    
   Text_Control2.Text = Text_control1.Text
    Text_control1.Text = Text_Control2.Text
    Form1.caption = Command$ & "- TextPad"
    fstate.dirty = False
    Form1.lblfilename.caption = Command$
Close #1
toobig:
If Err.Number <> 0 Then

'Form1.lblfilename.caption = ""
Close #1
Reset
Close #1
Exit Sub
End If
End Sub
Sub Resizenotewithtoolbar()
On Error GoTo Errresize:
If Form1.WindowState = vbMinimized Then Exit Sub
    If Form1.Toolbar.visible And Form1.Text1.visible = True Then
        Form1.Text1.Height = Form1.ScaleHeight - Form1.Toolbar.Height
        Form1.Text1.Width = Form1.ScaleWidth
        Form1.Text1.Top = Form1.Toolbar.Height
    Else
        If Form1.Toolbar.visible And Form1.Text2.visible = True Then
        Form1.Text2.Height = Form1.ScaleHeight - Form1.Toolbar.Height
        Form1.Text2.Width = Form1.ScaleWidth
        Form1.Text2.Top = Form1.Toolbar.Height

    Else
        If Form1.Text1.visible = True Then
        Form1.Text1.Height = Form1.ScaleHeight
        Form1.Text1.Width = Form1.ScaleWidth
        Form1.Text1.Top = 0
        Else
        If Form1.Text2.visible = True Then
        Form1.Text2.Height = Form1.ScaleHeight
        Form1.Text2.Width = Form1.ScaleWidth
        Form1.Text2.Top = 0
       
       End If
      End If
    End If

Errresize:
If Err.Number <> 0 Then
Exit Sub
End If
End If
End Sub
Sub check_if_textpad_is_associated()
Dim chkreg As String

Dim msg, response

chkreg = GetSettingString(HKEY_CLASSES_ROOT, _
"Txtfile\shell\open\command", _
"", "")

If chkreg = App.Path & "\" & App.EXEName & " %1" Then
Exit Sub

Else

msg = "Textpad is not currently associated with Text Files Would you like it too be ?" _
& Chr(13) & Chr(10) & Chr(13) & Chr(10) & "For this Message Not too show the next time you Run text pad," _
& Chr(13) & Chr(10) & "Check off in the options window :" _
& Chr(13) & Chr(10) & "(Textpad should check wether it is the default text viewer) "

response = MsgBox(msg, vbYesNo + vbInformation + vbDefaultButton1, "Text Pad")

Select Case response
Case vbYes
savestring HKEY_CLASSES_ROOT, _
"Txtfile\shell\open\command", _
"", App.Path & "\" & App.EXEName & ".EXE" & " %1"

Case vbNo
Exit Sub
Resume Next
End Select
End If

End Sub
Sub Readonlyerror()
Close #1
If Err.Number <> 0 Then
MsgBox "Error, The file you are trying too save too Exists with read only attributes please select a different filename.", vbExclamation, "Error,TextPad"
Form1.CommonDialog1.Filter = "Text documents (*.TXT) |*.TXT| INI Configuration Files (*.INI) |*.INI| All Files (*.*) |*.* "
Form1.CommonDialog1.Flags = cdlOFNHideReadOnly
On Error GoTo dialogerror:
Form1.CommonDialog1.ShowSave
If Form1.CommonDialog1.filename <> "" Then
Open Form1.CommonDialog1.filename For Output As #1
Print #1, Form1.ActiveControl.Text
Close #1
Form1.lblfilename.caption = Form1.CommonDialog1.filename
fstate.dirty = False
dialogerror:
If Err.Number <> 0 Then
Exit Sub
End If
End If
End If

End Sub
Sub Filequicksave()
Close #1

Dim strfilename As String

On Error GoTo Readonlyerr:
If Form1.lblfilename.caption <> "" Then
strfilename = Form1.lblfilename.caption
Open strfilename For Output As #1
Print #1, Form1.ActiveControl.Text
Close #1
Else
Form1.CommonDialog1.Filter = "Text documents (*.TXT) |*.TXT| INI Configuration Files (*.INI) |*.INI| Log Files (*.LOG) |*.LOG| All Files (*.*) |*.* "
Form1.CommonDialog1.DialogTitle = "Save As"
Form1.CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
On Error GoTo dialogerror
Form1.CommonDialog1.ShowSave
If Form1.CommonDialog1.filename <> "" Then
Open Form1.CommonDialog1.filename For Output As #1
Print #1, Form1.ActiveControl.Text
Close #1
Form1.lblfilename.caption = Form1.CommonDialog1.filename
Form1.caption = Form1.lblfilename.caption & " - TextPad"

Readonlyerr:
If Err.Number <> 0 Then
Readonlyerror
End If
End If
End If

dialogerror:
If Err.Number <> 0 Then
Exit Sub
End If
End Sub

Sub newfile()
Dim msg, response
If Form1.lblfilename.caption <> "" Then '2
 msg = "The Text in " _
 & Form1.lblfilename.caption & " File has changed" _
& Chr(13) & Chr(13) & "Do you wish too save The changes ?"
Else ' * If there isnt then Display The One Below
msg = "The Text in the untitled file has changed" _
& Chr(13) & Chr(13) & "Do you wish too save the changes?"
End If
   
Beep
response = MsgBox(msg, vbYesNoCancel + vbQuestion + vbDefaultButton2, "New File ")

Select Case response
Case vbYes     ' User chose Yes.
Filequicksave 'call quicksave procedure in modquicksave
Form1.caption = "Untitled - TextPad"
Form1.ActiveControl.Text = ""
fstate.dirty = False
Form1.lblfilename.caption = ""



Case vbNo ' user chose No.
Form1.caption = "Untitled - TextPad"
fstate.dirty = False
Form1.ActiveControl.Text = ""
Form1.lblfilename.caption = ""
Form1.CommonDialog1.filename = ("")
End Select
End Sub

Public Sub CreateKey(Hkey As Long, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long
' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!

lRegResult = RegCreateKey(Hkey, strPath, hCurKey)
If lRegResult <> ERROR_SUCCESS Then
MsgBox " an unknown Error has occured settings cannot be saved", vbCritical, "error"
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Sub
' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!

Public Function GetSettingString(Hkey As Long, _
strPath As String, strValue As String, Optional _
Default As String) As String
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

'Set up default value
If Not IsEmpty(Default) Then
GetSettingString = Default
Else
GetSettingString = ""
End If

lRegResult = RegOpenKey(Hkey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_SZ Then

strBuffer = String(lDataBufferSize, " ")
lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, _
ByVal strBuffer, lDataBufferSize)

intZeroPos = InStr(strBuffer, Chr$(0))
If intZeroPos > 0 Then
GetSettingString = Left$(strBuffer, intZeroPos - 1)
Else
GetSettingString = strBuffer
End If

End If

Else
MsgBox "An error has occured whiole calling the api function , settings cannot be saved .", vbCritical, "Error,api"

End If

lRegResult = RegCloseKey(hCurKey)
End Function
' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!

Public Sub SaveSettingString(Hkey As Long, strPath _
As String, strValue As String, strdata As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(Hkey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, _
ByVal strdata, Len(strdata))

If lRegResult <> ERROR_SUCCESS Then
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub ' Thank you again VBWORLD.com !!!!!!!
' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!


Sub Check_For_Registry_Entrys()

If GetSetting("Textpad", "Toolbar", "Visible") = "" _
And GetSetting("Textpad", "chckassociations", "show") = "" _
And GetSetting("Textpad", "Font", "Font") = "" _
And GetSetting("Textpad", "Font", "Fontsize") = "" _
And GetSetting("TextPad", "Wordwrap", "Wordwrap") = "" _
And GetSetting("TextPad", "UseExternalEditor", "Path") = "" Then
Load Frmnoreg
Frmnoreg.Show
 Form1.Hide
End If
End Sub

Sub form1_startup() ' Startup commands and declares
Dim onoffwordwrap As Boolean
On Error GoTo Inputpastendoffile:
Check_For_Registry_Entrys 'checks if textpad has any settings in the registry

RetrieveALLSettings 'get all the programs settings that were
'saved in the registry

If Command$ <> "" Then '  Check If Command Line Is being Used
If FileLen(Command$) >= 65000 Then

Select Case UseExternalEditor.use
Case True ' Use external editor
Query_TooBig (Command$)

Case False ' Dont use external editor
MsgBox Command$ & _
Chr(13) & "Is too large too be opened here" & vbNewLine & vbNewLine & "TIP : Be Sure too have " _
& vbNewLine & Chr$(34) & "Use external editor When opening files too large for textpad too open." & Chr$(34) & _
vbNewLine & "Enabled in the options Dialog .", vbExclamation, "Error,TextPad"
Form1.lblfilename.caption = ""
Close #1
Reset
Close #1
Exit Sub
End Select

Else
OpenFilecommandline ' If it is then Goto The Openfilecommandline Procedure In module 3
End If
End If



' togglewordwrap from usewordwraps value form modoptions BOOLEAN

Togglewordwrap (Usewordwrap)


'Declare Public Variables Boolean
fstate.dirty = False



'Recent files menu
If GetSetting("TextPad", "Recentfiles", "1") <= "" Then
Form1.mnurecentfile1.visible = False
Else
Form1.mnurecentfile1.caption = GetSetting("TextPad", "RecentFiles", "1")
Form1.mnurecentfile1.visible = True
Form1.line8.visible = True
End If

'Error Control
Inputpastendoffile:
If Err.Number <> 0 Then
Dim msg, style, response, title
msg = "TextPad has encountered the following error(s) while loading :" _
& vbCrLf & vbCrLf & "Error :  " & Err.Description & _
Chr(13) & Chr(10) & "source :  " & Err.Source & vbCrLf & vbCrLf & "Would you like too Continue loading Text pad anyway ?"
style = vbYesNo + vbExclamation + vbDefaultButton2
title = "TextPad"
response = MsgBox(msg, style, title)

Select Case response
Case vbYes
Close #1
Resume Next
Case vbNo
End

End Select

End If
End Sub
Sub Openfile()
    Close #1
Dim msg, response
Dim regval As String
   If fstate.dirty = False Then GoTo Fileopenproc:
   If fstate.dirty = True Then
    
    Dim caption, strfilerecent

If Form1.lblfilename.caption <> "" Then
msg = "The Text in " _
 & Form1.lblfilename.caption & " File has changed" _
  & Chr(13) & Chr(13) & "Do you wish too save The changes ?"
Else ' * If there isnt then Display The One Below
msg = "The Text in the untitled file has changed" _
& Chr(13) & Chr(13) & "Do you wish too save the changes?"
End If
End If

response = MsgBox(msg, vbYesNoCancel + vbExclamation + vbDefaultButton3, "TextPad")

Select Case response
Case vbYes
Filequicksave ' call procedure in module1
Form1.CommonDialog1.filename = ("")
Form1.lblfilename.caption = ""

Case vbNo
GoTo Fileopenproc:

Case vbCancel
Exit Sub
End Select


    
Fileopenproc:
    
    Reset
    Close #1 ' close before using it again
    FreeFile (1)
    Form1.CommonDialog1.Flags = Normal_Cdlogflags
    Form1.CommonDialog1.Cancelerror = True
    Form1.CommonDialog1.Filter = "Text Files (*.TXT) |*.TXT| Ini Files (*.INI) |*.INI| Log Files (*.LOG) |*.LOG| All Files (*.*) |*.*"
    Form1.CommonDialog1.DialogTitle = "Open File"
    On Error GoTo Cdlogerror:
    Form1.CommonDialog1.ShowOpen
    If Form1.CommonDialog1.filename <> "" Then
    Open Form1.CommonDialog1.filename For Binary Access Read As #1
        If FileLen(Form1.CommonDialog1.filename) > 65000 Then GoTo outofmemory:
        
          On Error GoTo outofmemory:
    Form1.ActiveControl.Text = Input(LOF(1), 1)
    fstate.dirty = False
    Form1.caption = UCase$(Form1.CommonDialog1.filename) & "  - TextPad"
   strfilerecent = Form1.CommonDialog1.filename
   Form1.mnurecentfile1.caption = strfilerecent
   SaveRegistryString "Recentfiles", "1", strfilerecent
  Form1.mnurecentfile1.visible = True
    Form1.line8.visible = True
     Form1.lblfilename.caption = Form1.CommonDialog1.filename

fstate.dirty = False
Close #1
Exit Sub

Cdlogerror:
If Err.Number = 32755 Then
Exit Sub
End If

End If

outofmemory: ' error That occurs when textpad runs out of memory

regval = GetSetting("TextPad", "UseExternaleditor", "Path", "")

If UseExternalEditor.use = True And regval = "" Then
No_Externaleditor_Detected
Form1.lblfilename.caption = ("")
Close #1
Reset
 fstate.dirty = False
Exit Sub
End If

Select Case UseExternalEditor.use
   Case True ' Case is True *******
Query_TooBig (Form1.CommonDialog1.filename)
Form1.lblfilename.caption = ("")
Close #1
Reset
 fstate.dirty = False
Exit Sub
   Case False ' Case is False ******
MsgBox Form1.CommonDialog1.filename _
& vbNewLine & "Is too large For TextPad too open." _
& vbNewLine & vbNewLine & "TIP : Be Sure too have " & vbNewLine & Chr$(34) & "Use external editor When opening files too large for textpad too open." & Chr$(34) & _
 vbNewLine & " Enabled in the options Dialog .", vbExclamation, "Error,TextPad"
Form1.lblfilename.caption = ("")
Close #1
Form1.ActiveControl.Text = ("")
Form1.caption = "Untitled - TextPad"
Reset
 fstate.dirty = False
Exit Sub
End Select
 End Sub
Sub recentfile1()
Close #1
Dim regfilename As String
Dim strfilerecent As String
'***********************************
'Declare RegFilename as A string
'Data Type : String
'We first Check If the file exists If
'The vba Command : Filedatetime([Expression]) doesnt
'find it Then an error Will
'occur and the error Control will
'Pick it up and Prompt the user about it.
'************************************
On Error GoTo Filenotfound: ' error control
 regfilename = GetSetting("TextPad", "RecentFiles", "1", strfilerecent)
FileDateTime (regfilename) ' check if file exists first
 Open GetSetting("TextPad", "RecentFiles", "1", strfilerecent) For Binary Access Read As #1
   On Error GoTo outofmemory:
   Form1.ActiveControl.Text = Input(LOF(1), 1)
Close #1
Form1.caption = regfilename & " - TextPad"
Form1.lblfilename.caption = regfilename
fstate.dirty = False

Filenotfound: ' error control
If Err.Number <> 0 Then

MsgBox "File not found" _
& vbNewLine & "The File may have been moved renamed or Deleted", vbCritical, "TextPad"

SaveRegistryString "Recentfiles", "1", ""
Form1.mnurecentfile1.visible = False
Form1.mnurecentfile1.caption = ""
Form1.caption = "Untitled - TextPad"
Form1.line8.visible = False

Exit Sub
End If

outofmemory: ' error That occurs when File overloads textpads limit
If Err.Number = 7 Then
  Form1.caption = "Untitled - TextPad"
 Form1.ActiveControl.Text = ("")
MsgBox regfilename & _
 vbNewLine & "Is too large Too be opened Here ", vbCritical, "Error"
SaveRegistryString "Recentfiles", "1", ""
Form1.lblfilename.caption = ("")
Close #1
Form1.CommonDialog1.filename = ("")
 fstate.dirty = False
Exit Sub
End If
End Sub
Sub ShellNewTextPad(Thewindowstyle As VbAppWinStyle)
On Error GoTo Filenotfound:
' this declares strapp as a private string that cant go out
' of scope
Dim strapp As String
  strapp = App.Path & "\" & App.EXEName
  Shell strapp, Thewindowstyle

Filenotfound:
 If Err.Number <> 0 Then
 MsgBox "TextPad Cannot find its Own executable File.", vbCritical, "Error, New instance"
 Exit Sub
End If

End Sub
Sub TextChangecontrol()
On Error GoTo outofmemoryerror:
If fstate.dirty = False Then
fstate.dirty = True
End If
outofmemoryerror:
If Err.Number <> 0 Then
MsgBox "TextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
& Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then TextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "TextPad "
Exit Sub
End If
End Sub
Sub SaveRegistryString(Thesection As String, thekey As String _
, thesetting As Variant)
SaveSetting "TextPad", Thesection, thekey, thesetting
End Sub
Sub ExecuteExternalEditor(Currentfilename As String)
Dim thepathname As String
Dim success As Long
  thepathname = GetSetting("TextPad", "UseExternalEditor", "path", "")
success = Shell(thepathname & Space(1) & Currentfilename, vbNormalFocus)
End Sub
Sub DetectExternalEditor()
' **************************************************
' Boy this Only Took Me About a Minute or so Too
' Write But It Actually WORKS !!!!
' What It Basically does is Use the Two Textboxes
' On the Dumby Form then Sets It Towards them in code
' then it Finds The String We Dont Want And Just
' Strips It too a Null Zero Length String ""
' And It actually works !!! Kudos For Me !!!!
' **************************************************
Dim retval As String
Dim ThisTextBox1
Dim ThisTextBox2
Set ThisTextBox1 = FrmDumbyform.TxtConvert1
Set ThisTextBox2 = FrmDumbyform.TxtConvert2
Dim SearchforWhat, Where
Dim i As Integer
retval = GetSettingString(HKEY_CLASSES_ROOT, _
"rtffile\shell\open\command", _
"", "")
ThisTextBox1.Text = Chr(34) & "%1" & Chr(34)
ThisTextBox2.Text = retval
For i = 1 To 3
SearchforWhat = ThisTextBox1.Text
Where = InStr(ThisTextBox2.Text, SearchforWhat)
If Where Then
     ThisTextBox2.SelStart = Where - 1  ' set selection start and
     ThisTextBox2.SelLength = Len(SearchforWhat)   ' set selection length.
     ThisTextBox2.SelText = ""
     retval = ThisTextBox2.Text
ExternalEditorPath = retval
SaveRegistryString "UseExternalEditor", "path", retval
End If
Next i
End Sub
Sub Query_TooBig(Thefilename As String)
Dim msg, response
msg = "This File Is Too large For TextPad Too Open." _
& vbNewLine & vbNewLine & "Would You Like The External Editor too Open it ?"
response = MsgBox(msg, vbDefaultButton3 + vbQuestion + vbYesNoCancel, "TextPad")
Select Case response
Case vbYes
ExecuteExternalEditor (Thefilename)
End
Case vbNo
Exit Sub
Case vbCancel
Exit Sub
End Select
End Sub
Sub No_Externaleditor_Detected()
Dim msg, response
msg = "This File is too Large for textpad too open. and also ," _
& vbNewLine & vbNewLine & "No External Editor Has been Detected" _
& vbNewLine & vbNewLine & "This means That If a file is too Large too open You will continue too see this message Until You Have TextPad Detect An External Editor." _
& vbNewLine & vbNewLine & "Do you wish Too Have TextPad Detect One For you Now ? (Recommended) "
response = MsgBox(msg, vbDefaultButton3 + vbExclamation + vbYesNo, "TextPad")
Select Case response
Case vbYes
DetectExternalEditor
MsgBox "External Editor Has been Succesfully Detected", vbInformation, "TextPad"
Exit Sub
Case vbNo
Exit Sub
End Select
End Sub
Sub Togglewordwrap(Optional ONOrOFF As Boolean)
On Error GoTo Err:
Dim fontname, fontsize As Integer
'onoff = IIf(us, True, False)
fontname = GetSetting("TextPad", "Font", "font")
fontsize = GetSetting("TextPad", "Font", "Fontsize")

Select Case ONOrOFF
Case False
Resizenotewithtoolbar
Form1.Text1.visible = False
Resizenotewithtoolbar

Form1.Text2.fontname = Form1.Text1.fontname
Form1.Text2.fontsize = Form1.Text1.fontsize

Form1.Text2.fontname = fontname
Form1.Text2.fontsize = fontsize

Form1.Text2.visible = True

If fstate.dirty = True Then
fstate.dirty = True
Else
If fstate.dirty = False Then
fstate.dirty = False
End If
End If

Resizenotewithtoolbar
Form1.Text2.Text = Form1.Text1.Text
If fstate.dirty = True Then
fstate.dirty = True
Else
If fstate.dirty = False Then
fstate.dirty = False
End If
End If

'Check off Menu too match current state
Form1.mnuwordwrap.Checked = False


Case True

'resize form too match current state
Resizenotewithtoolbar
Form1.Text2.visible = False
Resizenotewithtoolbar
'set font name and size from registry
Form1.Text1.fontname = Form1.Text2.fontname
Form1.Text1.fontsize = Form1.Text2.fontsize
Form1.Text1.visible = True
Form1.Text1.fontname = fontname
Form1.Text1.fontsize = fontsize
If fstate.dirty = True Then
fstate.dirty = True
Else
If fstate.dirty = False Then
fstate.dirty = False
End If
End If
Resizenotewithtoolbar
Form1.Text1.Text = Form1.Text2.Text
If fstate.dirty = True Then
fstate.dirty = True
Else
If fstate.dirty = False Then
fstate.dirty = False
End If
End If
Form1.mnuwordwrap.Checked = True



'error Handler
Err:
Resume Next
End Select

End Sub
