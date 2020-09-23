Attribute VB_Name = "Modoptions"
Type ToolbarBOOL
visible As Integer
End Type
Public isToolbarvisible As ToolbarBOOL

Type ExternalEditor
use As Integer
End Type
Public UseExternalEditor As ExternalEditor

Public Usewordwrap As Boolean

Type Checkassociations
check As Integer
End Type
Public Check_Associations_at_Startup As Checkassociations
'***************

Sub SaveSetting_Toolbar(BOOLobject As Object)
Dim retval As Long
retval = IIf(BOOLobject, True, False)
Select Case retval
Case True
SaveRegistryString "Toolbar", "Visible", 1
' Toolbar1 Will be Visible At startup
Form1.Toolbar.visible = True
Form1.mnuhidetoolbaritem.Checked = True
Call Resizenotewithtoolbar

Case False
SaveRegistryString "Toolbar", "Visible", 0
If Form1.Toolbar.visible = True Then
Form1.mnuhidetoolbaritem.Checked = False
Form1.Toolbar.visible = False
Call Resizenotewithtoolbar
End If

End Select

End Sub

Sub saveSetting_UseExternalEditor(BOOLobject As Object)
Dim retval As Long
retval = IIf(BOOLobject, True, False)
Select Case retval
Case True
SaveRegistryString "UseExternalEditor", "Use", 1
Case False
SaveRegistryString "UseExternalEditor", "Use", 0
End Select
End Sub
 
Sub SaveSetting_Wordwrap(BOOLobject As Object)
Dim retval As Long
Dim fontname As String, fontsize As Integer
fontname = GetSetting("Textpad", "Font", "Font", "")
fontsize = GetSetting("Textpad", "Font", "Fontsize", "")

retval = IIf(BOOLobject, True, False)
Select Case retval
Case True
SaveRegistryString "Wordwrap", "Wordwrap", 1
Form1.Text1.fontname = fontname
Form1.Text1.fontsize = fontsize
Resizenotewithtoolbar
Form1.Text1.visible = True
Form1.mnuwordwrap.Checked = True
Resizenotewithtoolbar

Case False
SaveRegistryString "Wordwrap", "Wordwrap", 0
Form1.Text2.fontsize = fontsize
Form1.Text2.fontname = fontname
Resizenotewithtoolbar
Form1.mnuwordwrap.Checked = False
Form1.Text1.visible = False
Form1.Text2.visible = True
Resizenotewithtoolbar

End Select

End Sub

Sub SaveSetting_chckassociations(BOOLobject As Object)
Dim retval As Long
retval = IIf(BOOLobject, True, False)
Select Case retval
Case True
SaveRegistryString "chckassociations", "show", 1
Case False
SaveRegistryString "chckassociations", "show", 0
End Select
End Sub

Sub RetrieveALLSettings()
Dim externaleditorval As String
Dim toolbarval As String
Dim wordwrapval As String
Dim Chckassociationsval As String
' Toolbar
toolbarval = GetSetting("TextPad", "Toolbar", "Visible")
Select Case toolbarval
Case 1
isToolbarvisible.visible = True
Form1.Toolbar.visible = True
Form1.mnuhidetoolbaritem.Checked = True
Case 0
isToolbarvisible.visible = False
Form1.Toolbar.visible = False
Form1.mnuhidetoolbaritem.Checked = False
End Select

' External Editor
externaleditorval = GetSetting("TextPad", "UseExternalEditor", "Use")
Select Case externaleditorval
Case 1
UseExternalEditor.use = True
Case 0
UseExternalEditor.use = False
End Select

'Word wrap
wordwrapval = GetSetting("TextPad", "Wordwrap", "Wordwrap")
Select Case wordwrapval
Case 1
Usewordwrap = True
Case 0
Usewordwrap = False
End Select

' check Associations
Chckassociationsval = GetSetting("TextPad", "chckassociations", "show")
Select Case Chckassociationsval
Case 1
Check_Associations_at_Startup.check = True
check_if_textpad_is_associated
Case 0
Check_Associations_at_Startup.check = False
End Select
End Sub
