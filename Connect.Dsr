VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10050
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   12135
   _ExtentX        =   21405
   _ExtentY        =   17727
   _Version        =   393216
   Description     =   "Correctly indents the code in a VB code window."
   DisplayName     =   "VB Code Auto Indenter"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'* Copyright (c) 2005 by Prakah Patel
'*
'* This software is the proprietary information of Pd Systems.
'* Use is subject to license terms.
'*
'* @author  Prakash Patel
'* @version 1.0
'* @date    31 March 2004
'*
'***************************************************************************

Option Explicit

Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1



Sub Run()
    Dim mCode As CodeModule
    
    If Not (VBInstance Is Nothing) Then
    
        On Error Resume Next
        Set mCode = VBInstance.ActiveCodePane.CodeModule
        
        If mCode Is Nothing Then
            MsgBox "You must be in a code pane to auto indent it.", vbInformation, "AutoIndent"
        Else
            Call IndentCode(mCode)
        End If
    Else
        MsgBox "NO VB IDE!", vbExclamation
    End If
    
End Sub

Private Sub AddinInstance_OnAddInsUpdate(custom() As Variant)
' Don't delete me!
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Call Run
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("Auto Indent Code")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
' Don't delete me!
End Sub


'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call Run
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarButton  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    ' Set the picture of the button by copying resource
    ' and then pasting it on to it.
    Clipboard.Clear
    Clipboard.SetData LoadResPicture("MENUPIC", vbResBitmap), vbCFBitmap
    cbMenuCommandBar.PasteFace
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

