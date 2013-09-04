Attribute VB_Name = "Run"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Run

'Where execution begins. Also, generic stuff for the whole app

'/// PUBLIC VARS //////////////////////////////////////////////////////////////////////

'Like `App.Path`, but the same place ("BUILD" folder) for MaSS1VE in IDE / compiled
Public Path As String

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY VersionString : A friendly version number displayed in some places _
 ======================================================================================
Public Property Get VersionString() As String
    Let VersionString = _
        "v" & App.Major & "." & App.Minor & "." & App.Revision & " pre-alpha"
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'MAIN : It all starts here! _
 ======================================================================================
Private Sub Main()
    Debug.Print "BEGIN"
    
    'When a user control is nested in another user control, the `Ambient.UserMode` _
     property returns the incorrect value of True when the control is being run in _
     "Design Mode" (on the form editor). This would cause the design mode controls _
     to be subclassed and crashes the IDE. To stop this, the variable below will _
     always be False when the controls are running in Design Mode. We set `UserMode` _
     to True in the `Sub Main()` to tell the controls it's okay to subclass. _
     (`Sub Main()` will only be run when you your app runs, not during design time)
    Let blu.UserMode = True
    
    'Set `Run.Path` so that program output goes to the BUILD folder when in IDE
    Let Run.Path = Lib.EndSlash(App.Path) & IIf(Run.InIDE, "BUILD\", "")
        
    'Check for Sonic 1 ROM _
     ----------------------------------------------------------------------------------
    'MaSS1VE requires access to an original Sonic 1 ROM when starting a new project _
     or exporting to a new ROM. Rather than just save a path to a ROM file located _
     somewhere in the user's files (which might get moved), we will keep a copy in _
     the app path or the user's app data so there's less chance of it going missing
    'MaSS1VE doesn't come with a Sonic 1 ROM, the user has to provide their own, so _
     display a form where they can drag-and-drop one if we can't find it
    
    'Check the two common locations for a ROM
    'WARNING: Incredibly, prefixing `Locate` with it's module, `ROM`, causes the _
     compiler to crash. It's insane, yes. This is the only reference on the web _
     I found about this rare bug: <bbs.csdn.net/topics/30000137>
    Let ROM.Path = Locate()
    
    'If no ROM was found, ask the user for it
    If ROM.Path = vbNullString Then
        Load frmROM: Call frmROM.Show
    
    Else
        'We have the ROM, we can start MaSS1VE proper
        'NOTE: At the moment we don't have UI for starting / loading projects, _
         so we will utilise the ROM form for now
        Load frmROM
        Let frmROM.UIState = Importing
        Call frmROM.Show
        
        Call ROM.Import
        Load mdiMain
        Call mdiMain.Show
        
        Unload frmROM
    End If
End Sub

'InIDE : Are we running the code from the Visual Basic IDE? _
 ======================================================================================
Public Function InIDE() As Boolean
    On Error GoTo Err_True
    
    'Do something that only faults in the IDE
    Debug.Print 1 \ 0
    InIDE = False
    Exit Function

Err_True:
    InIDE = True
End Function
