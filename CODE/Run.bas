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
    Let Run.Path = App.Path & IIf(Run.InIDE, "\BUILD\", "\")
        
    'Allow Windows to theme VB's controls
    'NOTE: This works because "CompiledInResources.res" contains a manifest file, _
     see <www.vbforums.com/showthread.php?606736-VB6-XP-Vista-Win7-Manifest-Creator>
    Call WIN32.InitCommonControls( _
        ICC_STANDARD_CLASSES Or ICC_INTERNET_CLASSES _
    )
        
    'Check for Sonic 1 ROM _
     ----------------------------------------------------------------------------------
'    Load frmROM: Call frmROM.Show

    Let ROM.Path = Run.Path & "Sonic the Hedgehog (1991)(Sega).sms"
    Call ROM.Import
    
    Load mdiMain: Call mdiMain.Show
    Load frmEditor: Call frmEditor.Show
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
