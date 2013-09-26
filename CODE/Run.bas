Attribute VB_Name = "Run"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Run

'Where execution begins. Also, generic stuff for the whole app

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
     (`Sub Main()` will only be run when your app runs, not during design time)
    Let blu.UserMode = True
        
    'Check for Sonic 1 ROM _
     ----------------------------------------------------------------------------------
    'MaSS1VE requires access to an original Sonic 1 ROM when starting a new project _
     or exporting to a new ROM. Rather than just save a path to a ROM file located _
     somewhere in the user's files (which might get moved), we will keep a copy in _
     the app path's data folder so there's less chance of it going missing.
    'MaSS1VE doesn't come with a Sonic 1 ROM, the user has to provide their own, _
     so display a form where they can drag-and-drop one if we can't find it
    
    'When a ROM is provided, it's copied to the "App Data" folder in the app directory, _
     test if it's currently there:
    If Lib.FileExists(Run.AppData & ROM.NameSMS) = True Then
        Let ROM.Path = Run.AppData & ROM.NameSMS
    End If
    
    'If no ROM was found, ask the user for it
    If ROM.Path = vbNullString Then
        Load frmROM
        Call frmROM.Show
    
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

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY AppData : The path where the application data is stored (ROM, Updates) _
 ======================================================================================
Public Property Get AppData() As String: Let AppData = Run.Path & "App Data\": End Property

'PROPERTY UserData : The path where the user data is stored (Project files) _
 ======================================================================================
Public Property Get UserData() As String: Let UserData = Run.Path & "User Data\": End Property

'PROPERTY InIDE : Are we running the code from the Visual Basic IDE? _
 ======================================================================================
Public Property Get InIDE() As Boolean
    On Error GoTo Err_True
    
    'Do something that only faults in the IDE
    Debug.Print 1 \ 0
    InIDE = False
    Exit Property

Err_True:
    InIDE = True
End Property

'PROPERTY Path : Like `App.Path` but normalised for IDE / EXE _
 ======================================================================================
Public Property Get Path() As String
    'Set `Run.Path` so that program output goes to the RELEASE folder when in IDE
    Let Path = Lib.EndSlash(App.Path) & IIf(Run.InIDE, "RELEASE\", "")
End Property

'PROPERTY VersionString : A friendly version number displayed in some places _
 ======================================================================================
Public Property Get VersionString() As String
    Let VersionString = _
        "v" & Format(App.Major & "." & App.Minor, "##0.0#") & _
        " #" & App.Revision & " pre-alpha"
End Property
