Attribute VB_Name = "Run"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Run

'Where execution begins. Also, generic stuff for the whole app

'Avoid having to search and replace these strings:
Public Const INI_Name = "MaSS1VE.ini"

Public Const UpdateFile = "Update.ini"
Public Const UpdateURL = "http://localhost/mass1ve/" & UpdateFile

'We need to know what action was taken on the update form after it was closed
Public UpdateResponse As VBA.VbMsgBoxResult

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
    
    'Begin Logging _
     ----------------------------------------------------------------------------------
    'Create the App Data folder if it doesn't exist so that the update _
     check won't fail
    If Lib.DirExists(Run.AppData) = False Then
        On Error GoTo ReadOnly
        'Attempt to create the "App Data" folder
        Call VBA.MkDir(Run.AppData)
        On Error GoTo 0
    End If
    
    'TODO: Open log file here
    
    'Has an update been downloaded? _
     ----------------------------------------------------------------------------------
    'The main form (`mdiMain`) downloads updates and displays a button to launch them, _
     if the user closes the app without launching the update then we will launch the _
     update automatically the next time the app is started
    
    'TODO: If a project file was double-clicked, we need to pass this through the _
           updater so that it gets loaded after the update
    
    'Check the necessary update files have been downloaded...
    If Run.UpdateWaiting = True Then
        'Display the update/changelog window and wait for a response
        Load frmUpdate
        Call frmUpdate.Show(vbModal)
        'If the user clicked to "Exit & Update" then do so now
        If Run.UpdateResponse = vbOK Then
            'Launch the installer with the path to the installation
            Call WIN32.shell32_ShellExecute( _
                0, vbNullString, Run.AppData & "Update.exe", _
                "/UPDATE /D=" & Left$(Run.Path, Len(Run.Path) - 1), _
                Run.AppData, SW_SHOWNORMAL _
            )
            'Quit the application _
             (we haven't shown any other UI so exiting `sub Main` will do)
            Exit Sub
        End If
    End If
    
    'Check for Sonic 1 ROM _
     ----------------------------------------------------------------------------------
    'MaSS1VE requires access to an original Sonic 1 ROM when starting a new project _
     or exporting to a new ROM. Rather than just save a path to a ROM file located _
     somewhere in the user's files (which might get moved), we will keep a copy in _
     the app path's "App Data" folder so there's less chance of it going missing.
    'MaSS1VE doesn't come with a Sonic 1 ROM, the user has to provide their own, _
     so we display a form where they can drag-and-drop one if we can't find it
    
    'Check if the ROM is already where we expect it
    If Lib.FileExists(Run.AppData & ROM.NameSMS) = True Then
        'WARNING: Incredibly, prefixing `NameSMS` with it's module, `ROM`, causes the _
         compiler to crash. It's insane, yes. This is the only reference on the web _
         I found about this rare bug: <bbs.csdn.net/topics/30000137>
        Let ROM.Path = Run.AppData & NameSMS
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
        
        'Rip the ROM data into an in-memory MaSS1VE project
        Call ROM.Import
        
        'Launch the main UI
        Load mdiMain
        Unload frmROM
        Call mdiMain.Show
    End If
    
    Exit Sub

ReadOnly:
    'If we can't write the log file, don't go any further, we can't operate in a _
     read-only environment
     MsgBox _
        "Cannot write to " & Chr(34) & Run.AppData & Chr(34) & ". " _
        & "MaSS1VE cannot be run from a folder it does not have write permissions " _
        & "for. (e.g. read-only media such as CD or an Administrator owned folder " _
        & "such as " & Chr(34) & "Program Files" & Chr(34) & ")", _
        vbCritical Or vbOKOnly
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
    Let Path = Lib.EndSlash(App.Path) & IIf(Run.InIDE, "RELEASE\", vbNullString)
End Property

'PROPERTY UpdateWaiting : Check if an update has already been downloaded _
 ======================================================================================
Public Property Get UpdateWaiting() As Boolean
    If Lib.FileExists(Run.AppData & "Update.ini") = True Then
        If Lib.FileExists(Run.AppData & "Update.html") = True Then
            If Lib.FileExists(Run.AppData & "Update.exe") = True Then
                Let UpdateWaiting = True
            End If
        End If
    End If
End Property

'PROPERTY VersionString : A friendly version number displayed in some places _
 ======================================================================================
Public Property Get VersionString() As String
    Let VersionString = _
        "v" & Format(App.Major & "." & App.Minor, "##0.0#") & _
        "," & App.Revision & " pre-alpha"
End Property
