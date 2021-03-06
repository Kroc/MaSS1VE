VERSION 5.00
Begin VB.Form frmROM 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFAF00&
   BorderStyle     =   0  'None
   Caption         =   "MaSS1VE"
   ClientHeight    =   5895
   ClientLeft      =   -45
   ClientTop       =   -285
   ClientWidth     =   6945
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin MaSS1VE.bluBorderless bluBorderless 
      Height          =   480
      Left            =   6360
      TabIndex        =   5
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   847
   End
   Begin VB.Timer Shake 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5880
      Top             =   4680
   End
   Begin VB.Image imgDrop 
      Appearance      =   0  'Flat
      Height          =   2775
      Left            =   0
      OLEDropMode     =   1  'Manual
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drag & Drop a Sonic 1 ROM here to begin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      UseMnemonic     =   0   'False
      Width           =   5700
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1920
      Left            =   2400
      Picture         =   "frmROM.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1920
   End
   Begin VB.Label lblCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright � Kroc Camen #YEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5400
      Width           =   5385
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Master System Sonic 1 Visual Editor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   375
      TabIndex        =   1
      Top             =   840
      Width           =   3225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MaSS1VE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1635
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "v0.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFEABA&
      Height          =   210
      Left            =   6240
      TabIndex        =   4
      Top             =   5400
      Width           =   450
   End
End
Attribute VB_Name = "frmROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'FORM :: frmROM

'WARNING: Drag and Drop will not work if the app is running as Administrator (elevated, _
 on Vista and above). I may fix this in the future, but MaSS1VE doesn't need to run _
 elevated; you might run into this problem if you have set the VB6 IDE to run elevated _
 (not required for compatibility AFAIK) and you try running MaSS1VE from the IDE _
 <social.msdn.microsoft.com/Forums/windowsdesktop/en-US/0ccf84fd-b78d-45b3-9b79-7366003cb19d/wmdropfiles-in-an-elevated-application-administrator>

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'What set of UI to show depending on actions taken
Private My_UIState As frmROM_UIState

Public Enum frmROM_UIState
    Default = 0                 'Before drag and when drag leaves the form
    ROMGood = 1                 'Dragging over the form, the ROM is verified
    ROMBad = 2                  'Dragging over the form, not a file or not a ROM
    Importing = 3               'Currently importing the ROM
End Enum

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'Track mouse in / out events
Private WithEvents MouseEvents As bluMouseEvents
Attribute MouseEvents.VB_VarHelpID = -1

'When the user drags a ROM over the form we pre-verify that is indeed a Sonic 1 ROM _
 before they drop it so as to use responsive UI
Private ROMVerified As Boolean

'Where the Window was positioned before shaking
Private FormLeft As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'FORM Load _
 ======================================================================================
Private Sub Form_Load()
    'Attach the mouse tracking
    Set MouseEvents = New bluMouseEvents
    Call MouseEvents.Attach(Me.hWnd)
    
    'Set the colour scheme
    Me.BackColor = blu.ActiveColour
    'Version number label
    Let Me.lblVersion.ForeColor = blu.InertColour
    Let Me.lblVersion.Caption = Run.VersionString
    'Add the current year to the copyright message
    Let Me.lblCopy.Caption = Replace(Me.lblCopy.Caption, "#YEAR", Year(Now))
    
    'Load the 32-bit icon from the EXE
    Call blu.SetIcon(frmROM.hWnd, "AAA")
End Sub

'FORM Resize _
 ======================================================================================
Private Sub Form_Resize()
    'If the form is invisible or minimised then don't bother resizing
    If Me.WindowState = vbMinimized Or Me.Visible = False Then Exit Sub
    
    'Position the control box in the corner
    Let Me.bluBorderless.Left = Me.ScaleWidth - Me.bluBorderless.Width
    
    'We use an empty image control to cover the form so that we don't get multiple _
     drag in / out events if the user drags over labels &c. (The Z-ordering of this _
     control is important, it's above all the other controls, except the control box)
    Call Me.imgDrop.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)
End Sub

'FORM Unload _
 ======================================================================================
Private Sub Form_Unload(Cancel As Integer)
    'Detacth the mouse tracking
    Set MouseEvents = Nothing
End Sub

'EVENT imgDrop OLEDRAGDROP : A file has been dropped on the form _
 ======================================================================================
Private Sub imgDrop_OLEDragDrop( _
    ByRef Data As DataObject, _
    ByRef Effect As VBRUN.OLEDropEffectConstants, _
    ByRef Button As Integer, ByRef Shift As Integer, _
    ByRef X As Single, ByRef Y As Single _
)
    'The `imgDrop_OLEDragOver` event handles verifying the ROM before the user lets _
     go of the mouse, we only need to action the drop
    If ROMVerified = False Then Exit Sub
    
    'Change the UI for importing the ROM: _
     ----------------------------------------------------------------------------------
    Let UIState = Importing
    'Clear the copy cursor otherwise it'll hang on screen
    Let Effect = vbDropEffectNone
    'Refresh the screen
    DoEvents
    
    'Copy the ROM to App Data: _
     ----------------------------------------------------------------------------------
    '`Run.Main` already checks if the "App Data" folder exists and creates it if not. _
     We only need attempt to copy the ROM. This could fail if a previously portable _
     installation was moved to a non-user area of the disk, or if the portable media _
     is made read-only
    On Error GoTo Continue
    Call VBA.FileCopy(ROM.Path, Run.AppData & ROM.NameSMS)
    'Update the location of the ROM path used here-in
    Let ROM.Path = Run.AppData & ROM.NameSMS
    
Continue:
    'If we were not able to copy the ROM we can still continue, using the file's _
     original location, but the user will have to repeat this process every time _
     MaSS1VE is started. For now we will stave off an error message or other action
    On Error GoTo 0
    
    'Since we have no project management yet we just start one in memory. _
     Later on we will show a form where the user can create / load projects
    'TODO: Import errors will have to be handled gracefully
    Call ROM.Import
    'Show the main application form
    Load mdiMain
    Call mdiMain.Show
    
    Unload Me
End Sub

'EVENT imgDrop OLEDRAGOVER : A file is being dragged over the form _
 ======================================================================================
Private Sub imgDrop_OLEDragOver( _
    ByRef Data As DataObject, _
    ByRef Effect As VBRUN.OLEDropEffectConstants, _
    ByRef Button As Integer, ByRef Shift As Integer, _
    ByRef X As Single, ByRef Y As Single, _
    ByRef State As Integer _
)
    'When something is dragged onto the form, check it before they drop it so that _
     we can show the right mouse pointer
    
    'Don't check with every mouse move, just when it enters for the first time
    If State = VBRUN.DragOverConstants.vbEnter Then
        'Is there file(s) being dragged in?
        If Data.GetFormat(VBRUN.ClipBoardConstants.vbCFFiles) = True Then
            'The user should only drag one file, but if for some strange reason they _
             drag many, check all of them for validity
            Dim i As Long
            For i = 1 To Data.Files.Count
                'Check the file for validity
                If ROM.Verify(Data.Files(i)) Then
                    'If so, show the copy cursor, change the UI and await the drop
                    Let Effect = Effect And vbDropEffectCopy
                    Let ROMVerified = True
                    Let ROM.Path = Data.Files(i)
                    Let UIState = ROMGood
                    Exit Sub
                End If
            Next
            'No valid files, show our displeasure
            Let Effect = vbDropEffectNone
            Let ROMVerified = False
            Let ROM.Path = vbNullString
            Let UIState = ROMBad
        Else
            'The user is dragging in something other than file(s), this shall not pass
            Let Effect = vbDropEffectNone
            Let ROMVerified = False
            Let ROM.Path = vbNullString
            Let UIState = ROMBad
            
        End If
    
    'If the drag leaves the form, reset the UI
    ElseIf State = VBRUN.DragOverConstants.vbLeave Then
        'There's a _strange_ bug / bit of behaviour in that if the the `vbOver` state _
         below is handled (i.e. we set the effect value) then when the user drops _
         this state fires, but if there's no code below, it won't!!! What we want is _
         that when a user drops an invalid file the UI warning remains until the _
         user mouses out of the form. Therefore, we need to avoid resetting the UI _
         here if the drop was invalid
        If ROMVerified = True Then
            Let ROMVerified = False
            Let ROM.Path = vbNullString
            Let UIState = Default
        End If
    
    'During continuous mouse drag over, keep the drag icon set _
     it will default to copy, but we want to show the "No" cursor for invalid drags
    ElseIf State = VBRUN.DragOverConstants.vbOver Then
        If ROMVerified = False Then Let Effect = vbDropEffectNone
    
    End If
End Sub

'EVENT Shake TIMER : Do the shaking animation _
 ======================================================================================
Private Sub Shake_Timer()
    'Decrement the time remaining and stop the animation if complete
    Let Me.Shake.Tag = CInt(Me.Shake.Tag) - 1
    If Me.Shake.Tag = 0 Then Let Shake.Enabled = False
    'Move the form backwards and forwards alternatively
    Let Me.Left = FormLeft + IIf(Me.Shake.Tag Mod 2 = 0, 30, -30)
End Sub

'EVENT MouseEvents MOUSEOUT _
 ======================================================================================
Private Sub MouseEvents_MouseOut()
    'If something other than a valid ROM is dropped on the form the UI is changed to _
     warn the user. We want to revert this once the mouse moves back out of the window
    If UIState = ROMBad Then Let UIState = Default
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY UIState : Manage the UI changes between drag/drop actions _
 ======================================================================================
Public Property Get UIState() As frmROM_UIState: Let UIState = My_UIState: End Property
Public Property Let UIState(ByVal State As frmROM_UIState)
    
    'During importing various elements are hidden / disabled
    Let Me.imgDrop.Enabled = (State <> Importing)       'Enable/Disable Drag-and-drop
    Let Me.bluBorderless.Visible = (State <> Importing) 'Show/Hide control box
    'During importing you get the hourglass cursor
    Let Me.MousePointer = IIf( _
        State = Importing, _
        VBRUN.MousePointerConstants.vbHourglass, _
        VBRUN.MousePointerConstants.vbDefault _
    )
    
    Select Case State
        Case frmROM_UIState.Default:
            Let Me.lblStatus.Caption = "Drag & Drop a Sonic 1 ROM to begin"
        
        Case frmROM_UIState.ROMBad:
            Let Me.lblStatus.Caption = "Sorry, that's not a Sonic 1 Master System ROM"
            'Give some visual feedback
            Call ShakeForm
            
        Case frmROM_UIState.ROMGood:
             Let Me.lblStatus.Caption = "ROM OK. Drop to begin"
            
        Case frmROM_UIState.Importing:
            Let Me.lblStatus = "Importing Sonic 1 ROM..."
    
    End Select
    
    Let My_UIState = State
End Property

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'ShakeForm : Show our disapproval :| _
 ======================================================================================
Private Sub ShakeForm()
    'Record the current position of the form so we can jiggle it around this point _
     and return to where it was originally afterwards
    Let FormLeft = Me.Left
    'This will be the length of the shake effect
    Let Me.Shake.Tag = 5
    'Begin the animation, see `Shake_Timer` event for details
    Let Me.Shake.Enabled = True
End Sub
