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
   Begin VB.Timer Shake 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5880
      Top             =   5280
   End
   Begin MaSS1VE.bluControlBox cbxMin 
      Height          =   480
      Left            =   6000
      TabIndex        =   5
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   1
      Kind            =   1
   End
   Begin MaSS1VE.bluControlBox cbxClose 
      Height          =   480
      Left            =   6480
      TabIndex        =   4
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   1
   End
   Begin MaSS1VE.bluWindow bluWindow1 
      Left            =   6360
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
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
      Caption         =   "Copyright © Kroc Camen #YEAR"
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
      Width           =   6465
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
End
Attribute VB_Name = "frmROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'FORM :: frmROM

'WARNING: Drag and Drop will not work if the app is running as Administrator (elevated, _
 on Vista and above). I may fix this in the future, but MaSS1VE doesn't need to run _
 elevated; you might run into this problem if you have set the VB6 IDE to run elevated _
 (not required for compatibility AFAIK) and you try running MaSS1VE from the IDE _
 <social.msdn.microsoft.com/Forums/windowsdesktop/en-US/0ccf84fd-b78d-45b3-9b79-7366003cb19d/wmdropfiles-in-an-elevated-application-administrator>

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'When the user drags a ROM over the form we pre-verify that is indeed a Sonic 1 ROM _
 before they drop it so as to use responsive UI
Dim ROMVerified As Boolean

'What set of UI to show depending on actions taken
Dim My_UIState As frmROM_UIState
Private Enum frmROM_UIState
    Default = 0                 'Before drag and when drag leaves the form
    ROMGood = 1                 'Dragging over the form, the ROM is verified
    ROMBad = 2                  'Dragging over the form, not a file or not a ROM
End Enum

'Where the Window was positioned before shaking
Dim FormLeft As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'FORM Load _
 ======================================================================================
Private Sub Form_Load()
    'Set the colour scheme
    Me.BackColor = blu.ActiveColour
    'Add the current year to the copyright message
    Let Me.lblCopy.Caption = Replace(Me.lblCopy.Caption, "#YEAR", Year(Now))
End Sub

'FORM Resize _
 ======================================================================================
Private Sub Form_Resize()
    'We use an empty image control to cover the form so that we don't get multiple _
     drag in / out events if the user drags over labels &c. (The Z-ordering of this _
     control is important, it's above all the other controls, except the control box)
    Call Me.imgDrop.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)
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
    If ROMVerified = True Then
        'Change the UI for importing the ROM:
        
        'Disable the drag and drop
        Let imgDrop.Enabled = False
        'Change the message
        Let Me.lblStatus = "Importing Sonic 1 ROM..."
        'Hide the control box buttons
        Let Me.cbxClose.Visible = False
        Let Me.cbxMin.Visible = False
        'Clear the copy cursor otherwise it'll hang on screen
        Let Effect = vbDropEffectNone
        'Set the busy cursor instead
        Let Me.MousePointer = VBRUN.MousePointerConstants.vbHourglass
        'Refresh the screen
        DoEvents
        
        'TODO: Copy the ROM to the app data
        'Since we have no project management yet we just start one in memory. _
         Later on we will show a form where the user can create / load projects
        'TODO: Import errors will have to be handled gracefully
        Call ROM.Import
        'Show the main application form
        Load mdiMain: Call mdiMain.Show
        'Since there's no level select screen yet, jump right into the editing
        Load frmEditor: Call frmEditor.Show
        
        Unload Me
    End If
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
            Next i
            'No valid files, show our displeasure
            Let Effect = vbDropEffectNone
            Let ROMVerified = False
            Let ROM.Path = vbNullString
            Let UIState = ROMBad
        Else
            'The user is dragging in something other than file(s), this shall not pass
            Let ROMVerified = False
            Let ROM.Path = vbNullString
            Let UIState = ROMBad
            
        End If
    
    'If the drag leaves the form, reset the UI
    ElseIf State = VBRUN.DragOverConstants.vbLeave Then
        Let ROMVerified = False
        Let ROM.Path = vbNullString
        Let UIState = Default
        
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

'/// PRIVATE PROPERTIES ///////////////////////////////////////////////////////////////

'PROPERTY UIState : Manage the UI changes between drag/drop actions _
 ======================================================================================
Private Property Get UIState() As frmROM_UIState: Let UIState = My_UIState: End Property
Private Property Let UIState(ByVal State As frmROM_UIState)
    Select Case State
        Case frmROM_UIState.Default:
            Let Me.lblStatus.Caption = "Drag & Drop a Sonic 1 ROM to begin"
        
        Case frmROM_UIState.ROMBad:
            Let Me.lblStatus.Caption = "Sorry, that's not a Sonic 1 Master System ROM"
            'Give some visual feedback
            Call ShakeForm
            
        Case frmROM_UIState.ROMGood:
             Let Me.lblStatus.Caption = "ROM OK. Drop to begin"
             
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
