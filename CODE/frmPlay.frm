VERSION 5.00
Begin VB.Form frmPlay 
   BackColor       =   &H00FFAF00&
   Caption         =   "Play"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MaSS1VE.bluControlBox cbxSizer 
      Height          =   360
      Left            =   14760
      TabIndex        =   3
      Top             =   8040
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Style           =   1
      Kind            =   3
   End
   Begin VB.PictureBox picROM 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFAF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3495
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin VB.Label lblExporting 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   5580
      End
      Begin VB.Label lblWhatToDo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drag and drop the cartridge below to a folder, or double-click it to launch it in an emulator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   4860
      End
      Begin VB.Image imgROM 
         Appearance      =   0  'Flat
         Height          =   1920
         Left            =   1800
         Picture         =   "frmPlay.frx":0000
         Stretch         =   -1  'True
         Top             =   960
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label lblFileName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sonic_the_Hedgehog_MaSS1VE.sms"
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
         Height          =   330
         Left            =   0
         TabIndex        =   1
         Top             =   3120
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   5580
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'FORM :: frmPlay

'ROM export screen

'WARNING: Drag and Drop will not work if the app is running as Administrator (elevated, _
 on Vista and above). I may fix this in the future, but MaSS1VE doesn't need to run _
 elevated; you might run into this problem if you have set the VB6 IDE to run elevated _
 (not required for compatibility AFAIK) and you try running MaSS1VE from the IDE _
 <social.msdn.microsoft.com/Forums/windowsdesktop/en-US/0ccf84fd-b78d-45b3-9b79-7366003cb19d/wmdropfiles-in-an-elevated-application-administrator>

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'This will hold the path to the temporary ROM we've cretaed
Private TempFile As String

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'FORM Activate : _
 ======================================================================================
Private Sub Form_Activate()
    'Show as busy
    Let Me.MousePointer = VBRUN.MousePointerConstants.vbHourglass
    'Force a screen refresh here otherwise the form hangs until the export is complete
    DoEvents
    'Export to the temporary folder, when the user drags and drops, it will copy it
    Let TempFile = WIN32.GetTemporaryFolder() & "Sonic_the_Hedgehog_MaSS1VE.sms"
    'Export the ROM
    Call ROM.Export(TempFile)
    'Ready
    Let Me.lblExporting.Visible = False
    Let Me.lblWhatToDo.Visible = True
    Let Me.lblFileName.Visible = True
    Let Me.imgROM.Visible = True
    Let Me.MousePointer = VBRUN.MousePointerConstants.vbDefault
End Sub

'FORM Resize _
 ======================================================================================
Private Sub Form_Resize()
    'If the form is invisible or minimised then don't bother resizing
    If Me.WindowState = vbMinimized Or Me.Visible = False Then Exit Sub
    'Ensure that the MDI child form always stays maximised when changing windows
    If Me.WindowState <> vbMaximized Then Let Me.WindowState = vbMaximized: Exit Sub
    
    'Position the sizing box in the corner
    Call Me.cbxSizer.Move( _
        Me.ScaleWidth - Me.cbxSizer.Width, Me.ScaleHeight - Me.cbxSizer.Height _
    )
    
    'Centre the cartridge icon and message, though slightly up looks better
    Call Me.picROM.Move( _
        (Me.ScaleWidth - Me.picROM.Width) \ 2, _
        ((Me.ScaleHeight - Me.picROM.Height) \ 2) - lblWhatToDo.Height _
    )
End Sub

'EVENT imgROM DBLCLICK : Launch the ROM in an emulator _
 ======================================================================================
Private Sub imgROM_DblClick()
    'Change the mouse pointer to background-working whilst we launch the ROM
    Let mdiMain.MousePointer = VBRUN.MousePointerConstants.vbArrowHourglass
    'Open the ROM file using explorer, if no file association is set, _
     the 'choose a program' window should appear
    Call WIN32.shell32_ShellExecute( _
        mdiMain.hWnd, vbNullString, TempFile, vbNullString, vbNullString, _
        SW_SHOWNORMAL _
    )
    'Reset the mouse pointer back to normal
    Let mdiMain.MousePointer = VBRUN.MousePointerConstants.vbDefault
End Sub

'EVENT imgROM MOUSEMOVE _
 ======================================================================================
Private Sub imgROM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Initiate the drag and drop
    If Button = VBRUN.MouseButtonConstants.vbLeftButton Then
        Call imgROM.OLEDrag
    End If
End Sub

'EVENT imgROM OLESTARTDRAG : The user has begun dragging the ROM off the form _
 ======================================================================================
Private Sub imgROM_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Let AllowedEffects = VBRUN.OLEDropEffectConstants.vbDropEffectCopy
    Call Data.Files.Add(TempFile)
    Call Data.SetData(, VBRUN.ClipBoardConstants.vbCFFiles)
End Sub
