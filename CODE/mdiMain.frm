VERSION 5.00
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFAF00&
   Caption         =   "MaSS1VE"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHelp 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   11460
      ScaleHeight     =   6855
      ScaleWidth      =   3660
      TabIndex        =   1
      Top             =   990
      Visible         =   0   'False
      Width           =   3660
      Begin MaSS1VE.bluWebView bluWebView 
         Height          =   1575
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2778
      End
      Begin VB.PictureBox picHelpToolbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFAF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   3735
         TabIndex        =   2
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.PictureBox toolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   15120
      TabIndex        =   0
      Top             =   0
      Width           =   15120
      Begin MaSS1VE.bluButton btnUpdate 
         Height          =   480
         Left            =   12840
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   847
         Caption         =   "UPDATE!"
         State           =   1
      End
      Begin MaSS1VE.bluControlBox cbxClose 
         Height          =   480
         Left            =   14640
         TabIndex        =   4
         Top             =   0
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin MaSS1VE.bluTab bluTab 
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   495
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
      End
      Begin MaSS1VE.bluButton btnHelp 
         Height          =   495
         Left            =   14160
         TabIndex        =   7
         Top             =   495
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "HELP"
      End
      Begin MaSS1VE.bluControlBox cbxMin 
         Height          =   480
         Left            =   13920
         TabIndex        =   5
         Top             =   0
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         Kind            =   1
      End
      Begin MaSS1VE.bluControlBox cbxMax 
         Height          =   480
         Left            =   14280
         TabIndex        =   6
         Top             =   0
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         Kind            =   2
      End
      Begin MaSS1VE.bluLabel lblMaSS1VE 
         Height          =   495
         Left            =   3840
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   873
         Caption         =   "MaSS1VE: The Master System Sonic 1 Visual Editor"
         Enabled         =   0   'False
      End
      Begin MaSS1VE.bluLabel lblTip 
         Height          =   495
         Left            =   10320
         Top             =   495
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
         Alignment       =   1
         Caption         =   "The quick brown fox jumps over the lazy dog"
         Enabled         =   0   'False
      End
      Begin MaSS1VE.bluLabel lblVersion 
         Height          =   480
         Left            =   12000
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   847
         Alignment       =   1
         Caption         =   "v0.0.0"
         Enabled         =   0   'False
      End
   End
   Begin MaSS1VE.bluDownload bluDownload 
      Left            =   720
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MaSS1VE.bluWindow bluWindow 
      Left            =   120
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "mdiMain"
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
'FORM :: mdiMain

'This is the main application form.

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'MDIFORM Load _
 ======================================================================================
Private Sub MDIForm_Load()
    'Load the 32-bit icon from the EXE
    Call WIN32.SetIcon(mdiMain.hWnd, "AAA")
    
    'Set the minimum allowed size of the form
    Let Me.bluWindow.MinWidth = 512
    Let Me.bluWindow.MinHeight = 320
    
    'Make it so that the window can be dragged via the top area
    Call Me.bluWindow.RegisterMoveHandler(Me.toolbar)
    
    'Set the version number label
    Let Me.lblVersion.Caption = Run.VersionString
    
    'Clear the help tip message _
     (shows contextual help when mousing over things)
    Call SetTip
    
    'Apply colour scheme
    Call Me.SetTheme
    
    'Configure the tab strip
    With Me.bluTab
        .Border = False
        .AutoSize = True
        .TabCount = 2
        .Caption(0) = "LEVELS"
        .Caption(1) = "PLAY"
        'Select no tab to begin with, the welcome screen will be shown by default
        .CurrentTab = -1
    End With
    
    'If on a small screen, start up maximised (we need at least 1024x600)
    If Screen.Width \ Screen.TwipsPerPixelX <= 1024 Then
        Let mdiMain.WindowState = VBRUN.FormWindowStateConstants.vbMaximized
    End If
    
    'Load the welcome form into the MDI window so the user has something to look at
    Load frmWelcome
    Call frmWelcome.Show
    
    'Check for updates: _
     ----------------------------------------------------------------------------------
    'Has an update already been downloaded and not installed yet?
    If Run.UpdateWaiting = True Then
        'Display the button for the update
        Let Me.lblVersion.Visible = False
        Let Me.btnUpdate.Visible = True
    Else
        'Access the "MaSS1VE.ini" file in the App Data folder. _
         It won't matter if it's missing, the class will just return default values
        Dim INI As INIFile
        Set INI = New INIFile
        Let INI.FilePath = Run.AppData & Run.INI_Name
    
        'Has an update check been done in the last day?
        If DateDiff("d", _
            CDate(INI.GetDouble("LastUpdateCheck", "Update")), Now() _
        ) > 1 Then
            'Download the "Update.ini" file. This is asynchronous, so the code will _
             not sit here waiting. The bluDownload control `Complete` event will fire _
             once the file is received so go there to follow the update process
            Let Me.bluDownload.Tag = Run.UpdateFile
            Call Me.bluDownload.Download( _
                Run.UpdateURL, _
                Run.AppData & Run.UpdateFile, vbAsyncReadForceUpdate _
            )
        End If
        Set INI = Nothing
    End If
End Sub

'MDIFORM Reisze _
 ======================================================================================
Private Sub MDIForm_Resize()
    'Resizing code will freak out if we're minimised!
    If Me.WindowState = vbMinimized Or Me.Visible = False Then Exit Sub
    
    'The dimensions for aligned controls on an MDIForm are *completely* unreliable. _
     We'll use the WIN32 API to get the size of the MDIForm in a reliable fashion
    Dim FormSize As RECT
    Call WIN32.user32_GetClientRect(Me.hWnd, FormSize)
    'WIN32 returns Pixels, so scale up to Twips
    Dim FormWidth As Long, FormHeight As Long
    Let FormWidth = blu.Xpx(FormSize.Right - FormSize.Left)
    Let FormHeight = blu.Ypx(FormSize.Bottom - FormSize.Top)
    
    Let Me.toolbar.Height = 2 * blu.Ypx(blu.Metric)
    
    'Our title text should only be visible if the form is borderless
    Let Me.lblMaSS1VE.Visible = Me.bluWindow.IsBorderless
    'Position the title in the centre
    Call Me.lblMaSS1VE.Move( _
        (FormWidth - Me.lblMaSS1VE.Width) \ 2, 0, _
        Me.lblMaSS1VE.Width, blu.Ypx(blu.Metric) _
    )
    
    Let Me.cbxClose.Left = FormWidth - Me.cbxClose.Width
    Let Me.cbxMax.Left = Me.cbxClose.Left - Me.cbxMax.Width
    Let Me.cbxMin.Left = Me.cbxMax.Left - Me.cbxMin.Width
    
    'If the window is borderless, there will be min/max/close controls that _
     the version number will go next to
    Dim LeftPos As Long
    Let LeftPos = IIf(Me.bluWindow.IsBorderless = True, Me.cbxMin.Left, FormWidth)
    
    Call Me.lblVersion.Move( _
        LeftPos - Me.lblVersion.Width, 0, _
        Me.lblVersion.Width, blu.Ypx(blu.Metric) _
    )
    Call Me.btnUpdate.Move( _
        LeftPos - Me.btnUpdate.Width, 0, _
        Me.btnUpdate.Width, blu.Xpx(blu.Metric) _
    )
    
    Let Me.bluTab.Height = blu.Ypx(blu.Metric)
    Let Me.bluTab.Top = Me.toolbar.Height - Me.bluTab.Height
    Let Me.bluTab.Height = Me.toolbar.Height - Me.bluTab.Top
    
    Call Me.btnHelp.Move( _
        FormWidth - Me.btnHelp.Width, Me.toolbar.ScaleHeight - blu.Ypx(blu.Metric), _
        Me.btnHelp.Width, blu.Ypx(blu.Metric) _
    )
    
    'Help pane
    If Me.picHelp.Visible = True Then
        'Help pane toolbar
        Call Me.picHelpToolbar.Move( _
            blu.Xpx, 0, Me.picHelp.ScaleWidth - blu.Xpx, blu.Ypx(blu.Metric) _
        )
        Call Me.bluWebView.Move( _
            0, Me.picHelpToolbar.Top + Me.picHelpToolbar.Height, _
            Me.picHelp.ScaleWidth, FormHeight - Me.picHelp.Top - Me.bluWebView.Top _
        )
    End If
    
    Call lblTip.Move( _
        Me.bluTab.Left + Me.bluTab.Width, blu.Ypx(blu.Metric), _
        Me.btnHelp.Left - Me.bluTab.Left - Me.bluTab.Width, _
        blu.Ypx(blu.Metric) _
    )
End Sub

'MDIFORM Terminate _
 ======================================================================================
Private Sub MDIForm_Terminate()
    'Clean up the project in memory
    Call GAME.Clear
End Sub

'MDIFORM Unload _
 ======================================================================================
Private Sub MDIForm_Unload(Cancel As Integer)
    'We don't need to show the application whilst we clean up
    Let Me.Visible = False
    
    'Unload any possible MDI children forms
    Unload frmWelcome
    Unload frmPlay
    Unload frmLevel
End Sub

'EVENT bluTab TABCHANGED : The top tabs have been clicked - change zone _
 ======================================================================================
Private Sub bluTab_TabChanged(ByVal Index As Integer)
    'During form load the current tab is changed to -1 so that no tab is selected. _
     this is for the benefit of the welcome screen, so don't go any further
    If Index = -1 Then Exit Sub
    
    'In any instance, get rid of the welcome zone
    Unload frmWelcome
    
    Select Case Index
        Case 0 'LEVELS ----------------------------------------------------------------
            Load frmLevel
            'The level editor will set the app colours since each level has a _
             different colour scheme
            Call frmLevel.SetTheme
            Call frmLevel.Show
            Call frmLevel.SetFocus
            
            'Don't keep the PLAY tab around
            Unload frmPlay
            
        Case 1 'PLAY ------------------------------------------------------------------
            'The PLAY zone exports the project to a Master System ROM, this happens _
             automatically upon loading the form
            Load frmPlay
            'Set the colour scheme to default, this will ensure that when changing _
             from the level editor the tab colours won't mismatch
            Call Me.SetTheme
            Call frmPlay.Show
            Call frmPlay.SetFocus
            
    End Select
End Sub

'EVENT btnHelp CLICK : Hide and show the help pane _
 ======================================================================================
Private Sub btnHelp_Click()
    If Me.picHelp.Visible = False Then
        Me.btnHelp.State = bluSTATE.Active
        Me.picHelp.Visible = True
        Call MDIForm_Resize
    Else
        Me.btnHelp.State = bluSTATE.Inactive
        Me.picHelp.Visible = False
    End If
End Sub

'EVENT bluWindow BORDERLESSSTATECHANGE _
 ======================================================================================
Private Sub bluWindow_BorderlessStateChange(ByVal Enabled As Boolean)
    'When the Desktop Window Manager switches on or off (i.e. user changes Windows _
     theme between hardware accelerated and non-accerlerated -- classic -- themes) _
     force a resize to shift the UI layout around. The min/max/close buttons will _
     hide themselves automatically, but we will need to hide the custom caption
    Call MDIForm_Resize
End Sub

'EVENT bluDownload PROGRESS : A file is being downloaded _
 ======================================================================================
Private Sub bluDownload_Progress( _
    ByVal StatusCode As AsyncStatusCodeConstants, ByVal Status As String, _
    ByVal BytesDownloaded As Long, ByVal BytesTotal As Long _
)
    Debug.Print "bluDownload: " & _
        bluDownload.StatusCodeText(StatusCode) & " " & Chr(34) & Status & Chr(34) & _
        " (" & BytesDownloaded & " / " & BytesTotal & ")"
End Sub

'EVENT bluDownload COMPLETE : The updater has finished downloading something _
 ======================================================================================
Private Sub bluDownload_Complete()
    On Error GoTo Fail
    Dim INI As INIFile
    
    'We tag the control with what we're downloading so we can separate actions
    Select Case Me.bluDownload.Tag
        '"Update.ini" contains the latest version number which we can compare with
        Case Run.UpdateFile '----------------------------------------------------------
            'Open the Update.ini file that was downloaded, ...
            Set INI = New INIFile
            Let INI.FilePath = Run.AppData & Run.UpdateFile
            
            '...and retrieve the latest version number
            Dim Version As String, InfoURL As String
            Let Version = INI.GetString("Version")
            Let InfoURL = INI.GetString("InfoURL")
            
            'Update MaSS1VE.ini with the last time the update check was performed
            Let INI.FilePath = Run.AppData & Run.INI_Name
            Call INI.SetValue(CDbl(Now()), "LastUpdateCheck", "Update")
            Call INI.Save: Set INI = Nothing
            
            'Is it different from ours?
            If Trim(Version) <> Run.VersionString Then
                'There's an update! Download first the release notes...
                Let Me.bluDownload.Tag = "Update.html"
                Call Me.bluDownload.Download( _
                    InfoURL, Run.AppData & "Update.html", vbAsyncReadForceUpdate _
                )
            Else
                'Same version. Delete the Update.ini file so that it doesn't confuse _
                 MaSS1VE should the user manually update over the top
                Call VBA.Kill(Run.AppData & Run.UpdateFile)
            End If
        
        Case "Update.html" '-----------------------------------------------------------
            'Once the release notes have been downloaded, download the installer
            'First read the download URL from the Update.ini file
            Set INI = New INIFile
            Let INI.FilePath = Run.AppData & Run.UpdateFile
            
            Dim URL As String
            Let URL = INI.GetString("InstallURL")
            If URL <> vbNullString Then
                Let Me.bluDownload.Tag = "Update.exe"
                Call Me.bluDownload.Download( _
                    URL, Run.AppData & "Update.exe", vbAsyncReadForceUpdate _
                )
            Else
'                Stop
            End If
            Set INI = Nothing
        
        Case "Update.exe" '------------------------------------------------------------
            'Once the Update.exe has been downloaded, notify the user in the UI
            Let Me.lblVersion.Visible = False
            Let Me.btnUpdate.Visible = True
    End Select
Fail:
End Sub

'btnUpdate CLICK : The update button that appears once an update has been downloaded _
 ======================================================================================
Private Sub btnUpdate_Click()
    'Display the changelog/update UI
    Load frmUpdate
    Call frmUpdate.bluWebView.Navigate(Run.AppData & "Update.html")
    Call frmUpdate.Show(vbModal, mdiMain)
    'Was the "Exit & Update" button clicked?
    If Run.UpdateResponse = vbOK Then
        'Launch the installer with the path to the installation
        Call WIN32.shell32_ShellExecute( _
            0, vbNullString, Run.AppData & "Update.exe", _
            "/UPDATE /D=" & Left$(Run.Path, Len(Run.Path) - 1), _
            Run.AppData, SW_SHOWNORMAL _
        )
        Unload Me
    End If
End Sub

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'SetTip _
 ======================================================================================
Public Sub SetTip(Optional ByVal Message As String = vbNullString)
    If Message <> Me.lblTip.Caption Then Let Me.lblTip.Caption = Message
End Sub

'SetTheme : Change the colour scheme of the form controls _
 ======================================================================================
Public Sub SetTheme( _
    Optional ByVal BaseColour As OLE_COLOR = blu.BaseColour, _
    Optional ByVal TextColour As OLE_COLOR = blu.TextColour, _
    Optional ByVal ActiveColour As OLE_COLOR = blu.ActiveColour, _
    Optional ByVal InertColour As OLE_COLOR = blu.InertColour _
)
    'Deal with all blu controls automatically
    Call blu.ApplyColoursToForm( _
        Me, BaseColour, TextColour, ActiveColour, InertColour _
    )
    'Specifics for this form
    Let Me.BackColor = ActiveColour
    Let Me.picHelpToolbar.BackColor = ActiveColour
    Let Me.lblVersion.TextColour = &HD0D0D0
End Sub
