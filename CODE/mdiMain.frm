VERSION 5.00
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFAF00&
   Caption         =   "MaSS1VE"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picHelp 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7425
      Left            =   11460
      ScaleHeight     =   7425
      ScaleWidth      =   3660
      TabIndex        =   1
      Top             =   990
      Visible         =   0   'False
      Width           =   3660
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
      Begin MaSS1VE.bluControlBox cbxClose 
         Height          =   480
         Left            =   14640
         TabIndex        =   5
         Top             =   0
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin MaSS1VE.bluTab bluTab 
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   495
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
      End
      Begin MaSS1VE.bluButton btnHelp 
         Height          =   495
         Left            =   14160
         TabIndex        =   3
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
         TabIndex        =   6
         Top             =   0
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         Kind            =   1
      End
      Begin MaSS1VE.bluControlBox cbxMax 
         Height          =   480
         Left            =   14280
         TabIndex        =   7
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
         Caption         =   "The quick brown fox jumps over the lazy dog"
         Enabled         =   0   'False
      End
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
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'FORM :: mdiMain

'The current selected level in the editor; _
 this is just temporary until we've completed the level selector (frmLevels)
Private LevelIndex As Byte

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'MDIFORM Load _
 ======================================================================================
Private Sub MDIForm_Load()
    Call Me.SetTheme
    
    Call SetTip
    
    With Me.bluTab
        .Border = False
        .AutoSize = True
        .TabCount = 2
        .Caption(0) = "LEVELS"
        .Caption(1) = "PLAY"
        'Select no tab to begin with, the welcome screen will be shown by default
        .CurrentTab = -1
    End With
    
    Call Me.bluWindow.RegisterMoveHandler(Me.toolbar)
    
    'If on a small screen, start up maximised (we need at least 1024x600)
    If Screen.Width \ Screen.TwipsPerPixelX <= 1024 Then
        Let mdiMain.WindowState = VBRUN.FormWindowStateConstants.vbMaximized
    End If
    
    Load frmWelcome
    Call frmWelcome.Show
End Sub

'MDIFORM Reisze _
 ======================================================================================
Private Sub MDIForm_Resize()
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
    
    Call Me.lblMaSS1VE.Move( _
        (FormWidth - Me.lblMaSS1VE.Width) \ 2, 0, _
        Me.lblMaSS1VE.Width, blu.Ypx(blu.Metric) _
    )
    
    Let Me.cbxClose.Left = FormWidth - Me.cbxClose.Width
    Let Me.cbxMax.Left = Me.cbxClose.Left - Me.cbxMax.Width
    Let Me.cbxMin.Left = Me.cbxMax.Left - Me.cbxMin.Width
    
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
    End If
    
    Call lblTip.Move( _
        Me.bluTab.Left + Me.bluTab.Width, blu.Ypx(blu.Metric), _
        Me.btnHelp.Left - Me.bluTab.Left - Me.bluTab.Width, _
        blu.Ypx(blu.Metric) _
    )
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
            Let frmLevel.WindowState = VBRUN.FormWindowStateConstants.vbMaximized
            Call frmLevel.Show
            
            'Don't keep the PLAY tab around
            Unload frmPlay
            
        Case 1 'PLAY ------------------------------------------------------------------
            'The PLAY zone exports the project to a Master System ROM, this happens _
             automatically upon loading the form
            Load frmPlay
            'Set the colour scheme to default, this will ensure that when changing _
             from the level editor the tab colours won't mismatch
            Call Me.SetTheme
            Let frmPlay.WindowState = VBRUN.FormWindowStateConstants.vbMaximized
            Call frmPlay.Show
            
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

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'SetTip _
 ======================================================================================
Public Sub SetTip(Optional ByVal Message As String = "")
    If Message <> Me.lblTip.Caption Then Let Me.lblTip.Caption = Message
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

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
End Sub
