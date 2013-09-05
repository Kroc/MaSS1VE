VERSION 5.00
Begin VB.Form frmLevel 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Level"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   14655
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picStatusbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   14655
      TabIndex        =   16
      Top             =   8040
      Width           =   14655
      Begin MaSS1VE.bluControlBox cbxSizer 
         Height          =   360
         Left            =   14280
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Kind            =   3
      End
      Begin MaSS1VE.bluButton btnZoom2 
         Height          =   360
         Left            =   13320
         TabIndex        =   17
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   635
         Caption         =   "2×"
      End
      Begin MaSS1VE.bluButton btnZoomTV 
         Height          =   360
         Left            =   13800
         TabIndex        =   18
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   635
         Caption         =   "TV"
      End
      Begin MaSS1VE.bluButton btnGrid 
         Height          =   360
         Left            =   11640
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         Caption         =   "OFF"
      End
      Begin MaSS1VE.bluButton btnZoom1 
         Height          =   360
         Left            =   12840
         TabIndex        =   20
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   635
         Caption         =   "1×"
      End
      Begin MaSS1VE.bluLabel lblMemory 
         Height          =   375
         Left            =   240
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Caption         =   "1,319 bytes floor layout memory free"
      End
      Begin MaSS1VE.bluLabel lblZoom 
         Height          =   360
         Left            =   12120
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         Alignment       =   1
         Caption         =   "zoom"
      End
      Begin MaSS1VE.bluLabel lblGrid 
         Height          =   360
         Left            =   11040
         Top             =   0
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         Alignment       =   1
         Caption         =   "grid"
      End
   End
   Begin MaSS1VE.bluViewport vwpLevel 
      Height          =   7575
      Left            =   3720
      TabIndex        =   13
      Top             =   480
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   13361
   End
   Begin MaSS1VE.bluTab bluTab 
      Height          =   1200
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   2117
      AutoSize        =   -1  'True
      Border          =   0   'False
      Orientation     =   1
   End
   Begin VB.PictureBox picToolbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFAF00&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   14655
      TabIndex        =   3
      Top             =   0
      Width           =   14655
      Begin VB.ComboBox cmbLevels 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmLevel.frx":0000
         Left            =   60
         List            =   "frmLevel.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   60
         Width           =   3615
      End
      Begin VB.PictureBox picRings 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFAF00&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3840
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   0
         Width           =   735
      End
      Begin MaSS1VE.bluButton btnUndo 
         Height          =   480
         Left            =   4920
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   847
         Caption         =   "UNDO"
         Style           =   1
      End
      Begin MaSS1VE.bluButton btnRedo 
         Height          =   480
         Left            =   5760
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   847
         Caption         =   "REDO"
         Style           =   1
      End
      Begin MaSS1VE.bluButton btnShare 
         Height          =   480
         Left            =   13680
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   847
         Caption         =   "SHARE"
         Style           =   1
      End
      Begin MaSS1VE.bluButton btnCut 
         Height          =   480
         Left            =   6720
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   847
         Caption         =   "CUT"
         Style           =   1
      End
      Begin MaSS1VE.bluButton btnCopy 
         Height          =   480
         Left            =   7440
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   847
         Caption         =   "COPY"
         Style           =   1
      End
      Begin MaSS1VE.bluButton btnPaste 
         Height          =   480
         Left            =   8280
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   847
         Caption         =   "PASTE"
         Style           =   1
      End
      Begin VB.Line lineSplit 
         BorderColor     =   &H00FFAF00&
         BorderStyle     =   3  'Dot
         Index           =   2
         Visible         =   0   'False
         X1              =   6600
         X2              =   6600
         Y1              =   90
         Y2              =   360
      End
      Begin VB.Line lineSplit 
         BorderColor     =   &H00FFAF00&
         BorderStyle     =   3  'Dot
         Index           =   1
         Visible         =   0   'False
         X1              =   4800
         X2              =   4800
         Y1              =   90
         Y2              =   360
      End
      Begin VB.Line lineSplit 
         BorderColor     =   &H00FFAF00&
         BorderStyle     =   3  'Dot
         Index           =   0
         Visible         =   0   'False
         X1              =   3720
         X2              =   3720
         Y1              =   90
         Y2              =   360
      End
   End
   Begin VB.PictureBox picSidePane 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   5535
      Index           =   0
      Left            =   480
      ScaleHeight     =   5535
      ScaleWidth      =   3255
      TabIndex        =   2
      Top             =   480
      Width           =   3255
      Begin MaSS1VE.bluViewport vwpBlocks 
         Height          =   1575
         Left            =   0
         TabIndex        =   14
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2778
         Centre          =   0   'False
      End
      Begin VB.PictureBox picBlocksToolbar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFAF00&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   1200
         Left            =   0
         ScaleHeight     =   1200
         ScaleWidth      =   3255
         TabIndex        =   4
         Top             =   0
         Width           =   3255
         Begin VB.PictureBox picBlockSelectRight 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   0
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   600
            Width           =   510
         End
         Begin VB.PictureBox picBlockSelectLeft 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   0
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   510
         End
         Begin MaSS1VE.bluLabel lblLMB 
            Height          =   495
            Left            =   480
            Top             =   0
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
            Caption         =   "LMB"
            Style           =   1
         End
         Begin MaSS1VE.bluLabel lblRMB 
            Height          =   495
            Left            =   480
            Top             =   600
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
            Caption         =   "RMB"
            Style           =   1
         End
      End
   End
End
Attribute VB_Name = "frmLevel"
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
'FORM :: frmLevel

'View and edit a specific level

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'A couple of look-up tables for, perhaps, no speed gain, but it's good practice. _
 We'll populate these in `Form_Initialize`
Private x32(0 To 256) As Long               'Multiples of 32 (for blocks)
Private x8(0 To 256) As Long                'Multiples of 8 (for tiles)

'The block offset of the mouse hover (to draw a rectangle around the current block)
Private Hover As POINT

'How wide the side-pane is measured in blocks
Private Const PaneBlockWidth As Long = 7

'Caches: _
 --------------------------------------------------------------------------------------
'Cache the block mappings so we can repaint the level quickly _
 (it's far too slow to repaint tile-by-tile)
Private BlocksCache As bluImage
'An image of the water line across the width of the whole level _
 (this is quite slow to paint every time)
Private WaterLineCache As bluImage
'This is a cache object to hold the sprites once constructed out of the tilesets
Private Sprites As S1Sprites

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'Which level is attached to the editor
Private WithEvents My_Level As S1Level
Attribute My_Level.VB_VarHelpID = -1

'Which blocks are selected to the left / right mouse buttons
Private My_BlockSelectLeft As Byte
Private My_BlockSelectRight As Byte

'The zoom level
Private My_Zoom As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'FORM Initialize _
 ======================================================================================
Private Sub Form_Initialize()
    'Allow Windows to theme VB's controls
    'NOTE: This works because "CompiledInResources.res" contains a manifest file, _
     see <www.vbforums.com/showthread.php?606736-VB6-XP-Vista-Win7-Manifest-Creator>
    'We're doing this here, instead of in `Sub Main`, because other forms do not _
     contain any common controls and this causes the EXE to crash on exit. _
     See the below function comments for further details
    Call WIN32.InitCommonControls( _
        ICC_STANDARD_CLASSES Or ICC_INTERNET_CLASSES _
    )
    
    'Setup the look up tables
    Dim i As Long: For i = 0 To 256: Let x32(i) = i * 32: Let x8(i) = i * 8: Next i
End Sub

'FORM Load _
 ======================================================================================
Private Sub Form_Load()
    'Configure the tabstrip _
     (I haven't coded in proper property storage for it yet)
    With Me.bluTab
        .TabCount = 3
        .Caption(0) = "Layout"
        .Caption(1) = "Objects"
        .Caption(2) = "Theme"
        .TabCount = 1
        .CurrentTab = 0
    End With
    
    'Populate the level select combobox
    Dim i As Long
    For i = LBound(GAME.Levels) To UBound(GAME.Levels)
        'Exclude empty levels
        If Not GAME.Levels(i) Is Nothing Then
            'As we exclude the levels the combobox index does not match up with the _
             level index, so we have to store that (awkwardly) in the `ItemData` array
            Call cmbLevels.AddItem(GAME.Levels(i).Title)
            Let cmbLevels.ItemData(cmbLevels.ListCount - 1) = i
        End If
    Next i
    
    'Set the zoom to default
    Let Me.Zoom = 1
    
    'Load the first level into the editor
    Set Me.Level = GAME.Levels(0)
End Sub

'FORM Resize _
 ======================================================================================
Private Sub Form_Resize()
    'If the form is invisible or minimised then don't bother resizing
    If Me.WindowState = vbMinimized Or Me.Visible = False Then Exit Sub
    'Ensure that the MDI child form always stays maximised when changing windows
    If Me.WindowState <> vbMaximized Then Let Me.WindowState = vbMaximized: Exit Sub
    
    'Size the toolbar along the top, we need this to position everything below it
    Call Me.picToolbar.Move( _
        0, 0, Me.ScaleWidth, blu.Ypx(blu.Metric) _
    )
    
    'Statusbar: _
     ----------------------------------------------------------------------------------
    'Size the status bar along the bottom
    Call Me.picStatusbar.Move( _
        0, Me.ScaleHeight - blu.Ypx(24), _
        Me.ScaleWidth, blu.Ypx(24) _
    )
    'Put the sizing box in the corner, _
     make it square according to the statusbar height
    Call cbxSizer.Move( _
        Me.picStatusbar.ScaleWidth - Me.picStatusbar.ScaleHeight, _
        0, Me.picStatusbar.ScaleHeight, Me.picStatusbar.ScaleHeight _
    )
    
    Let Me.lblMemory.Left = Me.bluTab.Width - blu.Xpx(8)
    'Zoom levels
    Let Me.btnZoomTV.Left = Me.cbxSizer.Left - Me.btnZoomTV.Width
    Let Me.btnZoom2.Left = Me.btnZoomTV.Left - Me.btnZoom2.Width
    Let Me.btnZoom1.Left = Me.btnZoom2.Left - Me.btnZoom1.Width
    Let Me.lblZoom.Left = Me.btnZoom1.Left - Me.lblZoom.Width
    'Grid ON/OFF
    Let Me.btnGrid.Left = Me.lblZoom.Left - Me.btnGrid.Width
    Let Me.lblGrid.Left = Me.btnGrid.Left - Me.lblGrid.Width
    
    'Side Pane Area: _
     ----------------------------------------------------------------------------------
    'Move the vertical tab strip into place
    Call Me.bluTab.Move( _
        0, Me.picToolbar.Top + Me.picToolbar.Height, _
        blu.Xpx(blu.Metric) _
    )
    
    'Move the side panes into place
    Dim i As Long
    For i = 0 To Me.picSidePane.Count - 1: Call Me.picSidePane(i).Move( _
        Me.bluTab.Left + Me.bluTab.Width, _
        Me.picToolbar.Top + Me.picToolbar.Height, _
        blu.Xpx( _
            (x32(PaneBlockWidth) + PaneBlockWidth) + _
            WIN32.GetSystemMetric(SM_CXVSCROLL) _
        ), _
        Me.ScaleHeight - ( _
            Me.picToolbar.Top + Me.picToolbar.Height + Me.picStatusbar.Height _
        ) _
    ): Next i
    
    'Blocks (Layout) Side Pane: _
     ----------------------------------------------------------------------------------
    'Position the tool bar
    Call Me.picBlocksToolbar.Move( _
        0, 0, Me.picSidePane(0).Width, 1200 _
    )
    
    'The current block selections
    '(Not certain of their layout just yet; leaving this up to the VB form layout)
'        With Me.picBlockSelectLeft
'            Let .Width = blu.Xpx(34)
'            Let .Height = blu.Ypx(34)
'            Let .Top = Me.fraBlocksToolbar.Height - .Height - blu.Ypx(7)
'            Let .Left = blu.Xpx(7)
'        End With
'        With Me.picBlockSelectRight
'            Let .Width = blu.Xpx(34)
'            Let .Height = blu.Ypx(34)
'            Let .Top = Me.picBlockSelectLeft.Top
'            Let .Left = Me.picBlockSelectLeft.Left + Me.picBlockSelectLeft.Height + blu.Xpx(7)
'        End With
    
    'Position the block picker
    Call Me.vwpBlocks.Move( _
        0, Me.picBlocksToolbar.Top + Me.picBlocksToolbar.Height + blu.Ypx, _
        Me.picSidePane(0).Width, _
        Me.picSidePane(0).Height - ( _
            Me.picBlocksToolbar.Top + Me.picBlocksToolbar.Height + blu.Ypx _
        ) _
    )
    
    'Toolbar Buttons: _
     ----------------------------------------------------------------------------------
    'The level select combobox; this will eventually be a custom text box
    Call Me.cmbLevels.Move( _
        blu.Xpx(4), (Me.picToolbar.ScaleHeight - Me.cmbLevels.Height) \ 2, _
        Me.picSidePane(0).Width + Me.bluTab.Width - blu.Xpx(4) _
    )
    
    With Me.lineSplit(0)
        Let .X1 = Me.picSidePane(0).Left + Me.picSidePane(0).Width: Let .X2 = .X1
    End With
    
    'The display that shows the number of rings on the level
    Let Me.picRings.Left = Me.picSidePane(0).Left + Me.picSidePane(0).Width 'Me.lineSplit(0).X1 + blu.Xpx(8)
    
    With Me.lineSplit(1)
        Let .X1 = Me.picRings.Left + Me.picRings.Width + blu.Xpx(8): Let .X2 = .X1
    End With
    
    'undo / redo buttons (not implemented yet)
    Let Me.btnUndo.Left = Me.lineSplit(1).X1 + blu.Xpx(8)
    Let Me.btnRedo.Left = Me.btnUndo.Left + Me.btnUndo.Width
    
    'divider between undo/redo buttons and cut/copy/paste
    With Me.lineSplit(2)
        Let .X1 = Me.btnRedo.Left + Me.btnRedo.Width + blu.Xpx(8): Let .X2 = .X1
    End With
    
    'Cut / copy / paste (not implemented yet)
    Let Me.btnCut.Left = Me.lineSplit(2).X1 + blu.Xpx(8)
    Let Me.btnCopy.Left = Me.btnCut.Left + Me.btnCut.Width
    Let Me.btnPaste.Left = Me.btnCopy.Left + Me.btnCopy.Width
    
    'Share button, aligned to the right
    Let Me.btnShare.Left = Me.picToolbar.Width - Me.btnShare.Width
    
    'Level viewport: _
     ----------------------------------------------------------------------------------
    'Reposition the level viewport
    Call Me.vwpLevel.Move( _
        Me.picSidePane(0).Left + Me.picSidePane(0).Width, _
        Me.picToolbar.Top + Me.picToolbar.Height, _
        Me.ScaleWidth - (Me.picSidePane(0).Left + Me.picSidePane(0).Width), _
        Me.ScaleHeight - ( _
            Me.picToolbar.Top + Me.picToolbar.Height + Me.picStatusbar.Height _
        ) _
    )
End Sub

'FORM Terminate _
 ======================================================================================
Private Sub Form_Terminate()
    'Clear the lookup tables set up in `Form_Initialize`
    Erase x8, x32
End Sub

'FORM Unload _
 ======================================================================================
Private Sub Form_Unload(Cancel As Integer)
    'Detatch the current level from the form _
     (this will also clean up the caches)
    Set Level = Nothing
End Sub

'EVENT My_Level BLOCKMAPPINGCHANGE : Something has modified the block mappings _
 ======================================================================================
Private Sub My_Level_BlockMappingChange(ByVal BlockIndex As Byte, ByVal TileIndex As Byte, ByVal Value As Byte)
    Call CacheBlocks
    'TODO: Repaint the block selections!
    Call RepaintLevel
End Sub

'EVENT My_Level FLOORLAYOUTCHANGE _
 ======================================================================================
Private Sub My_Level_FloorLayoutChange(ByVal X As Long, ByVal Y As Long, ByVal NewIndex As Byte, ByVal OldIndex As Byte)
    'Change just that block in the cache (saves having to repaint the entire level)
    Call PaintBlock( _
        vwpLevel.hDC, x32(X), x32(Y), NewIndex, _
        (Y >= My_Level.ObjectLayout.WaterLevel) _
    )
    Call vwpLevel.Refresh
End Sub

'EVENT My_Level RINGCOUNTCHANGE : Update the ring count display _
 ======================================================================================
Private Sub My_Level_RingCountChange()
    With picRings
        Call .Cls
        Call GAME.HUD.ApplyPalette(My_Level.SpritePalette)
        Call GAME.HUD.PaintSprite(.hDC, 8 - 4, 8 - 1, 38)
        Call GAME.HUD.PaintSprite(.hDC, 16 - 4, 8 - 1, 40)
        Call WriteNumbers(.hDC, 16, 8 - 2, My_Level.Rings)
        Call .Refresh
    End With
End Sub

'EVENT My_Level WATERLEVELCHANGE : The water line has been moved _
 ======================================================================================
'The level class will trigger this event when object &H40 is set which controls _
 where the water line is in the level (i.e. Labyrinth)
Private Sub My_Level_WaterLevelChange()
    'Repaint the level
    Call RepaintLevel
End Sub

'EVENT btnShare CLICK _
 ======================================================================================
Private Sub btnShare_Click()
'    'Create a temporary image object
'    Dim Screenshot As New bluImage
'    Call Screenshot.Create24Bit( _
'        ImageWidth:=x32(Level.Width), ImageHeight:=x32(Level.Height) _
'    )
'
'    'Paint the level into the image
'    Call WIN32.gdi32_BitBlt( _
'        Screenshot.hDC, 0, 0, LevelCache.Width, LevelCache.Height, _
'        LevelCache.hDC, 0, 0, vbSrcCopy _
'    )
'    Call WIN32.gdi32_GdiTransparentBlt( _
'        Screenshot.hDC, 0, 0, LevelCache.Width, LevelCache.Height, _
'        ObjectCache.hDC, 0, 0, LevelCache.Width, LevelCache.Height, _
'        &H123456 _
'    )
'
'    'Save it to a bitmap file
'    Call Screenshot.Save(Run.Path & "Screenshot.bmp")
'
'    'Clean up
'    Set Screenshot = Nothing
End Sub

'EVENT btnZoom1/2/TV CLICK : Change the zoom level _
 ======================================================================================
Private Sub btnZoom1_Click(): Let Me.Zoom = 1: End Sub
Private Sub btnZoom2_Click(): Let Me.Zoom = 2: End Sub
Private Sub btnZoomTV_Click(): Let Me.Zoom = 3: End Sub

'EVENT cmbLevels CLICK : The drop-down level list has changed value, load the level _
 ======================================================================================
Private Sub cmbLevels_Click()
    If Me.cmbLevels.Enabled = False Then Exit Sub
    Set Me.Level = GAME.Levels( _
        cmbLevels.ItemData(cmbLevels.ListIndex) _
    )
End Sub

'EVENT vwpLevel MOUSE[DOWN/MOVE/UP] : Mouse interacting with the level _
 ======================================================================================
Private Sub vwpLevel_MouseDown(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single, ByVal ImageX As Long, ByVal ImageY As Long)
    Call HandleMouse(Button, Shift, X, Y, ImageX, ImageY)
End Sub
Private Sub vwpLevel_MouseMove(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single, ByVal ImageX As Long, ByVal ImageY As Long)
    Call HandleMouse(Button, Shift, X, Y, ImageX, ImageY)
End Sub
Private Sub vwpLevel_MouseUp(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single, ByVal ImageX As Long, ByVal ImageY As Long)
    Call HandleMouse(Button, Shift, X, Y, ImageX, ImageY)
End Sub

'EVENT vwpLevel MOUSEOUT : Mouse has gone out of the level viewport _
 ======================================================================================
Private Sub vwpLevel_MouseOut()
    Call mdiMain.SetTip
    'Clear the hover box
    Let Hover.X = -1: Let Hover.Y = -1
    
    'Refresh the viewport so the hover rectangle disappears
    Call vwpLevel.Refresh
End Sub

'EVENT vwpLevel PAINT : whenever the level viewport changes, paint on the selection _
 ======================================================================================
Private Sub vwpLevel_Paint(ByVal hDC As Long)
    'If the mouse is not hovered over any block nor is there a selection, then skip
    If Hover.X = -1 Then Exit Sub
    
    'Determine where in the viewport the selection rectangle begins
    Dim X As Long, Y As Long
    Let X = Me.vwpLevel.CentreX + (x32(Hover.X) - Me.vwpLevel.ScrollX) * My_Zoom
    Let Y = Me.vwpLevel.CentreY + (x32(Hover.Y) - Me.vwpLevel.ScrollY) * My_Zoom
    
    Dim Box As RECT
    Call WIN32.user32_SetRect(Box, X, Y, X + x32(My_Zoom) + 3, Y + x32(My_Zoom) + 3)
    Call WIN32.user32_FrameRect(hDC, Box, WIN32.gdi32_GetStockObject(BLACK_BRUSH))
    Call WIN32.user32_SetRect(Box, X - 1, Y - 1, X + x32(My_Zoom) + 1, Y + x32(My_Zoom) + 1)
    Call WIN32.user32_FrameRect(hDC, Box, WIN32.gdi32_GetStockObject(WHITE_BRUSH))
    Call WIN32.user32_SetRect(Box, X - 2, Y - 2, X + x32(My_Zoom) + 2, Y + x32(My_Zoom) + 2)
    Call WIN32.user32_FrameRect(hDC, Box, WIN32.gdi32_GetStockObject(WHITE_BRUSH))
End Sub

'EVENT vwpBlocks MOUSEUP _
 ======================================================================================
Private Sub vwpBlocks_MouseUp(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single, ByVal ImageX As Long, ByVal ImageY As Long)
    'Which block was clicked?
    Dim Index As Long
    Let Index = _
        ((ImageY \ 33) * PaneBlockWidth) _
        + (ImageX \ 33)
    
    'Only allow clicks within the block range
    If Index >= 0 Or Index = (Level.BlockMapping.Length - 1) Then
        'Which mouse button was clicked?
        If Button = VBRUN.MouseButtonConstants.vbLeftButton Then
            Let Me.BlockSelectLeft = Index
        ElseIf Button = VBRUN.MouseButtonConstants.vbRightButton Then
            Let Me.BlockSelectRight = Index
        End If
    End If
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////

'PROPERTY BlockSelectLeft : Which block is currently set to the left mouse button _
 ======================================================================================
Public Property Get BlockSelectLeft() As Byte: Let BlockSelectLeft = My_BlockSelectLeft: End Property
Public Property Let BlockSelectLeft(ByVal Index As Byte)
    Let My_BlockSelectLeft = Index
    'Paint the selected block into the selection box
    If Not Level Is Nothing Then _
        Call PaintBlock(Me.picBlockSelectLeft.hDC, 1, 1, Index): _
        Me.picBlockSelectLeft.Refresh
End Property

'PROPERTY BlockSelectLeft : Which block is currently set to the right mouse button _
 ======================================================================================
Public Property Get BlockSelectRight() As Byte: Let BlockSelectRight = My_BlockSelectRight: End Property
Public Property Let BlockSelectRight(ByVal Index As Byte)
    Let My_BlockSelectRight = Index
    'Paint the selected block into the selection box
    If Not Level Is Nothing Then _
        Call PaintBlock(Me.picBlockSelectRight.hDC, 1, 1, Index): _
        Me.picBlockSelectRight.Refresh
End Property

'PROPERTY Level : Attach a level object to the editor _
 ======================================================================================
Public Property Get Level() As S1Level: Set Level = My_Level: End Property
Public Property Set Level(ByVal TheLevel As S1Level)
    'Keep a hold of the level
    Set My_Level = TheLevel
    
    'Clean up
    Set BlocksCache = Nothing
    Set WaterLineCache = Nothing
    Set Sprites = Nothing
    
    'Is the selected level valid?
    If Not My_Level Is Nothing Then
        'When the level form loads, the level select combobox text is blank, we need _
         to set it but doing so would trigger the `Click` event, causing an infinite _
         loop of trying to load the level
        If Me.cmbLevels.Text = "" Then
            Let Me.cmbLevels.Enabled = False
            Let Me.cmbLevels.Text = My_Level.Title
            Let Me.cmbLevels.Enabled = True
        End If
        
        'Cache the block mappings for speedy painting of the level
        Call CacheBlocks
        
        'Reset the select block indexes _
         (these will be repainted; dependent on the above)
        Let Me.BlockSelectLeft = 1
        Let Me.BlockSelectRight = 0
        
        'Cache an image of the water line across the whole level _
         (this is significantly faster than painting the line every repaint)
        Set WaterLineCache = New bluImage
        Call WaterLineCache.Create8Bit( _
            x32(My_Level.Width), 8, _
            My_Level.SpritePalette.Colours, True _
        )
        Dim WaterLineCache_hDC As Long
        Let WaterLineCache_hDC = WaterLineCache.hDC
        
        Dim MyLevel_SpriteArt_Tiles_hDC As Long
        Let MyLevel_SpriteArt_Tiles_hDC = My_Level.SpriteArt.Tiles.hDC
        Dim MyLevel_SpritePalette_Colour0 As Long
        Let MyLevel_SpritePalette_Colour0 = My_Level.SpritePalette.Colour(0)
        
        Dim X As Long
        For X = 0 To My_Level.Width - 1
            Call WIN32.gdi32_GdiTransparentBlt(WaterLineCache.hDC, x32(X), 0, 8, 8, MyLevel_SpriteArt_Tiles_hDC, 0, 0, 8, 8, MyLevel_SpritePalette_Colour0)
            Call WIN32.gdi32_GdiTransparentBlt(WaterLineCache.hDC, x32(X) + 8, 0, 8, 8, MyLevel_SpriteArt_Tiles_hDC, 16, 0, 8, 8, MyLevel_SpritePalette_Colour0)
            Call WIN32.gdi32_GdiTransparentBlt(WaterLineCache.hDC, x32(X) + 16, 0, 8, 8, MyLevel_SpriteArt_Tiles_hDC, 0, 0, 8, 8, MyLevel_SpritePalette_Colour0)
            Call WIN32.gdi32_GdiTransparentBlt(WaterLineCache.hDC, x32(X) + 24, 0, 8, 8, MyLevel_SpriteArt_Tiles_hDC, 16, 0, 8, 8, MyLevel_SpritePalette_Colour0)
        Next
        
        'Update the ring count display
        Call My_Level_RingCountChange
        
        'Reset the cached sprite images, they will repaint on-demand
        Set Sprites = New S1Sprites
        Set Sprites.SpriteArt = My_Level.SpriteArt
        Set Sprites.Palette = My_Level.SpritePalette
        
        'Level viewport: _
         ------------------------------------------------------------------------------
        'Resize the viewport's image buffer to the size of the level
        Call Me.vwpLevel.SetImageProperties( _
            x32(My_Level.Width), x32(My_Level.Height) _
        )
        'Add a layer to the viewport for the level objects
        Call Me.vwpLevel.AddLayer
        'Now paint the level in the viewport
        Call RepaintLevel
        
        'Centre the viewport on Sonic
        Call Me.vwpLevel.ScrollTo( _
            x32(My_Level.StartX + 1) - (((vwpLevel.Width \ Screen.TwipsPerPixelX) \ My_Zoom) \ 2), _
            x32(My_Level.StartY + 1) - (((vwpLevel.Height \ Screen.TwipsPerPixelY) \ My_Zoom) \ 2) _
        )
        
        'Block mappings side pane: _
         ------------------------------------------------------------------------------
        'Create the blocks image in the viewport
        Let vwpBlocks.BackColor = blu.BaseColour
        'Determine the height of the block list based on _
         a. the number of blocks in the mapping _
         b. the number of blocks in the width of the side pane
        'This is trickier than it sounds because of rounding problems where you need _
         one extra row for only one or two blocks (which would normally round down)
        Call vwpBlocks.SetImageProperties( _
            Width:=x32(PaneBlockWidth) + PaneBlockWidth - 1, _
            Height:=x32(Lib.RoundUp(Level.BlockMapping.Length / PaneBlockWidth)) + _
                    Lib.RoundUp(Level.BlockMapping.Length / PaneBlockWidth) _
        )
        Dim i As Byte
        For i = 0 To My_Level.BlockMapping.Length - 1
            Call PaintBlock( _
                hndDeviceContext:=vwpBlocks.hDC, _
                X:=1 + (i Mod PaneBlockWidth) * 33, _
                Y:=(i \ PaneBlockWidth) * 33, _
                Index:=i _
            )
        Next i
        Call vwpBlocks.Refresh
        
        'Set the app colour scheme
        Call Me.SetTheme
    End If
End Property

'PROPERTY Zoom _
 ======================================================================================
Public Property Get Zoom() As Long: Let Zoom = My_Zoom: End Property
Public Property Let Zoom(ByVal ZoomLevel As Long)
    Let My_Zoom = ZoomLevel
    
    Let Me.vwpLevel.Zoom = My_Zoom
    
    Me.btnZoom1.State = IIf(My_Zoom = 1, bluSTATE.Active, bluSTATE.Inactive)
    Me.btnZoom2.State = IIf(My_Zoom = 2, bluSTATE.Active, bluSTATE.Inactive)
    Me.btnZoomTV.State = IIf(My_Zoom = 3, bluSTATE.Active, bluSTATE.Inactive)
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'SetTheme : Change the colour scheme of the form controls _
 ======================================================================================
Public Sub SetTheme()
    
    Dim ActiveColour As Long, InertColour As Long
    Dim HSLColour As HSL
    
    'If no level is attached yet, we can't set the colour scheme based on the level
    If My_Level Is Nothing Then
        'Default to a grey colour
        Let ActiveColour = VBRUN.SystemColorConstants.vbApplicationWorkspace
    Else
        'For now this is hardcoded until such a time we allow changing of level themes
        Select Case cmbLevels.ItemData(cmbLevels.ListIndex)
            'Green Hill Zone (+End Sequence), Bridge
            Case 0 To 5, 18: Let ActiveColour = blu.ActiveColour
            'Jungle
            Case 6 To 8: Let ActiveColour = &H5000&
            'Labyrinth
            Case 9 To 10: Let ActiveColour = &HAFFF&
            'Sky Base Act 1
            Case 11: Let ActiveColour = &H50AF00
            'Sky Base Act 2 including Emerald Maze / Ballhog Area
            Case 12 To 14, 20 To 25: Let ActiveColour = &HAFAF50
            Case 15 To 16: Let ActiveColour = &H500000
            'Sky Base Act 2 / 3 Interior
            Case 17, 26 To 27: Let ActiveColour = &HFFAF50
            'Special stages
            Case 28 To 35: Let ActiveColour = &H5000FF
        End Select
    End If
    
    'Calculate the inert text colour from the main active colour
    Let HSLColour = Lib.RGBToHSL(ActiveColour)
    'Use light text on dark background and dark text on light background
    If HSLColour.Luminance < 100 Then Let HSLColour.Luminance = 85
    'Duller colour
    Let HSLColour.Saturation = 27
    
    Let InertColour = Lib.HSLToRGB( _
        HSLColour.Hue, HSLColour.Saturation, HSLColour.Luminance _
    )
    
    Call mdiMain.SetTheme(, , ActiveColour, InertColour)
    
    'Deal with all blu controls automatically
    Call blu.ApplyColoursToForm( _
        Me, blu.BaseColour, blu.TextColour, ActiveColour, InertColour _
    )
    
    'Some frmLevel specifics
    Let Me.picToolbar.BackColor = ActiveColour
    Let Me.picBlocksToolbar.BackColor = ActiveColour
    
    Dim i As Long
    For i = 0 To Me.lineSplit.Count - 1
        Let Me.lineSplit(i).BorderColor = ActiveColour
        'For some reason, this is necessary otherwise the line can randomly disappear!
        Call Me.lineSplit(i).Refresh
    Next i
    
    'Update the ring counter
    Let Me.picRings.BackColor = ActiveColour
    If Not My_Level Is Nothing Then Call My_Level_RingCountChange
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'CacheBlocks : Save time by caching the 32x32 block mappings _
 ======================================================================================
'The Floor Layout is made up of blocks, each block is a 4x4 arrangement of tiles from _
 the level art. It is expensive to paint the whole level tile-by-tile so we cache the _
 block mappings and paint 32x32px at a time using those
Private Sub CacheBlocks()
    Set BlocksCache = Nothing
    Set BlocksCache = New bluImage
    Call BlocksCache.Create8Bit( _
        ImageWidth:=x32(My_Level.BlockMapping.Length), ImageHeight:=32, _
        Palette_LongArray:=My_Level.LevelPalette.Colours _
    )
    
    'Any object we have to reference in a loop is slow
    Dim BlocksCache_hDC As Long
    Let BlocksCache_hDC = BlocksCache.hDC
    
    Dim i As Long, iX As Long, iY As Long
    For i = 0 To My_Level.BlockMapping.Length - 1
        For iY = 0 To 3: For iX = 0 To 3
            Call My_Level.LevelArt.PaintTile( _
                hDC:=BlocksCache_hDC, X:=x32(i) + x8(iX), Y:=x8(iY), _
                Index:=My_Level.BlockMapping.Tile( _
                    BlockIndex:=i, TileIndex:=(iY * 4) + iX _
                ) _
            )
        Next iX: Next iY
    Next i
End Sub

'HandleMouse : The MouseDown / Move / Up events are handled the same _
 ======================================================================================
Private Sub HandleMouse( _
    ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single, _
    ByVal ImageX As Long, ByVal ImageY As Long _
)
    Dim Block As POINT
    
    'If the level is thinner/shorter than the viewport it'll be centred; if the mouse _
     is outside of where the level is centred, there's no point going any further, _
     skip ahead to "mouse out" now
    If ImageX < 0 Then GoTo Invalid
    If ImageY < 0 Then GoTo Invalid
    If ImageX >= Me.vwpLevel.ImageWidth Then GoTo Invalid
    If ImageY >= Me.vwpLevel.ImageHeight Then GoTo Invalid
    
    'Calculate which block the mouse is over
    Let Block.X = ImageX \ 32
    Let Block.Y = ImageY \ 32
    'Has the hover rectangle moved from one block to the next? _
     (we want to ensure we repaint only when the mouse moves from one block to another _
      and not with every single mouse move event)
    Dim DoRefresh As Boolean
    If Block.X <> Hover.X Or Block.Y <> Hover.Y Then Let DoRefresh = True
    'Remember which block the mouse is over so that whenever the viewport repaints, _
     we can draw the hover rectangle
    Let Hover.X = Block.X: Let Hover.Y = Block.Y
    
    'Just hovering, thanks
    If Button = 0 Then
        Call mdiMain.SetTip( _
            "L/R-CLICK: Set blocks M-CLICK: Pick block" _
        )
        
    'Left Click:
    ElseIf Button = VBRUN.MouseButtonConstants.vbLeftButton Then
        'Change the block to the one set to the left mouse button. _
         The `My_Level_FloorLayoutChange` event will paint the changed block _
         and refresh the viewport for us (we don't need to repaint the whole level)
        Let Level.FloorLayout.Block(Block.X, Block.Y) = BlockSelectLeft
        'We can leave here; if the selection rectangle moved it will already have _
         been redrawn with the above code causing a repaint
        Exit Sub
        
    'Right Click:
    ElseIf Button = VBRUN.MouseButtonConstants.vbRightButton Then
        'As above, but with the right mouse button
        Let Level.FloorLayout.Block(Block.X, Block.Y) = BlockSelectRight
        Exit Sub
        
    'Middle Click:
    ElseIf Button = VBRUN.MouseButtonConstants.vbMiddleButton Then
        'Pick the current block under the mouse cursor
        Let BlockSelectLeft = Level.FloorLayout.Block(Block.X, Block.Y)
        
    End If
    
    'Is the mouse in a different block than it was last time?
    If DoRefresh = True Then Call Me.vwpLevel.Refresh
    
    Exit Sub
    
Invalid:
    'The mouse is not over any block; _
     if this is a change to previously, refresh
    If Hover.X <> -1 Then
        'No hover rectangle
        Let Hover.X = -1: Let Hover.Y = -1
        Call Me.vwpLevel.Refresh
    End If
    'Clear the help tip when in the viewport, but not on a map block
    Call mdiMain.SetTip
End Sub

'PaintBlock _
 ======================================================================================
Private Sub PaintBlock( _
    ByVal hndDeviceContext As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal Index As Byte, _
    Optional ByVal UnderWater As Boolean = False _
)
    'We don't want to have to set the palette with every call to this function, so _
     we remember what was used last time
    Static IsAlreadyUnderWater As Boolean
    
    'Do we need to change the palette?
    If UnderWater <> IsAlreadyUnderWater Then
        If UnderWater = False Then
            Call My_Level.LevelPalette.ApplyToImage(BlocksCache)
        Else
            Call GAME.UnderwaterLevelPalette.ApplyToImage(BlocksCache)
        End If
    End If
    Let IsAlreadyUnderWater = UnderWater
    
    'Paint the block from the cache, which is a whole ton faster than painting _
     every tile in the block
    
    Call WIN32.gdi32_BitBlt( _
        hndDeviceContext, X, Y, 32, 32, _
        BlocksCache.hDC, x32(Index), 0, vbSrcCopy _
    )
End Sub

'PaintObject _
 ======================================================================================
Private Sub PaintObject(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Index As OBJECT_TYPE)
    Select Case Index
        'POWER UPS
        Case OBJECT_TYPE.Monitor_Ring
            Call Sprites.Monitor_Ring.Paint(hDC, X + 4, Y - 7)
        Case OBJECT_TYPE.Monitor_Speed
            Call Sprites.Monitor_Speed.Paint(hDC, X + 4, Y - 7)
        Case OBJECT_TYPE.Monitor_Life
            Call Sprites.Monitor_Life.Paint(hDC, X + 4, Y - 7)
        Case OBJECT_TYPE.Monitor_Shield
            Call Sprites.Monitor_Shield.Paint(hDC, X + 4, Y - 7)
        Case OBJECT_TYPE.Monitor_Stars
            Call Sprites.Monitor_Stars.Paint(hDC, X + 4, Y - 7)
        Case OBJECT_TYPE.Monitor_Check
            Call Sprites.Monitor_Check.Paint(hDC, X + 4, Y - 7)
        Case OBJECT_TYPE.Monitor_Cont
            Call Sprites.Monitor_Cont.Paint(hDC, X + 4, Y - 7)
        Case OBJECT_TYPE.Emerald
            Call Sprites.Emerald.Paint(hDC, X + 8, Y)
            
        Case OBJECT_TYPE.END_SIGN
            Call Sprites.EndSign.Paint(hDC, X, Y + 9)
            
        Case OBJECT_TYPE.BAD_MOTO
            Call Sprites.Badnick_Motobug.Paint(hDC, X, Y)
        Case OBJECT_TYPE.BAD_CRAB
            Call Sprites.Badnick_Crabmeat.Paint(hDC, X, Y - 8)
        Case OBJECT_TYPE.BAD_BUZZ
            Call Sprites.Badnick_BuzzBomber.Paint(hDC, X, Y)
        Case OBJECT_TYPE.BAD_NEWT
            Call Sprites.Badnick_Newtron.Paint(hDC, X + 8, Y)
        Case OBJECT_TYPE.BAD_CHOP
            Call Sprites.Badnick_Chopper.Paint(hDC, X + 8, Y + 8)
    End Select
End Sub

'RepaintLevel : Redraw the whole level and refresh the viewport _
 ======================================================================================
Private Sub RepaintLevel()
    'Try avoid painting if not necessary
    If My_Level Is Nothing Then Exit Sub
    
    Dim i As Long
    Dim iX As Long, iY As Long
    
    'Temporarily cache the water level Y index so that we don't need to do this _
     lookup within a loop (can be very slow)
    Dim MyLevel_ObjectLayout_WaterLevel As Long
    Let MyLevel_ObjectLayout_WaterLevel = My_Level.ObjectLayout.WaterLevel
    
    'Floor Layout: _
     ----------------------------------------------------------------------------------
    Dim StartTimeL As Single
    Let StartTimeL = Timer
    
    Dim vwpLevel_hDC0 As Long
    Let vwpLevel_hDC0 = vwpLevel.hDC
    
    'Cache the complete floor layout to save having to reference through objects _
     every time. This makes a big speed difference
    Dim MyLevel_FloorLayout() As Byte
    Let MyLevel_FloorLayout = My_Level.FloorLayout.GetByteStream()
    
    Dim MyLevel_Width As Long, MyLevel_Height As Long
    Let MyLevel_Width = My_Level.Width - 1
    Let MyLevel_Height = My_Level.Height - 1
    
    'Since the raw floor layout data is a 1-dimensional array we need to convert _
     X,Y to a single index number. We create a small lookup table to help having _
     to calculate the index in the loop
    Dim xY() As Long
    ReDim xY(0 To MyLevel_Height) As Long
    For i = 0 To MyLevel_Height: Let xY(i) = i * (MyLevel_Width + 1): Next i
    
    'Paint the entire floor layout in one go
    For iY = 0 To MyLevel_Height: For iX = 0 To MyLevel_Width
        Call PaintBlock( _
            hndDeviceContext:=vwpLevel_hDC0, X:=x32(iX), Y:=x32(iY), _
            Index:=MyLevel_FloorLayout(xY(iY) + iX), _
            UnderWater:=(iY >= MyLevel_ObjectLayout_WaterLevel) _
        )
    Next iX: Next iY
    
    'Clean up the temporary copy of the floor layout
    Erase MyLevel_FloorLayout
    
    Debug.Print "Repaint Floor Layout - " & Round(Timer - StartTimeL, 4)
    
    'Object Layout: _
     ----------------------------------------------------------------------------------
    Dim StartTimeO As Single
    Let StartTimeO = Timer
    
    Dim vwpLevel_hDC1 As Long
    Let vwpLevel_hDC1 = vwpLevel.hDC(1)
    
    'Loop over all the objects in the level
    For i = 0 To 255
        With My_Level.ObjectLayout.Object(i)
            'If there's an object in this slot
            If .O > 0 Then
                'Above or below the water line?
                If .Y < MyLevel_ObjectLayout_WaterLevel Then
                    'Above, use the level's given palette
                    If Not Sprites.Palette Is My_Level.SpritePalette _
                        Then Set Sprites.Palette = My_Level.SpritePalette
                Else
                    'Below, use the specific underwater palette
                    If Not Sprites.Palette Is GAME.UnderwaterSpritePalette _
                        Then Set Sprites.Palette = GAME.UnderwaterSpritePalette
                End If
                'Now paint the object
                Call PaintObject( _
                    hDC:=vwpLevel_hDC1, X:=x32(.X), Y:=x32(.Y), _
                    Index:=.O _
                )
            End If
        End With
    Next i
    
    'Paint Sonic at the starting point. This will also require checking if above _
     or below the water line
    If My_Level.StartY < MyLevel_ObjectLayout_WaterLevel Then
        'Above, use the level's given palette
        If Not Sprites.Palette Is My_Level.SpritePalette _
            Then Set Sprites.Palette = My_Level.SpritePalette
    Else
        'Below, use the specific underwater palette
        If Not Sprites.Palette Is GAME.UnderwaterSpritePalette _
            Then Set Sprites.Palette = GAME.UnderwaterSpritePalette
    End If
    Call Sprites.Sonic.Paint( _
        hndDeviceContext:=vwpLevel_hDC1, _
        DestLeft:=x32(My_Level.StartX), _
        DestTop:=x32(My_Level.StartY) + 16, _
        DestWidth:=x8(3), DestHeight:=x8(4) _
    )
    
    If My_Level.IsUnderWater = True Then
        Call WaterLineCache.Paint( _
            vwpLevel_hDC1, 0, x32(MyLevel_ObjectLayout_WaterLevel) - 4 _
        )
    End If
    
    Debug.Print "Repaint Object Layout - " & Round(Timer - StartTimeO, 4)
    
    Call Me.vwpLevel.Refresh
End Sub

'WriteNumbers : Paint a number using the HUD graphics _
 ======================================================================================
Public Sub WriteNumbers(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal NumberString As String)
    Call GAME.HUD.ApplyPalette(My_Level.SpritePalette)
    
    Dim i As Long, Letter As String * 1, Index As Byte
    For i = 1 To Len(NumberString)
        Let Letter = Mid(NumberString, i, 1)
        Select Case Letter
            Case "x": Let Index = 36
            Case ":": Let Index = 48
            Case Else
                Let Index = CByte(Mid(NumberString, i, 1)) * 2
        End Select
        Call GAME.HUD.PaintSprite(hDC, X + x8(i), Y, Index)
    Next i
End Sub
