VERSION 5.00
Begin VB.Form frmEditor 
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
   Begin VB.PictureBox picSizer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   14400
      Picture         =   "frmEditor.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   8160
      Width           =   255
   End
   Begin MaSS1VE.bluViewport vwpLevel 
      Height          =   7935
      Left            =   3720
      TabIndex        =   19
      Top             =   480
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   13996
   End
   Begin MaSS1VE.bluTab bluTab 
      Height          =   1200
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   495
      _extentx        =   873
      _extenty        =   2117
      border          =   0   'False
      orientation     =   1
      autosize        =   -1  'True
   End
   Begin VB.Frame fraToolbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFAF00&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14655
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
      Begin MaSS1VE.bluLabel lblTitle 
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         Caption         =   "Green Hill Act 1"
         State           =   1
         Style           =   1
      End
      Begin MaSS1VE.bluButton btnZoom2 
         Height          =   480
         Left            =   12480
         TabIndex        =   4
         Top             =   0
         Width           =   495
         _extentx        =   873
         _extenty        =   847
         caption         =   "2×"
         style           =   1
      End
      Begin MaSS1VE.bluButton btnZoomTV 
         Height          =   480
         Left            =   12960
         TabIndex        =   5
         Top             =   0
         Width           =   495
         _extentx        =   873
         _extenty        =   847
         caption         =   "TV"
         style           =   1
      End
      Begin MaSS1VE.bluLabel lblGrid 
         Height          =   480
         Left            =   10320
         Top             =   0
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   847
         Alignment       =   1
         Caption         =   "grid"
         Style           =   1
      End
      Begin MaSS1VE.bluButton btnGrid 
         Height          =   480
         Left            =   10800
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   615
         _extentx        =   1085
         _extenty        =   847
         caption         =   "OFF"
         style           =   1
      End
      Begin MaSS1VE.bluLabel lblZoom 
         Height          =   480
         Left            =   11280
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   847
         Alignment       =   1
         Caption         =   "zoom"
         Style           =   1
      End
      Begin MaSS1VE.bluButton btnZoom1 
         Height          =   480
         Left            =   12000
         TabIndex        =   7
         Top             =   0
         Width           =   495
         _extentx        =   873
         _extenty        =   847
         caption         =   "1×"
         style           =   1
      End
      Begin MaSS1VE.bluButton btnUndo 
         Height          =   480
         Left            =   4920
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   735
         _extentx        =   1296
         _extenty        =   847
         caption         =   "UNDO"
         style           =   1
      End
      Begin MaSS1VE.bluButton btnRedo 
         Height          =   480
         Left            =   5760
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   735
         _extentx        =   1296
         _extenty        =   847
         caption         =   "REDO"
         style           =   1
      End
      Begin MaSS1VE.bluButton btnShare 
         Height          =   480
         Left            =   13680
         TabIndex        =   15
         Top             =   0
         Width           =   975
         _extentx        =   1720
         _extenty        =   847
         caption         =   "SHARE"
         style           =   1
      End
      Begin MaSS1VE.bluButton btnCut 
         Height          =   480
         Left            =   6720
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   615
         _extentx        =   1085
         _extenty        =   847
         caption         =   "CUT"
         style           =   1
      End
      Begin MaSS1VE.bluButton btnCopy 
         Height          =   480
         Left            =   7440
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   735
         _extentx        =   1296
         _extenty        =   847
         caption         =   "COPY"
         style           =   1
      End
      Begin MaSS1VE.bluButton btnPaste 
         Height          =   480
         Left            =   8280
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         _extentx        =   1508
         _extenty        =   847
         caption         =   "PASTE"
         style           =   1
      End
      Begin VB.Line lineSplit 
         BorderColor     =   &H00FFAF00&
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   13560
         X2              =   13560
         Y1              =   90
         Y2              =   360
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
   Begin VB.Frame fraSidePane 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   3255
      Begin MaSS1VE.bluViewport vwpBlocks 
         Height          =   1575
         Left            =   0
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2778
      End
      Begin VB.Frame fraBlocksToolbar 
         BackColor       =   &H00FFAF00&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1200
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3255
         Begin VB.CommandButton Command3 
            Caption         =   "ROM!"
            Height          =   255
            Left            =   2400
            TabIndex        =   12
            Top             =   0
            Width           =   735
         End
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
            TabIndex        =   10
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
            TabIndex        =   9
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
Attribute VB_Name = "frmEditor"
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
'FORM :: frmEditor

'View and edit a specific level

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'A couple of look-up tables for, perhaps, no speed gain, but it's good practice. _
 We'll populate these in `Form_Initialize`
Dim x32(0 To 256) As Long               'Multiples of 32 (for blocks)
Dim x8(0 To 256) As Long                'Multiples of 8 (for tiles)

'To speed things up we cache lots of calculations that don't change every frame
Private Type CACHEVARS
    'The size of a block, 32 normally, but multiplied for zooming
    BlockSize As Long
    'When the viewport is larger than the level, we centre the level in the viewport
    Centre As POINT
    'When centred, the portion of level to paint is smaller than the viewport, _
     otherwise these will be the size of the viewport
    Dest As SIZE
    'The portion of the level to display in the viewport. At zoom level 1, this is _
     the same as the DestWidth/Height, but zoomed in, it will be smaller
    Src As SIZE
    'The offset to the block in the upper left hand corner, that is, how far scrolled _
     we are measured in blocks
    Block As POINT
    'Since we can scroll per-pixel, this is a cache of how many pixels offset we are _
     from the nearest block (above)
    BlockPxOffset As POINT
    'The block offset of the mouse hover (to draw a rectangle around the current block)
    Hover As POINT
    
    'The water line position, in pixels, in the object layout cache image _
     (that is, where to paint the water line waves across the whole level)
    WaterLevelPx As Long
    
    'The height in pixels of the block mappings image in the side pane, _
     this is used so we know how to set the scrollbar
    BlockListHeight As Long
End Type
Dim c As CACHEVARS

'How wide the side-pane is measured in blocks
Private Const PaneBlockWidth As Long = 7

'Caches: _
 --------------------------------------------------------------------------------------
'Cache the block mappings so we can repaint the level quickly _
 (it's far too slow to repaint tile-by-tile)
Dim BlocksCache As bluImage
'An image of the water line across the width of the whole level _
 (this is quite slow to paint every time)
Dim WaterLineCache As bluImage
'This is a cache object to hold the sprites once constructed out of the tilesets
Dim Sprites As S1Sprites

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
    'Setup the look up tables
    Dim i As Long
    For i = 0 To 256: Let x32(i) = i * 32: Let x8(i) = i * 8: Next i
End Sub

'FORM Load _
 ======================================================================================
Private Sub Form_Load()
    Call mdiMain.bluWindow.RegisterSizeHandler(Me.picSizer)
    
    With Me.bluTab
        .Style = bluSTYLE.Normal
        .TabCount = 3
        .Caption(0) = "Layout"
        .Caption(1) = "Objects"
        .Caption(2) = "Properties"
        .CurrentTab = 0
    End With
    
    'Set the zoom to default
    Let Me.Zoom = 1
    
    Call frmEditor.Show
    
    'Load the first level into the editor
    Set Me.Level = GAME.Levels(0)
End Sub

'FORM Resize _
 ======================================================================================
Private Sub Form_Resize()
    'If the form is minimised then don't bother resizing
    If Me.WindowState = vbMinimized Or Me.Visible = False Then Exit Sub
    
    'Size the toolbar along the top, we need this to position everything below it
    Call Me.fraToolbar.Move( _
        0, 0, Me.ScaleWidth, blu.Ypx(blu.Metric) _
    )
    
    'Side Pane Area: _
     ----------------------------------------------------------------------------------
    'Move the vertical tab strip into place
    Call Me.bluTab.Move( _
        0, Me.fraToolbar.Top + Me.fraToolbar.Height, _
        blu.Xpx(blu.Metric) _
    )
    
    'Move the side panes into place
    Dim i As Long
    For i = 0 To Me.fraSidePane.Count - 1: Call Me.fraSidePane(i).Move( _
        Me.bluTab.Left + Me.bluTab.Width, _
        Me.fraToolbar.Top + Me.fraToolbar.Height, _
        blu.Xpx( _
            (x32(PaneBlockWidth) + PaneBlockWidth) + _
            WIN32.GetSystemMetric(SM_CXVSCROLL) _
        ), _
        Me.ScaleHeight - (Me.fraToolbar.Top + Me.fraToolbar.Height) _
    ): Next i
    
    'Blocks (Layout) Side Pane: _
     ----------------------------------------------------------------------------------
    'Position the tool bar
    Call Me.fraBlocksToolbar.Move( _
        0, 0, Me.fraSidePane(0).Width, 1200 _
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
        0, Me.fraBlocksToolbar.Top + Me.fraBlocksToolbar.Height + blu.Ypx, _
        Me.fraSidePane(0).Width, _
        Me.fraSidePane(0).Height _
        - (Me.fraBlocksToolbar.Top + Me.fraBlocksToolbar.Height + blu.Ypx) _
    )
    
    'Toolbar Buttons: _
     ----------------------------------------------------------------------------------
    'The title of the level; this will eventually be a custom text box
    Call Me.lblTitle.Move( _
        0, Me.fraSidePane(0).Left + Me.fraSidePane(0).Width _
    )
    
    With Me.lineSplit(0)
        Let .X1 = Me.fraSidePane(0).Left + Me.fraSidePane(0).Width: Let .X2 = .X1
    End With
    
    'The display that shows the number of rings on the level
    Let Me.picRings.Left = Me.fraSidePane(0).Left + Me.fraSidePane(0).Width 'Me.lineSplit(0).X1 + blu.Xpx(8)
    
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
    
    'Right-hand aligned buttons (working backwards from the right):
    Let Me.btnShare.Left = Me.fraToolbar.Width - Me.btnShare.Width
    With Me.lineSplit(3)
        Let .X1 = Me.btnShare.Left - blu.Xpx(8): Let .X2 = .X1
    End With
    'Zoom levels
    Let Me.btnZoomTV.Left = Me.lineSplit(3).X1 - blu.Xpx(8) - Me.btnZoomTV.Width
    Let Me.btnZoom2.Left = Me.btnZoomTV.Left - Me.btnZoom2.Width
    Let Me.btnZoom1.Left = Me.btnZoom2.Left - Me.btnZoom1.Width
    Let Me.lblZoom.Left = Me.btnZoom1.Left - Me.lblZoom.Width
    'Grid ON/OFF
    Let Me.btnGrid.Left = Me.lblZoom.Left - Me.btnGrid.Width
    Let Me.lblGrid.Left = Me.btnGrid.Left - Me.lblGrid.Width
    
    'Level viewport: _
     ----------------------------------------------------------------------------------
    'Reposition the level viewport
    Call Me.vwpLevel.Move( _
        Me.fraSidePane(0).Left + Me.fraSidePane(0).Width, _
        Me.fraToolbar.Top + Me.fraToolbar.Height, _
        Me.ScaleWidth - (Me.fraSidePane(0).Left + Me.fraSidePane(0).Width), _
        Me.ScaleHeight - (Me.fraToolbar.Top + Me.fraToolbar.Height) _
    )
    Call picSizer.Move( _
        Me.ScaleWidth - Me.picSizer.Width, _
        Me.ScaleHeight - Me.picSizer.Height _
    )
        
    'Set the minimum size of the form: _
     ----------------------------------------------------------------------------------
    'TODO: This isn't finalised. Should be done on mdiMain rather than from MDI child
    
    Let mdiMain.bluWindow.MinWidth = ( _
        Me.bluTab.Width + Me.fraSidePane(0).Width + _
        Me.picRings.Width + _
        Me.lblZoom.Width + Me.btnZoom1.Width + Me.btnZoom2.Width + Me.btnZoomTV.Width + _
        Me.btnShare.Width _
    ) \ Screen.TwipsPerPixelX
    
    Let mdiMain.bluWindow.MinHeight = ( _
        mdiMain.toolbar.Height + Me.bluTab.Top + Me.bluTab.Height _
    ) \ Screen.TwipsPerPixelY
End Sub

'FORM Unload _
 ======================================================================================
Private Sub Form_Unload(Cancel As Integer)
    'Detatch the current level from the form
    Set Level = Nothing
    Set Sprites = Nothing
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
    'Recalculate where the water line exists on the level height
    Let c.WaterLevelPx = x32(My_Level.ObjectLayout.WaterLevel) - 4
    'And refresh
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

'EVENT vwpLevel MOUSEOUT : Mouse has gone out of the level viewport _
 ======================================================================================
Private Sub vwpLevel_MouseOut()
    Call mdiMain.SetTip
    'Clear the hover box
    Let c.Hover.X = -1: Let c.Hover.Y = -1

    'TODO: Render selection rectangle on the viewport
'   Call RepaintLevel
End Sub

'EVENT vwpBlocks MOUSEUP _
 ======================================================================================
Private Sub vwpBlocks_MouseUp(Button As MouseButtonConstants, Shift As ShiftConstants, X As Single, Y As Single)
    'Which block was clicked?
    Dim Index As Long
    Let Index = _
        ((Y \ 33) * PaneBlockWidth) _
        + (X \ 33)
    
    'Only allow clicks within the block range
    If Index >= 0 Or Index = (Level.BlockMapping.Length - 1) Then
        'Which mouse button was clicked?
        If Button = 1 Then
            Let Me.BlockSelectLeft = Index
        ElseIf Button = 2 Then
            Let Me.BlockSelectRight = Index
        End If
    End If
End Sub

Private Sub Command3_Click()
    ROM.Export Run.Path & "MaSS1VE.sms", _
               Run.Path & "Sonic the Hedgehog (1991)(Sega).sms", _
               0
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
        'Set the level title
        Let Me.lblTitle.Caption = My_Level.Title
        
        'Cache the block mappings for speedy painting of the level
        Call CacheBlocks
        
        'Reset the select block indexes _
         (these will be repainted; dependent on the above)
        Let Me.BlockSelectLeft = 1
        Let Me.BlockSelectRight = 0
        
        'Recalculate where the water line exists on the level height
        Let c.WaterLevelPx = x32(My_Level.ObjectLayout.WaterLevel) - 4
        
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
        'Determine the height of the block list based on _
         a. the number of blocks in the mapping _
         b. the number of blocks in the width of the side pane
        'This is trickier than it sounds because of rounding problems where you need _
         one extra row for one or two blocks
        Let c.BlockListHeight = _
            x32(Lib.RoundUp(Level.BlockMapping.Length / PaneBlockWidth)) + _
            Lib.RoundUp(Level.BlockMapping.Length / PaneBlockWidth)
        
        'Create the blocks image in the viewport
        Let vwpBlocks.BackColor = blu.BaseColour
        Call vwpBlocks.SetImageProperties( _
            Width:=x32(PaneBlockWidth) + PaneBlockWidth - 1, _
            Height:=c.BlockListHeight _
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
    End If
End Property

'PROPERTY Zoom _
 ======================================================================================
Public Property Get Zoom() As Long: Let Zoom = My_Zoom: End Property
Public Property Let Zoom(ByVal ZoomLevel As Long)
    Let My_Zoom = ZoomLevel
    
    Let c.BlockSize = x32(My_Zoom)
    Me.btnZoom1.State = IIf(My_Zoom = 1, bluSTATE.Active, bluSTATE.Inactive)
    Me.btnZoom2.State = IIf(My_Zoom = 2, bluSTATE.Active, bluSTATE.Inactive)
    Me.btnZoomTV.State = IIf(My_Zoom = 3, bluSTATE.Active, bluSTATE.Inactive)
    
    Call Form_Resize
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

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
    
    'Some frmEditor specifics
    Let Me.fraToolbar.BackColor = ActiveColour
    Let Me.lineSplit(0).BorderColor = ActiveColour
    Let Me.lineSplit(1).BorderColor = ActiveColour
    Let Me.lineSplit(2).BorderColor = ActiveColour
    Let Me.lineSplit(3).BorderColor = ActiveColour
    Let Me.fraBlocksToolbar.BackColor = ActiveColour
    
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

'HandleMouse : The MouseDown / Move / Up events are handled similarly _
 ======================================================================================
Private Sub HandleMouse(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    Dim iX As Long, iY As Long
    
    'If the level is thinner/shorter than the viewport it'll be centred; if the mouse _
     is above / to the left of where the level is centred, there's no point going any _
     further, skip ahead to "mouse out" now
    If X - c.Centre.X <= 0 Then GoTo Invalid
    If Y - c.Centre.Y <= 0 Then GoTo Invalid
    
    'Calculate which block the mouse is over
    Let iX = c.Block.X + (((X - c.Centre.X) + (c.BlockPxOffset.X * My_Zoom)) \ c.BlockSize)
    Let iY = c.Block.Y + (((Y - c.Centre.Y) + (c.BlockPxOffset.Y * My_Zoom)) \ c.BlockSize)
    
    'Is this within level boundaries? (If the level is centred in the viewport then _
     there will be space below / to the right)
    If iX >= 0 And iX < Level.Width And _
       iY >= 0 And iY < Level.Height _
    Then
        'Just hovering, thanks
        If Button = 0 Then
            Call mdiMain.SetTip( _
                "L/R CLICK: Set blocks CTRL: Select an area" _
            )
            
        'Left Click:
        ElseIf Button = 1 Then
            'Change the block to the one set to the left mouse button. _
             The `My_Level_BlockMappingChange` event will paint the changed block _
             and refresh the viewport for us (we don't need to repaint the whole level)
            Let Level.FloorLayout.Block(iX, iY) = BlockSelectLeft
            
        'Right Click:
        ElseIf Button = 2 Then
            'As above, but with the right mouse button
            Let Level.FloorLayout.Block(iX, iY) = BlockSelectRight
        
        'Middle Click:
        ElseIf Button = 4 Then
            'Pick the current block under the mouse cursor
            Let BlockSelectLeft = Level.FloorLayout.Block(iX, iY)
            
        End If
        
        'Is the mouse in a different block than it was last time?
        Dim DoRender As Boolean
        If iX <> c.Hover.X Or iY <> c.Hover.Y Then
            Let c.Hover.X = iX: Let c.Hover.Y = iY
            'TODO: Render the selection rectangle
        End If
        
    Else
Invalid:
        'Clear the help tip when in the viewport, but not on a map block
        Call mdiMain.SetTip
        'The mouse is not over any block; if this is a change to previously, refresh
        If iX <> -1 And iY <> -1 Then
            Let c.Hover.X = -1: Let c.Hover.Y = -1
            'TODO: Render the selection rectangle
        Else
            Let c.Hover.X = -1: Let c.Hover.Y = -1
        End If
    End If
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
            vwpLevel_hDC1, 0, c.WaterLevelPx _
        )
    End If
    
    Debug.Print "Repaint Object Layout - " & Round(Timer - StartTimeO, 4)
    
'    'Mouse hover rectangle: _
'     ----------------------------------------------------------------------------------
'    'Is the mouse hovered over a block?
'    If c.Hover.X <> -1 And c.Hover.Y <> -1 Then
'        Dim px As Long, pY As Long
'        Let px = c.Centre.X + ((c.Hover.X - c.Block.X) * c.BlockSize) - (c.BlockPxOffset.X * My_Zoom)
'        Let pY = c.Centre.Y + ((c.Hover.Y - c.Block.Y) * c.BlockSize) - (c.BlockPxOffset.Y * My_Zoom)
'
'        Me.picRender.Line (px, pY)-(px + c.BlockSize + 2, pY + c.BlockSize + 2), &H202020, B
'        Me.picRender.Line (px - 1, pY - 1)-(px + c.BlockSize, pY + c.BlockSize), &HF8F8F8, B
'        Me.picRender.Line (px - 2, pY - 2)-(px + c.BlockSize + 1, pY + c.BlockSize + 1), &HF8F8F8, B
'    End If
        
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
