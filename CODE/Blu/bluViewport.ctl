VERSION 5.00
Begin VB.UserControl bluViewport 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "bluViewport.ctx":0000
End
Attribute VB_Name = "bluViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'blu : A Modern Metro-esque graphical toolkit; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CONTROL :: bluViewport

'Provides a backbuffered image display with built in API-driven native scroll bars, _
 including mouse wheel support
 
'A multi-layered full image backbuffer is provided for you so that you don't have to _
 maintain your own, and so bluViewport can manage the scrolling and painting _
 automatically. It's great for caching as you only have to repaint a layer when it _
 changes, not every time the window scrolls

'Status             Ready, but incomplete
'Dependencies       blu.bas, bluImage.cls, bluMagic.cls, bluMouseEvents.cls
'Last Updated       24-JUN-14
'Last Update        Removed Lib.bas dependency, _
                    Fixed old bug with scrollbars wrongly showing when image width _
                     and/or height is the same as the viewport

'TODO: When scrolling, include mouse button / key state with the mouse move event sent _
       (this info should come from bluMouseEvents)

'--------------------------------------------------------------------------------------

'This was made with the help of "Adding Scroll Bars to Forms, PictureBoxes and _
 User Controls" by Steve McMahon, though my own work _
 <www.vbaccelerator.com/article.asp?id=2185>
 
'NOTE: The `AutoRedraw` property of this control is False. Since we are handling _
 `WM_PAINT` ourselves, we do not need to use VB's backbuffer

'"If AutoRedraw is set to True for a form or PictureBox container, hDC acts as a _
  handle to the device context of the persistent graphic. When AutoRedraw is False, _
  hDC is the actual hDC value of the Form window or the PictureBox container." _
 <msdn.microsoft.com/en-us/library/aa267506%28v=vs.60%29.aspx>

'/// API DEFS /////////////////////////////////////////////////////////////////////////

'Show / hide scrollbars _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb787601%28v=vs.85%29.aspx>
Private Declare Function user32_ShowScrollBar Lib "user32" Alias "ShowScrollBar" ( _
    ByVal hndWindow As Long, _
    ByVal Bar As SB, _
    ByVal Show As BOOL _
) As BOOL

Private Enum SB
    SB_LEFT = 6
    SB_TOP = 6
    
    SB_RIGHT = 7
    SB_BOTTOM = 7
    
    SB_ENDSCROLL = 8
    
    SB_LINELEFT = 0
    SB_LINEUP = 0
    
    SB_PAGELEFT = 2
    SB_PAGEUP = 2
    
    SB_LINERIGHT = 1
    SB_LINEDOWN = 1
    
    SB_PAGERIGHT = 3
    SB_PAGEDOWN = 3
    
    SB_THUMBPOSITION = 4
    SB_THUMBTRACK = 5
    
End Enum

'Structure used when getting / setting the scroll bar properties _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb787537%28v=vs.85%29.aspx>
Private Type SCROLLINFO
    SizeOfMe As Long
    Mask As SIF                 'Filter which properties to read / write
    Min As Long                 'Lowest value, must be positive
    Max As Long                 'Highest value, must be positive. Note that this must _
                                 also include 1x Page size below
    Page As Long                'The size of one page of scroll (i.e. the viewport _
                                 width / height), e.g. when the user clicks the track
    Pos As Long                 'Current value of the scrollbar
    TrackPos As Long            'The value of the scroll box when being dragged
End Type

Private Enum SIF
    SIF_RANGE = &H1             'Get/set Min and max
    SIF_PAGE = &H2              'Get/set Page value
    SIF_POS = &H4               'Get/set the scroll value
    SIF_DISABLENOSCROLL = &H8   'Disable the scroll bar instead of hiding
    SIF_TRACKPOS = &H10         'Get position of the scroll box when dragging it
    'All of the above:
    SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
End Enum

'<msdn.microsoft.com/en-us/library/windows/desktop/bb787583%28v=vs.85%29.aspx>
Private Declare Function user32_GetScrollInfo Lib "user32" Alias "GetScrollInfo" ( _
    ByVal hndWindow As Long, _
    ByVal Bar As bluScrollBar, _
    ByRef Info As SCROLLINFO _
) As BOOL

'<msdn.microsoft.com/en-us/library/windows/desktop/bb787595%28v=vs.85%29.aspx>
Private Declare Function user32_SetScrollInfo Lib "user32" Alias "SetScrollInfo" ( _
    ByVal hndWindow As Long, _
    ByVal Bar As bluScrollBar, _
    ByRef Info As SCROLLINFO, _
    ByVal Redraw As BOOL _
) As Long

'Stuff happening in the subclass _
 --------------------------------------------------------------------------------------

'<msdn.microsoft.com/en-us/library/windows/desktop/dd162768%28v=vs.85%29.aspx>
Private Type PAINTSTRUCT
  hndDC As Long
  Erase As BOOL
  PaintRECT As RECT
  Restore As BOOL
  IncUpdate As BOOL
  Reserved(0 To 31) As Byte
End Type

'<msdn.microsoft.com/en-us/library/windows/desktop/dd183362%28v=vs.85%29.aspx>
Private Declare Function user32_BeginPaint Lib "user32" Alias "BeginPaint" ( _
    ByVal hndWindow As Long, _
    ByRef Paint As PAINTSTRUCT _
) As Long

'<msdn.microsoft.com/en-us/library/windows/desktop/dd162598%28v=vs.85%29.aspx>
Private Declare Function user32_EndPaint Lib "user32" Alias "EndPaint" ( _
    ByVal hndWindow As Long, _
    ByRef Paint As PAINTSTRUCT _
) As BOOL

'Scroll a window - shifts the pixels and sends a WM_PAINT message to fill in the blanks _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb787593%28v=vs.85%29.aspx>
Private Declare Function user32_ScrollWindowEx Lib "user32" Alias "ScrollWindowEx" ( _
    ByVal hndWindow As Long, _
    ByVal ScrollX As Long, _
    ByVal ScrollY As Long, _
    ByVal ptrScrollRECT As Long, _
    ByVal ptrClipRECT As Long, _
    ByVal ptrUpdateRegion As Long, _
    ByRef ptrUpdateRECT As Long, _
    ByVal Flags As Long _
) As Long
Private Const SW_INVALIDATE As Long = &H2

'<msdn.microsoft.com/en-us/library/windows/desktop/dd145167%28v=vs.85%29.aspx>
Private Declare Function user32_UpdateWindow Lib "user32" Alias "UpdateWindow" ( _
    ByVal hndWindow As Long _
) As BOOL

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'We'll need to subclass the control to listen into the scroll bar events
Private Magic As bluMagic

'This will track mouse in / out and mouse wheel events
Private WithEvents MouseEvents As bluMouseEvents
Attribute MouseEvents.VB_VarHelpID = -1

'bluViewport allows you to manage multiple layers for the whole image to minimise _
 the amount of painting you have to do and so bluViewport can manage the scrolling
Private Type Layer
    IMAGE As bluImage
End Type
Private NumberOfLayers As Long
Private Layers() As Layer

'This is a buffer the size of the viewport, used for flicker-free drawing
Private Buffer As bluImage

'To try be fast as possible we cache various values here:
Private Type CACHEVARS
    UserControl_BackColor As Long       'Back colour, but already OLE translated
    ClientRECT As RECT                  'The width / height of the viewport
    ImageRECT As RECT                   'The whole image's size
    DC_BRUSH As Long                    'The stock colour brush built in to DCs
    
    Info(0 To 1) As SCROLLINFO          'The scroll bar properties (HORZ / VERT)
    
    'If the viewport is larger than the image, then we will centre it. This will give _
     the Top / Left offset for where the image is centred in the viewport
    Centre As POINT
    'The destination size to paint. If the image is centred because it is smaller than _
     the viewport, the destination size will be less than the viewport width / height
    Dst As SIZE
    'The source portion of the image. At Zoom=1, this is the same as Dst, but when _
     zoomed, it's a smaller area, that is stretched to the Dst size
    Src As SIZE
End Type
Private c As CACHEVARS

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Public Enum bluScrollBar
    HORZ = 0
    VERT = 1
End Enum

Private My_ScrollAmount(0 To 1) As Long 'Amount to scroll clicking scroll arrow once
Private My_ScrollLineSize As Long       'Size of a "line" for mouse wheel scrolling
Private My_ScrollCharSize As Long       'Size of a "char" for horizontal mouse wheel

Private My_Centre As Boolean            'Centre the image if smaller than the viewport?

Private My_Zoom As Long                 'Zoom level
Private My_ZoomMin As Long              'Minimum zoom level (i.e. 1)
Private My_ZoomMax As Long              'Maximum zoom level

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'When the mouse leaves the viewport
Event MouseOut()
'And when it enters. This will fire just once instead of continuously like MouseMove
Event MouseIn()
'When the mouse stays in place for a brief period of time. _
 This is used for tooltips, for example
Event MouseHover( _
    ByVal Button As VBRUN.MouseButtonConstants, _
    ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single, _
    ByVal ImageX As Long, ByVal ImageY As Long _
)

'The standard mouse events we'll forward
Event Click()
Attribute Click.VB_UserMemId = -600
Event MouseDown( _
    ByVal Button As VBRUN.MouseButtonConstants, ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single, _
    ByVal ImageX As Long, ByVal ImageY As Long _
)
Event MouseMove( _
    ByVal Button As VBRUN.MouseButtonConstants, ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single, _
    ByVal ImageX As Long, ByVal ImageY As Long _
)
Event MouseUp( _
    ByVal Button As VBRUN.MouseButtonConstants, ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single, _
    ByVal ImageX As Long, ByVal ImageY As Long _
)

'An event that occurs when painting, allowing you to alter the viewport's display
Event Paint(ByVal hDC As Long)

'When a scroll occurs
Event Scroll(ByVal ScrollX As Long, ByVal ScrollY As Long)

'When the user zooms the viewport via Ctrl+Zoom. This event doesn't fire when the _
 zoom is changed via the property to avoid potential infinite loops with your code
Event Zoom(ByVal Direction As Long)

'CONTROL Click _
 ======================================================================================
Private Sub UserControl_Click(): RaiseEvent Click: End Sub

'CONTROL Initialize _
 ======================================================================================
Private Sub UserControl_Initialize()
    'The DC brush helps us avoid having to create and destroy a brush when we want _
     to paint. It acts as a built-in brush that we can set the colour on at will _
     <blogs.msdn.com/b/oldnewthing/archive/2005/04/20/410031.aspx>
    Let c.DC_BRUSH = blu.gdi32_GetStockObject(DC_BRUSH)
    
    'The `SCROLLINFO` structures must state their size, _
     let's do this just once
    Let c.Info(HORZ).SizeOfMe = Len(c.Info(HORZ))
    Let c.Info(VERT).SizeOfMe = Len(c.Info(VERT))
    
    Let My_Zoom = 1
End Sub

'CONTROL InitProperties : New instance of the control plopped on a form _
 ======================================================================================
Private Sub UserControl_InitProperties()
    Let Me.ScrollAmountH = 32
    Let Me.ScrollAmountV = 32
    Let Me.ScrollLineSize = 16
    Let Me.ScrollCharSize = 16
    Let Me.ZoomMin = 1
    Let Me.ZoomMax = 16
End Sub

'CONTROL KeyDown : Handle keyboard control of scrolling _
 ======================================================================================
Private Sub UserControl_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    Dim Send As Long, Scroll As WM
    Let Send = -1
    
    'If holding SHIFT, do horizontal scroll
    Let Scroll = IIf( _
        (Shift And VBRUN.ShiftConstants.vbShiftMask) <> 0, _
        WM_HSCROLL, WM_VSCROLL _
    )
    
    Select Case KeyCode
        'Arrow keys are specific and not overriden with SHIFT
        Case vbKeyLeft:     Let Scroll = WM_HSCROLL: Let Send = SB.SB_LINELEFT
        Case vbKeyRight:    Let Scroll = WM_HSCROLL: Let Send = SB.SB_LINERIGHT
        Case vbKeyUp:       Let Scroll = WM_VSCROLL: Let Send = SB.SB_LINEUP
        Case vbKeyDown:     Let Scroll = WM_VSCROLL: Let Send = SB.SB_LINEDOWN
        'Page Up/Down & Home/End can be horizontal if SHIFT held
        Case vbKeyPageUp:   Let Send = SB.SB_PAGEUP
        Case vbKeyPageDown: Let Send = SB.SB_PAGEDOWN
        Case vbKeyHome:     Let Send = SB.SB_TOP
        Case vbKeyEnd:      Let Send = SB.SB_BOTTOM
    End Select
    If Send <> -1 Then
        'Send the `WM_HSCROLL` / `WM_VSCROLL` message. _
         See the subclass section at the bottom of the file for details
        Call blu.user32_SendMessage(UserControl.hWnd, Scroll, Send, 0)
    End If
End Sub

'CONTROL MouseDown _
 ======================================================================================
Private Sub UserControl_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y, GetImageX(X), GetImageY(Y))
End Sub

'CONTROL MouseMove _
 ======================================================================================
Private Sub UserControl_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y, GetImageX(X), GetImageY(Y))
End Sub

'CONTROL MouseUp _
 ======================================================================================
Private Sub UserControl_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y, GetImageX(X), GetImageY(Y))
End Sub

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Let Me.BackColor = .ReadProperty(Name:="BackColor", DefaultValue:=VBRUN.SystemColorConstants.vbApplicationWorkspace)
        Let My_Centre = .ReadProperty(Name:="Centre", DefaultValue:=True)
        Let My_ScrollAmount(HORZ) = .ReadProperty(Name:="ScrollAmountH", DefaultValue:=32)
        Let My_ScrollAmount(VERT) = .ReadProperty(Name:="ScrollAmountV", DefaultValue:=32)
        Let My_ScrollLineSize = .ReadProperty(Name:="ScrollLineSize", DefaultValue:=16)
        Let My_ScrollCharSize = .ReadProperty(Name:="ScrollCharSize", DefaultValue:=64)
        Let My_ZoomMin = .ReadProperty(Name:="ZoomMin", DefaultValue:=1)
        Let My_ZoomMax = .ReadProperty(Name:="ZoomMax", DefaultValue:=16)
    End With
    
    'Only subclass if not in VB's design mode
    If blu.UserMode = True Then
        'Attach the mouse events to look for mouse enter / leave / wheel
        Set MouseEvents = New bluMouseEvents
        Call MouseEvents.Attach( _
            UserControl.hWnd, _
            blu.GetParentForm(StartWith:=UserControl.Parent, GetMDIParent:=True).hWnd _
        )
        'Subclass the control to listen to scroll bar events
        Set Magic = New bluMagic
        Call Magic.ssc_Subclass(UserControl.hWnd, 0, 1, Me)
        Call Magic.ssc_AddMsg( _
            UserControl.hWnd, MSG_BEFORE, _
            WM_PAINT, WM_ERASEBKGND, WM_HSCROLL, WM_VSCROLL, _
            WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMBUTTONDOWN _
        )
    End If
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize()
    'Reconfigure the scroll bar parameters, this will deal with the change in client _
     size as the showing / hiding of the scrollbars changes the client size
    Call InitScrollBars
    'Refresh the viewport
    Call Me.Refresh
End Sub

'CONTROL Terminate _
 ======================================================================================
Private Sub UserControl_Terminate()
    Erase Layers
    Set Buffer = Nothing
    
    'Carefully detatch the subclassing
    Set MouseEvents = Nothing
    If Not Magic Is Nothing Then
        Call Magic.ssc_DelMsg( _
            UserControl.hWnd, MSG_BEFORE, _
            WM_PAINT, WM_ERASEBKGND, WM_HSCROLL, WM_VSCROLL, _
            WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMBUTTONDOWN _
        )
        Call Magic.ssc_UnSubclass(UserControl.hWnd)
        Set Magic = Nothing
    End If
End Sub

'CONTROL WriteProperties _
 ======================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(Name:="BackColor", Value:=UserControl.BackColor, DefaultValue:=VBRUN.SystemColorConstants.vbApplicationWorkspace)
        Call .WriteProperty(Name:="Centre", Value:=My_Centre, DefaultValue:=True)
        Call .WriteProperty(Name:="ScrollAmountH", Value:=My_ScrollAmount(HORZ), DefaultValue:=32)
        Call .WriteProperty(Name:="ScrollAmountV", Value:=My_ScrollAmount(VERT), DefaultValue:=32)
        Call .WriteProperty(Name:="ScrollLineSize", Value:=My_ScrollLineSize, DefaultValue:=16)
        Call .WriteProperty(Name:="ScrollCharSize", Value:=My_ScrollCharSize, DefaultValue:=64)
        Call .WriteProperty(Name:="ZoomMin", Value:=My_ZoomMin, DefaultValue:=1)
        Call .WriteProperty(Name:="ZoomMax", Value:=My_ZoomMax, DefaultValue:=16)
    End With
End Sub

'EVENT MouseEvents MOUSEIN : Mouse entered the viewport control _
 ======================================================================================
Private Sub MouseEvents_MouseIn(): RaiseEvent MouseIn: End Sub

'EVENT MouseEvents MOUSEOUT : Mouse went out of the viewport control _
 ======================================================================================
Private Sub MouseEvents_MouseOut(): RaiseEvent MouseOut: End Sub

'EVENT MouseEvents MOUSEHOVER _
 ======================================================================================
Private Sub MouseEvents_MouseHover(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single)
    RaiseEvent MouseHover(Button, Shift, X, Y, GetImageX(X), GetImageY(Y))
End Sub

'EVENT MouseEvents MOUSEHSCROLL _
 ======================================================================================
Private Sub MouseEvents_MouseHScroll(ByVal CharsScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single)
    'Scroll the viewport...
    With c.Info(HORZ)
        Let .Mask = SIF_POS
        'For increased zoom, we need to dampen the scrolling speed!
        Let .Pos = blu.Range( _
            InputNumber:=.Pos - blu.NotZero( _
                InputNumber:=(CharsScrolled * My_ScrollCharSize) \ My_Zoom, _
                AtLeast:=Sgn(CharsScrolled) _
            ), _
            Maximum:=Me.ScrollMax(HORZ), Minimum:=.Min _
        )
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, HORZ, c.Info(HORZ), API_TRUE)
    RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
    
    'Raise a mouse move event since the pointer is no longer under the part of the _
     image it was before and the controller might need the new ImageX/Y values
    RaiseEvent MouseMove(Button, Shift, X, Y, GetImageX(X), GetImageY(Y))
    
    'The viewport is refreshed _after_ the events fire so that your controller _
     does *not* have to call `Refresh` itself, saving repaints
    Call Me.Refresh
End Sub

'EVENT MouseEvents MOUSEVSCROLL _
 ======================================================================================
Private Sub MouseEvents_MouseVScroll(ByVal LinesScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single)
    'If Ctrl is held, zoom
    If (Shift And vbCtrlMask) > 0 Then
        'The zoom property will automatically handle the min / max limit
        Let Me.Zoom = Me.Zoom + Sgn(LinesScrolled)
        RaiseEvent Zoom(Sgn(LinesScrolled))
    Else
        'Scroll the viewport...
        With c.Info(VERT)
            Let .Mask = SIF_POS
            'For increased zoom, we need to dampen the scrolling speed!
            Let .Pos = blu.Range( _
                InputNumber:=.Pos - blu.NotZero( _
                    InputNumber:=(LinesScrolled * My_ScrollLineSize) \ My_Zoom, _
                    AtLeast:=Sgn(LinesScrolled) _
                ), _
                Maximum:=Me.ScrollMax(VERT), Minimum:=.Min _
            )
        End With
        Call user32_SetScrollInfo(UserControl.hWnd, VERT, c.Info(VERT), API_TRUE)
        RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
    End If
    
    'Raise a mouse move event since the pointer is no longer under the part of the _
     image it was before and the controller might need the new ImageX/Y values
    RaiseEvent MouseMove(Button, Shift, X, Y, GetImageX(X), GetImageY(Y))
    
    'The viewport is refreshed _after_ the events fire so that your controller _
     does *not* have to call `Refresh` itself, saving repaints
    Call Me.Refresh
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////

'PROPERTY BackColor : the colour behind the viewport's image _
 ======================================================================================
Public Property Get BackColor() As OLE_COLOR: Let BackColor = UserControl.BackColor: End Property
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let BackColor(ByVal Color As OLE_COLOR)
    Let UserControl.BackColor = Color
    'Cache the new back colour ready for painting. If it's a system colour _
     (e.g. `vbApplicationWorkspace`, then translate it to the real colour)
    Let c.UserControl_BackColor = blu.OLETranslateColor(Color)
    'Apply the colour to the back buffer DC, this will automatically be used at paint
    If Not Buffer Is Nothing Then
        Call blu.gdi32_SetDCBrushColor( _
            Buffer.hDC, c.UserControl_BackColor _
        )
    End If
    Call Me.Refresh
    Call UserControl.PropertyChanged("BackColor")
End Property

'PROPERTY Centre : Whether to centre the image if smaller than the viewport _
 ======================================================================================
Public Property Get Centre() As Boolean: Let Centre = My_Centre: End Property
Public Property Let Centre(ByVal State As Boolean)
    Let My_Centre = State
    'Recalculate the scroll bar limits (will update the centering)
    Call InitScrollBars
    
    'Raise a mouse move event since the pointer is no longer under the part of the _
     image it was before and the controller might need the new ImageX/Y values
    Dim MousePos As blu.POINT
    Let MousePos = blu.GetMousePos_Window(UserControl.hWnd)
    'TODO: Get mouse button / key state
    RaiseEvent MouseMove( _
        0, 0, MousePos.X, MousePos.Y, GetImageX(MousePos.X), GetImageY(MousePos.Y) _
    )
    
    'The viewport is refreshed _after_ the events fire so that your controller _
     does *not* have to call `Refresh` itself, saving repaints
    Call Me.Refresh
    Call UserControl.PropertyChanged("Centre")
End Property

'PROPERTY CentreX : The horizontal image offset if image is narrower than the viewport _
 ======================================================================================
Public Property Get CentreX() As Long: Let CentreX = c.Centre.X: End Property

'PROPERTY CentreY : The vertical image offset if image is shorter than the viewport _
 ======================================================================================
Public Property Get CentreY() As Long: Let CentreY = c.Centre.Y: End Property

'PROPERTY hDC : Handle to the device context for the image layer -- not the control _
 ======================================================================================
Public Property Get hDC(Optional ByVal Layer As Long = 0) As Long
    'If you want to paint directly on the viewport use the viewport's `Paint` event, _
     this is double-buffered so you won't get any flicker
    If NumberOfLayers <> 0 And Layer >= 0 And Layer < NumberOfLayers Then
        Let hDC = Layers(Layer).IMAGE.hDC
    End If
End Property

'PROPERTY ImageWidth _
 ======================================================================================
Public Property Get ImageWidth() As Long
    If NumberOfLayers = 0 Then Exit Property
    Let ImageWidth = Layers(0).IMAGE.Width
End Property

'PROPERTY ImageHeight _
 ======================================================================================
Public Property Get ImageHeight() As Long
    If NumberOfLayers = 0 Then Exit Property
    Let ImageHeight = Layers(0).IMAGE.Height
End Property

'GET PaletteColour : Get an individual palette colour for an 8-bit layer: _
 ======================================================================================
Public Property Get PaletteColour( _
    ByVal Index As Byte, ByRef Layer As Long _
) As OLE_COLOR
    Dim QuadColour As blu.RGBQUAD
    
    'TODO: Is layer within range?
    'Only available for 8-bit layers!
    If Layers(Layer).IMAGE.Depth <> [8-Bit] Then Let PaletteColour = -1: Exit Property
    
    'Try to fetch the colour
    If gdi32_GetDIBColorTable( _
        hndDeviceContext:=Layers(Layer).IMAGE.hDC, _
        StartIndex:=Index, Count:=1, _
        ptrRGBQUAD:=QuadColour _
    ) = 1 Then
        Let PaletteColour = VBA.RGB(QuadColour.Red, QuadColour.Green, QuadColour.Blue)
    Else
        Let PaletteColour = -1
    End If
End Property

'LET PaletteColour : Set an individual palette colour for an 8-bit layer: _
 ======================================================================================
Public Property Let PaletteColour( _
    ByVal Index As Byte, ByRef Layer As Long, _
    NewColour As OLE_COLOR _
)
    'TODO: Is layer within range?
    'Only available for 8-bit layers!
    If Layers(Layer).IMAGE.Depth <> [8-Bit] Then Exit Property
    
    'Should the colour be a system colour (like "button face"), get the real colour
    Let NewColour = blu.OLETranslateColor(NewColour)

    'Prepre the colour structure to use with the API call
    Dim QuadColour As blu.RGBQUAD
    With QuadColour
        .Blue = (NewColour And &HFF0000) \ &H10000
        .Green = (NewColour And &HFF00&) \ &H100
        .Red = (NewColour And &HFF)
    End With

    'Push the colour in
    Call gdi32_SetDIBColorTable( _
        hndDeviceContext:=Layers(Layer).IMAGE.hDC, _
        StartIndex:=Index, Count:=1, _
        ptrRGBQUAD:=QuadColour _
    )
    
    'Need to repaint to see the colour change
    Call Me.Refresh
End Property

'PROPERTY ScrollMax : Return the maximum scroll value _
 ======================================================================================
'You can't set this value as it is automatically managed by the viewport based on the _
 image size and viewport size
Public Property Get ScrollMax(ByVal Bar As bluScrollBar) As Long
Attribute ScrollMax.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Let ScrollMax = blu.Min( _
        c.Info(Bar).Max - c.Info(Bar).Page _
    )
End Property

'PROPERTY ScrollAmountH : The amount to scroll when clicking the scroll arrows once _
 ======================================================================================
Public Property Get ScrollAmountH() As Long: Let ScrollAmountH = My_ScrollAmount(HORZ): End Property
Attribute ScrollAmountH.VB_Description = "The amount to scroll (horizontally) clicking the scroll arrow once"
Attribute ScrollAmountH.VB_ProcData.VB_Invoke_Property = ";Behavior"
Public Property Let ScrollAmountH(ByVal Value As Long)
    Let My_ScrollAmount(HORZ) = Value
    Call UserControl.PropertyChanged("ScrollAmountH")
End Property

'PROPERTY ScrollAmountV : The amount to scroll when clicking the scroll arrows once _
 ======================================================================================
Public Property Get ScrollAmountV() As Long: Let ScrollAmountV = My_ScrollAmount(VERT): End Property
Attribute ScrollAmountV.VB_Description = "The amount to scroll (vertically) clicking the scroll arrow once"
Attribute ScrollAmountV.VB_ProcData.VB_Invoke_Property = ";Behavior"
Public Property Let ScrollAmountV(ByVal Value As Long)
    Let My_ScrollAmount(VERT) = Value
    Call UserControl.PropertyChanged("ScrollAmountV")
End Property

'PROPERTY ScrollLineSize : The size of a "line" for mouse wheel scrolling _
 ======================================================================================
Public Property Get ScrollLineSize() As Long: Let ScrollLineSize = My_ScrollLineSize: End Property
Attribute ScrollLineSize.VB_Description = "The size (in px) of a ""line"" for vertical mouse wheel scrolling."
Attribute ScrollLineSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
Public Property Let ScrollLineSize(ByVal Value As Long)
    Let My_ScrollLineSize = Value
    Call UserControl.PropertyChanged("ScrollLineSize")
End Property

'PROPERTY ScrollCharSize : The size of a "char" for horizontal wheel scrolling _
 ======================================================================================
Public Property Get ScrollCharSize() As Long: Let ScrollCharSize = My_ScrollCharSize: End Property
Attribute ScrollCharSize.VB_Description = "The size (in px) of a ""char"" for horizontal mouse wheel scrolling."
Attribute ScrollCharSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
Public Property Let ScrollCharSize(ByVal Value As Long)
    Let My_ScrollCharSize = Value
    Call UserControl.PropertyChanged("ScrollCharSize")
End Property

'PROPERTY ScrollX : Scroll the viewport to a specific place horizontally _
 ======================================================================================
Public Property Get ScrollX() As Long: Let ScrollX = c.Info(HORZ).Pos: End Property
Public Property Let ScrollX(ByVal Value As Long)
    'Scroll the viewport...
    With c.Info(HORZ)
        Let .Mask = SIF_POS
        Let .Pos = blu.Max(.Pos, Me.ScrollMax(HORZ))
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, HORZ, c.Info(HORZ), API_TRUE)
    RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
    
    'Raise a mouse move event since the pointer is no longer under the part of the _
     image it was before and the controller might need the new ImageX/Y values
    Dim MousePos As blu.POINT
    Let MousePos = blu.GetMousePos_Window(UserControl.hWnd)
    'TODO: Get mouse button / key state
    RaiseEvent MouseMove( _
        0, 0, MousePos.X, MousePos.Y, GetImageX(MousePos.X), GetImageY(MousePos.Y) _
    )
    
    'The viewport is refreshed _after_ the events fire so that your controller _
     does *not* have to call `Refresh` itself, saving repaints
    Call Me.Refresh
End Property

'PROPERTY ScrollY : Scroll the viewport to a specific place vertically _
 ======================================================================================
Public Property Get ScrollY() As Long: Let ScrollY = c.Info(VERT).Pos: End Property
Public Property Let ScrollY(ByVal Value As Long)
    'Scroll the viewport...
    With c.Info(VERT)
        Let .Mask = SIF_POS
        Let .Pos = blu.Max(.Pos, Me.ScrollMax(VERT))
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, VERT, c.Info(VERT), API_TRUE)
    RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
    
    'Raise a mouse move event since the pointer is no longer under the part of the _
     image it was before and the controller might need the new ImageX/Y values
    Dim MousePos As blu.POINT
    Let MousePos = blu.GetMousePos_Window(UserControl.hWnd)
    'TODO: Get mouse button / key state
    RaiseEvent MouseMove( _
        0, 0, MousePos.X, MousePos.Y, GetImageX(MousePos.X), GetImageY(MousePos.Y) _
    )
    
    'The viewport is refreshed _after_ the events fire so that your controller _
     does *not* have to call `Refresh` itself, saving repaints
    Call Me.Refresh
End Property

'PROPERTY Zoom _
 ======================================================================================
Public Property Get Zoom() As Long: Let Zoom = My_Zoom: End Property
Public Property Let Zoom(ByVal ZoomLevel As Long)
    'Keep within the defined bounds
    Let ZoomLevel = blu.Range(ZoomLevel, My_ZoomMax, My_ZoomMin)
    'Don't change the zoom if not necessary (avoid repaints)
    If My_Zoom = ZoomLevel Then Exit Property Else Let My_Zoom = ZoomLevel
    
    'Recalculate the scroll bar limits, when we send events they may want to refer to _
     the min / max / centre values and so forth
    Call InitScrollBars
    
    'Raise a mouse move event since the pointer is no longer under the part of the _
     image it was before and the controller might need the new ImageX/Y values
    Dim MousePos As blu.POINT
    Let MousePos = blu.GetMousePos_Window(UserControl.hWnd)
    'TODO: Get mouse button / key state
    RaiseEvent MouseMove( _
        0, 0, MousePos.X, MousePos.Y, GetImageX(MousePos.X), GetImageY(MousePos.Y) _
    )
    
    'The viewport is refreshed _after_ the events fire so that your controller _
     does *not* have to call `Refresh` itself, saving repaints
    Call Me.Refresh
End Property

'PROPERTY ZoomMin : Minimum zoom size (e.g. when Ctrl+Scroll zooming) _
 ======================================================================================
Public Property Get ZoomMin() As Long: Let ZoomMin = My_ZoomMin: End Property
Public Property Let ZoomMin(ByVal ZoomLevel As Long)
    'Let's not divide by zero!
    If ZoomLevel < 1 Then Let ZoomLevel = 1
    'The minimum cannot be greater than the maximum
    If ZoomLevel > My_ZoomMax Then Let ZoomLevel = My_ZoomMax
    
    'Set the property value now as it will assessed if the zoom is changed below
    Let My_ZoomMin = ZoomLevel
    'If the current zoom is less than that, change the zoom level
    If My_Zoom < My_ZoomMin Then Let Me.Zoom = My_ZoomMin
    
    Call UserControl.PropertyChanged("ZoomMin")
End Property

'PROPERTY ZoomMax : Maximum zoom size (e.g. when Ctrl+Scroll zooming) _
 ======================================================================================
Public Property Get ZoomMax() As Long: Let ZoomMax = My_ZoomMax: End Property
Public Property Let ZoomMax(ByVal ZoomLevel As Long)
    'Let's not divide by zero!
    If ZoomLevel < 1 Then Let ZoomLevel = 1
    'Zoom max cannot be less than zoom min!
    If ZoomLevel < My_ZoomMin Then Let ZoomLevel = My_ZoomMin
    
    'Set the property value now as it will assessed if the zoom is changed below
    Let My_ZoomMax = ZoomLevel
    'If the current zoom is greater than that, change the zoom level
    If My_Zoom > My_ZoomMax Then Let Me.Zoom = My_ZoomMax
    
    Call UserControl.PropertyChanged("ZoomMin")
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'AddLayer : Adds a layer to the whole image _
 ======================================================================================
Public Function AddLayer( _
    Optional ByVal TransparentColour As OLE_COLOR = &H123456 _
) As Long
    'Add an extra layer item to the array
    ReDim Preserve Layers(NumberOfLayers) As Layer
    'Set the image size for the layer
    With Layers(NumberOfLayers)
        Set .IMAGE = New bluImage
        Call .IMAGE.Create24Bit( _
            ImageWidth:=c.ImageRECT.Right, ImageHeight:=c.ImageRECT.Bottom, _
            BackgroundColour:=TransparentColour, UseTransparency:=True _
        )
    End With
    
    Let NumberOfLayers = NumberOfLayers + 1
    Let AddLayer = NumberOfLayers
    
    'NOTE: This procedure does not refresh the viewport! It is expected that you might _
     want to add multiple layers and paint into them before refreshing the viewport
End Function

'Cls : Clears the image _
 ======================================================================================
Public Sub Cls(Optional ByVal Layer As Long = -1)
    'If no layer is specified ("-1"), clear all of them
    Dim i As Long
    For i = _
        IIf(Layer = -1, LBound(Layers), Layer) To _
        IIf(Layer = -1, UBound(Layers), Layer)
        'Paint the layer clear
        Call Layers(i).IMAGE.Cls
    Next
    
    'NOTE: This procedure does not refresh the viewport! When you clear the image _
     (or layer), you might be beginning to paint on it and won't want flicker
End Sub

'Refresh _
 ======================================================================================
Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
    'Queue a `WM_PAINT` message to repaint the whole viewport area
    Call blu.user32_InvalidateRect(UserControl.hWnd, c.ClientRECT, API_FALSE)
End Sub

'ScrollTo : Scroll to an X and Y location in one call _
 ======================================================================================
Public Sub ScrollTo(ByVal X As Long, ByVal Y As Long)
    'Scroll X
    With c.Info(HORZ)
        Let .Mask = SIF_POS
        Let .Pos = blu.Range(X, Me.ScrollMax(HORZ), 0)
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, HORZ, c.Info(HORZ), API_TRUE)
    
    'Scroll Y
    With c.Info(VERT)
        Let .Mask = SIF_POS
        Let .Pos = blu.Range(Y, Me.ScrollMax(VERT), 0)
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, VERT, c.Info(VERT), API_TRUE)
    
    'Alert the controller to the scroll happening
    RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
    
    'Raise a mouse move event since the pointer is no longer under the part of the _
     image it was before and the controller might need the new ImageX/Y values
    Dim MousePos As blu.POINT
    Let MousePos = blu.GetMousePos_Window(UserControl.hWnd)
    'TODO: Get mouse button / key state
    RaiseEvent MouseMove( _
        0, 0, MousePos.X, MousePos.Y, GetImageX(MousePos.X), GetImageY(MousePos.Y) _
    )
    
    'The viewport is refreshed _after_ the events fire so that your controller _
     does *not* have to call `Refresh` itself, saving repaints
    Call Me.Refresh
End Sub

'SetImageProperties : Set the size of the back buffer image to scroll around _
 ======================================================================================
Public Sub SetImageProperties( _
             ByRef Width As Long, ByRef Height As Long, _
    Optional ByRef BitDepth As bluImage_Depth = [24-Bit] _
)
    'Changing the image size will destroy all existing layers
    Erase Layers
    ReDim Layers(0) As Layer
    With Layers(0)
        Set .IMAGE = New bluImage
        If BitDepth = [8-Bit] Then
            Call .IMAGE.Create8Bit(Width, Height, , True)
        ElseIf BitDepth = [24-Bit] Then
            Call .IMAGE.Create24Bit(Width, Height, c.UserControl_BackColor, True)
        ElseIf BitDepth = [32-Bit] Then
            Call .IMAGE.Create32Bit(Width, Height, c.UserControl_BackColor, True)
        End If
    End With
    Let NumberOfLayers = 1
    
    'Cache details of the image for faster painting
    Call blu.user32_SetRect(c.ImageRECT, 0, 0, Width, Height)
    
    'When the back buffer changes, recalculate the scrollbars
    Call InitScrollBars
    
    'NOTE: This procedure does not refresh the viewport! It is expected that you might _
     want to add multiple layers and paint into them before refreshing the viewport
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'GetImageX : Given the mouse X position, return the X-image-pixel _
 ======================================================================================
Private Function GetImageX(ByVal X As Long) As Long
    Let GetImageX = c.Info(HORZ).Pos + (X - c.Centre.X) \ My_Zoom
End Function

'GetImageY : Given the mouse Y position, return the Y-image-pixel _
 ======================================================================================
Private Function GetImageY(ByVal Y As Long) As Long
    Let GetImageY = c.Info(VERT).Pos + (Y - c.Centre.Y) \ My_Zoom
End Function

'InitScrollBars _
 ======================================================================================
Private Sub InitScrollBars()
    'Show / Hide scrollbars? _
     ----------------------------------------------------------------------------------
    'The size of the image, accounting for zooming
    Dim ImageSize As SIZE
    Let ImageSize.Width = c.ImageRECT.Right * My_Zoom
    Let ImageSize.Height = c.ImageRECT.Bottom * My_Zoom
    
    'Get the current size of the viewport
    Call blu.user32_GetClientRect(UserControl.hWnd, c.ClientRECT)
    'If the image is wider than the viewport, show the horizontal scrollbar
    Call user32_ShowScrollBar( _
        UserControl.hWnd, HORZ, Abs(ImageSize.Width > c.ClientRECT.Right) _
    )
    'If the image were no longer to fit in the height of the viewport because of the _
     appearance of the horizontal scrollbar at the bottom, we need to account for _
     this and add a vertical scrollbar. Fetch new viewport dimensions before checking
    Call blu.user32_GetClientRect(UserControl.hWnd, c.ClientRECT)
    'If the image is taller than the viewport, show the vertical scrollbar
    Call user32_ShowScrollBar( _
        UserControl.hWnd, VERT, Abs(ImageSize.Height > c.ClientRECT.Bottom) _
    )
    'Again, this could cause a situation where the appearance of the veritcal _
     scrollbar means that the image no longer fits horizontally, get an updated _
     viewport size, ...
    Call blu.user32_GetClientRect(UserControl.hWnd, c.ClientRECT)
    '... and show the horizontal scrollbar if necessary
    Call user32_ShowScrollBar( _
        UserControl.hWnd, HORZ, Abs(ImageSize.Width > c.ClientRECT.Right) _
    )
    'Lastly, grab the final size of the viewport (sans the scrollbars)
    Call blu.user32_GetClientRect(UserControl.hWnd, c.ClientRECT)
    
    'Setup the back buffer: _
     ----------------------------------------------------------------------------------
    'Recreate the back buffer (the size of the viewport) for flicker-free painting
    Set Buffer = Nothing
    Set Buffer = New bluImage
    Call Buffer.Create24Bit( _
        c.ClientRECT.Right, c.ClientRECT.Bottom, _
        c.UserControl_BackColor _
    )
    
    'Set the background to be used at painting to the back buffer
    Call blu.gdi32_SetDCBrushColor( _
        Buffer.hDC, c.UserControl_BackColor _
    )
    
    'Calculate portion of image to be displayed: _
     ----------------------------------------------------------------------------------
    'If the image is narrower than than the viewport then centre it horizontally
    If My_Centre = True And ImageSize.Width < c.ClientRECT.Right Then
        'Offset the image by half the difference in space between the image and _
         the viewport's width so that it appears centred
        Let c.Centre.X = (c.ClientRECT.Right - ImageSize.Width) \ 2
        'We will be painting the full width of the image, nothing clipped
        Let c.Dst.Width = ImageSize.Width
    Else
        Let c.Centre.X = 0
        'When zoomed in, if the viewport is not an exact multiple of the zoom level _
         (i.e. odd numbered width when zoom=2) then the image stretching would cause _
         one or more of the pixels row/columns to be thicker than the others, that is, _
         if your viewport is 21px wide and the image is 10px wide then one of those _
         10 pixel columns will be 3px wide, not 2px. This could be a problem with the _
         controller who might expect a rigid, consistent grid when zoomed. _
         To fix this we have to normalise the destination width/height so that it is _
         a multiple of the zoom factor
        Let c.Dst.Width = _
            c.ClientRECT.Right + (My_Zoom - 1) - (c.ClientRECT.Right Mod My_Zoom)
    End If
        
    'If the image is shorter than the viewport then centre it vertically
    If My_Centre = True And ImageSize.Height < c.ClientRECT.Bottom Then
        'Offset the image by half the difference in space between the image and _
         the viewport's height so that it appears centred
        Let c.Centre.Y = (c.ClientRECT.Bottom - ImageSize.Height) \ 2
        'We will be painting the full height of the image, nothing clipped
        Let c.Dst.Height = ImageSize.Height
    Else
        Let c.Centre.Y = 0
        Let c.Dst.Height = _
            c.ClientRECT.Bottom + (My_Zoom - 1) - (c.ClientRECT.Bottom Mod My_Zoom)
    End If
    
    Let c.Src.Width = c.Dst.Width \ My_Zoom
    Let c.Src.Height = c.Dst.Height \ My_Zoom
    
    'Recalculate the scroll bars limits: _
     ----------------------------------------------------------------------------------
    'Resizing the control might cause the scroll position to change if it's against _
     the ends of the scroll limit, we need to check this and send a Scroll event
    Dim OldHpos As Long, OldVPos As Long
    
    'Recalculate the scroll bars:
    With c.Info(HORZ)
        Let OldHpos = .Pos
        Let .Mask = SIF_PAGE Or SIF_RANGE Or SIF_POS
        Let .Page = c.ClientRECT.Right \ My_Zoom
        Let .Max = blu.Min(c.ImageRECT.Right)
        Let .Pos = blu.Range(.Pos, Me.ScrollMax(HORZ), .Min)
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, HORZ, c.Info(HORZ), API_TRUE)
    
    Call user32_ShowScrollBar( _
        UserControl.hWnd, HORZ, Abs(ImageSize.Width > c.ClientRECT.Right) _
    )
    
    With c.Info(VERT)
        Let OldVPos = .Pos
        Let .Mask = SIF_PAGE Or SIF_RANGE Or SIF_POS
        Let .Page = c.ClientRECT.Bottom \ My_Zoom
        Let .Max = blu.Min(c.ImageRECT.Bottom)
        Let .Pos = blu.Range(.Pos, Me.ScrollMax(VERT), .Min)
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, VERT, c.Info(VERT), API_TRUE)
    
    Call user32_ShowScrollBar( _
        UserControl.hWnd, VERT, Abs(ImageSize.Height > c.ClientRECT.Bottom) _
    )
    
    'Now send a scroll event if the scroll value changed
    If OldHpos <> c.Info(HORZ).Pos _
    Or OldVPos <> c.Info(VERT).Pos _
    Then
        RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
    End If
    
    'NOTE: This procedure does not refesh the viewport, the various callers handle _
     that so as to not do unecessary repaints
End Sub

'/// SUBCLASS /////////////////////////////////////////////////////////////////////////
'bluMagic helps us tap into the Windows message stream going on in the background _
 so that we can trap mouse / window events and a whole lot more. This works using _
 "function ordinals", therefore this procedure has to be the last one on the page

'SubclassWindowProcedure : THIS MUST BE THE LAST PROCEDURE ON THIS PAGE _
 ======================================================================================
Private Sub SubclassWindowProcedure( _
    ByVal Before As Boolean, _
    ByRef Handled As Boolean, _
    ByRef ReturnValue As Long, _
    ByVal hndWindow As Long, _
    ByVal Message As WM, _
    ByVal wParam As Long, _
    ByVal lParam As Long, _
    ByRef UserParam As Long _
)
    Select Case Message
        '<msdn.microsoft.com/en-us/library/windows/desktop/dd145213%28v=vs.85%29.aspx>
        Case WM_PAINT '----------------------------------------------------------------
            If NumberOfLayers <> 0 Then
                'Prepare the surface for painting
                Dim Paint As PAINTSTRUCT
                Call user32_BeginPaint(UserControl.hWnd, Paint)
                             
                'Clear the image with the background colour
                Call blu.user32_FillRect( _
                    Buffer.hDC, c.ClientRECT, c.DC_BRUSH _
                )
                
                'Work downward through the layers:
                Dim i As Long
                For i = 0 To NumberOfLayers - 1
                    'The bottom layer does not need to be painted transparently
                    If i = 0 Then
                        'With no zoom, it's ever so slightly faster to non-stretch `BitBlt`
                        If My_Zoom = 1 Then
                            Call blu.gdi32_BitBlt( _
                                Buffer.hDC, _
                                c.Centre.X, c.Centre.Y, _
                                c.Dst.Width, c.Dst.Height, _
                                Layers(0).IMAGE.hDC, _
                                c.Info(HORZ).Pos, c.Info(VERT).Pos, _
                                vbSrcCopy _
                            )
                        Else
                            'When zoomed, stretch the 1:1 image to the viewport. The source _
                             and destination sizes are calculated in `InitScrollBars`, usually _
                             called upon resizing the viewport
                            Call blu.gdi32_StretchBlt( _
                                Buffer.hDC, _
                                c.Centre.X, c.Centre.Y, _
                                c.Dst.Width, c.Dst.Height, _
                                Layers(0).IMAGE.hDC, _
                                c.Info(HORZ).Pos, c.Info(VERT).Pos, _
                                c.Src.Width, c.Src.Height, _
                                vbSrcCopy _
                            )
                        End If
                    Else
                        'For the other layers, mask out their background colour
                        Call blu.gdi32_GdiTransparentBlt( _
                            Buffer.hDC, _
                            c.Centre.X, c.Centre.Y, _
                            c.Dst.Width, c.Dst.Height, _
                            Layers(i).IMAGE.hDC, _
                            c.Info(HORZ).Pos, c.Info(VERT).Pos, _
                            c.Src.Width, c.Src.Height, _
                            Layers(i).IMAGE.BackgroundColour _
                        )
                    End If
                Next
                
                'Give the controller the opportunity to paint over the final display
                RaiseEvent Paint(Buffer.hDC)
                
                'Copy the back buffer onto the display
                Call blu.gdi32_BitBlt( _
                    Paint.hndDC, 0, 0, c.ClientRECT.Right, c.ClientRECT.Bottom, _
                    Buffer.hDC, 0, 0, vbSrcCopy _
                )
                
                'Finish painting, let Windows know we've handled it ourselves
                Call user32_EndPaint(UserControl.hWnd, Paint)
                Let ReturnValue = 0
                Let Handled = True
            End If
        
        '`WM_ERASEBKGND` _
         <msdn.microsoft.com/en-us/library/windows/desktop/ms648055%28v=vs.85%29.aspx>
        Case WM_ERASEBKGND '-----------------------------------------------------------
            'Don't paint the background so as to avoid flicker, _
             all painting will be done in `WM_PAINT`
            Let ReturnValue = 1
            Let Handled = True
        
        '`WM_HSCROLL` and `WM_VSCROLL` - the scroll bars have been clicked _
         <msdn.microsoft.com/en-us/library/windows/desktop/bb787575%28v=vs.85%29.aspx> _
         <msdn.microsoft.com/en-us/library/windows/desktop/bb787577%28v=vs.85%29.aspx>
        Case WM.WM_HSCROLL, WM.WM_VSCROLL '--------------------------------------------
            'Which scroll bar?
            Dim Bar As bluScrollBar
            If Message = WM_HSCROLL Then Let Bar = HORZ Else Let Bar = VERT
            
            With c.Info(Bar)
                'Record the current position so we can know how far we're scrolling
                Dim ScrollBy(0 To 1) As Long
                Let ScrollBy(Bar) = .Pos
                'Prepare to update the scroll value
                Let .Mask = SIF_POS
                'What part of the bar has been clicked?
                Select Case wParam And &HFFFF&
                    'The user is dragging the scroll bar; `lParam` contains the value, _
                     but only up to 16 bits, we get the full value with `GetScrollInfo`
                    Case SB.SB_THUMBTRACK
                        'Fetch the current `TrackPos` value
                        Let .Mask = .Mask Or SIF_TRACKPOS
                        Call user32_GetScrollInfo(UserControl.hWnd, Bar, c.Info(Bar))
                        'Move the scroll bar to this value
                        Let .Pos = .TrackPos
                    
                    'Home -- jump right to the beginning
                    Case SB.SB_LEFT, SB.SB_TOP
                        Let .Pos = .Min
                    
                    'End -- jump right to the end
                    Case SB.SB_RIGHT, SB.SB_BOTTOM
                        Let .Pos = Me.ScrollMax(Bar)
                    
                    'Left
                    Case SB.SB_LINELEFT, SB.SB_LINEUP
                        Let .Pos = .Pos - My_ScrollAmount(Bar)
                    
                    'Right
                    Case SB.SB_LINERIGHT, SB.SB_LINEDOWN
                        Let .Pos = .Pos + My_ScrollAmount(Bar)
                    
                    'Page left / up
                    Case SB.SB_PAGELEFT, SB.SB_PAGEUP
                        Let .Pos = .Pos - .Page
                        
                    'Page right / down
                    Case SB.SB_PAGERIGHT, SB.SB_PAGEDOWN
                        Let .Pos = .Pos + .Page
                    
                    'Any other kind of interaction doesn't change the position
                    Case Else
                        Let ReturnValue = 0
                        Let Handled = True
                        Exit Sub
                        
                End Select
                'Make sure the new value isn't out of range
                Let .Pos = blu.Range(.Pos, Me.ScrollMax(Bar), .Min)
                
                'Convert the old position to a relative value (+/-...)
                Let ScrollBy(Bar) = ScrollBy(Bar) - .Pos
            End With
            
            'Update the scroll bar
            Call user32_SetScrollInfo(UserControl.hWnd, Bar, c.Info(Bar), API_TRUE)
            
            'Alert the owner of the move
            RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
            
            'Raise a mouse move event since the pointer is no longer under the part _
             of the image it was before and the controller might need the new _
             ImageX/Y values
            Dim MousePos As blu.POINT
            Let MousePos = blu.GetMousePos_Window(UserControl.hWnd)
            'When the viewport is scrolled by keyboard or by the controller _
            (`ScrollTo`), then we want to fire a `MouseMove` event to say that the _
            mouse pointer is under a different part of the image than before _
            (`ImageX/Y`), but the mouse position is not always immediately available _
            to us in that event (e.g. keyboard scrolling)
            'TODO: Get mouse button / key state
            RaiseEvent MouseMove( _
                0, 0, MousePos.X, MousePos.Y, GetImageX(MousePos.X), GetImageY(MousePos.Y) _
            )
            
            'Scroll the pixels in the window
            Call user32_ScrollWindowEx( _
                UserControl.hWnd, ScrollBy(HORZ) * My_Zoom, ScrollBy(VERT) * My_Zoom, _
                0, 0, 0, 0, SW_INVALIDATE _
            )
            'The viewport is refreshed _after_ the events fire so that your controller _
             does *not* have to call `Refresh` itself, saving repaints
            Call user32_UpdateWindow(UserControl.hWnd)
            
            Let ReturnValue = 0
            Let Handled = True
    
        'The scroll bars have been clicked -- set focus to the control
        Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMBUTTONDOWN
            'The user control does not automatically gain focus when you click on the _
             scrollbars (you might expect to click on the scrollbar and then use the _
             keys to move it)
            Call UserControl.SetFocus
    End Select

'======================================================================================
'    C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
'--------------------------------------------------------------------------------------
'           DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'======================================================================================
End Sub
