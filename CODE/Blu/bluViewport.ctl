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
End
Attribute VB_Name = "bluViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
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
'Dependencies       bluImage.cls, bluMagic.cls, bluMouseEvents.cls, Lib.bas, WIN32.bas
'Last Updated       12-AUG-13

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

Private Enum WM
    WM_PAINT = &HF
    WM_ERASEBKGND = &H14
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
End Enum

'Send a window message _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms644950%28v=vs.85%29.aspx>
Private Declare Function user32_SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hndWindow As Long, _
    ByVal Message As WM, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

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
    ByVal Flags As SW _
) As Long

Private Enum SW
    SW_INVALIDATE = &H2
End Enum

'<msdn.microsoft.com/en-us/library/windows/desktop/dd145002%28v=vs.85%29.aspx>
Private Declare Function user32_InvalidateRect Lib "user32" Alias "InvalidateRect" ( _
    ByVal hndWindow As Long, _
    ByRef InvalidRECT As RECT, _
    ByVal EraseBG As BOOL _
) As BOOL

'<msdn.microsoft.com/en-us/library/windows/desktop/dd145167%28v=vs.85%29.aspx>
Private Declare Function user32_UpdateWindow Lib "user32" Alias "UpdateWindow" ( _
    ByVal hndWindow As Long _
) As BOOL

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'We'll need to subclass the control to listen into the scroll bar events
Dim Magic As bluMagic

'This will track mouse in / out and mouse wheel events
Dim WithEvents MouseEvents As bluMouseEvents
Attribute MouseEvents.VB_VarHelpID = -1

'bluViewport allows you to manage multiple layers for the whole image to minimise _
 the amount of painting you have to do and so bluViewport can manage the scrolling
Private Type Layer
    Image As bluImage
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
End Type
Dim c As CACHEVARS

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Public Enum bluScrollBar
    HORZ = 0
    VERT = 1
End Enum

Private My_ScrollAmount(0 To 1) As Long 'Amount to scroll clicking scroll arrow once
Private My_ScrollLineSize As Long       'Size of a "line" for mouse wheel scrolling
Private My_ScrollCharSize As Long       'Size of a "char" for horizontal mouse wheel

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
    ByVal X As Single, ByVal Y As Single _
)

'The standard mouse events we'll forward
Event Click()
Attribute Click.VB_UserMemId = -600
Event MouseDown(ByVal Button As VBRUN.MouseButtonConstants, ByVal Shift As VBRUN.ShiftConstants, ByVal X As Single, ByVal Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(ByVal Button As VBRUN.MouseButtonConstants, ByVal Shift As VBRUN.ShiftConstants, ByVal X As Single, ByVal Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(ByVal Button As VBRUN.MouseButtonConstants, ByVal Shift As VBRUN.ShiftConstants, ByVal X As Single, ByVal Y As Single)
Attribute MouseUp.VB_UserMemId = -607

'An event that occurs when painting, allowing you to alter the viewport's display
Event Paint(ByVal hDC As Long, ByVal ScrollX As Long, ByVal ScrollY As Long)

'When a scroll occurs
Event Scroll(ByVal ScrollX As Long, ByVal ScrollY As Long)

'CONTROL Click _
 ======================================================================================
Private Sub UserControl_Click(): RaiseEvent Click: End Sub

'CONTROL Initialize _
 ======================================================================================
Private Sub UserControl_Initialize()
    'The DC brush helps us avoid having to create and destroy a brush when we want _
     to paint. It acts as a built-in brush that we can set the colour on at will _
     <blogs.msdn.com/b/oldnewthing/archive/2005/04/20/410031.aspx>
    Let c.DC_BRUSH = WIN32.gdi32_GetStockObject(DC_BRUSH)
    
    'The `SCROLLINFO` structures must state their size, _
     let's do this just once
    Let c.Info(HORZ).SizeOfMe = LenB(c.Info(HORZ))
    Let c.Info(VERT).SizeOfMe = LenB(c.Info(VERT))
End Sub

'CONTROL InitProperties : New instance of the control plopped on a form _
 ======================================================================================
Private Sub UserControl_InitProperties()
    Let Me.ScrollAmountH = 32
    Let Me.ScrollAmountV = 32
    Let Me.ScrollLineSize = 16
    Let Me.ScrollCharSize = 16
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
        Call user32_SendMessage(UserControl.hWnd, Scroll, Send, 0)
    End If
End Sub

'CONTROL MouseDown _
 ======================================================================================
Private Sub UserControl_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'CONTROL MouseMove _
 ======================================================================================
Private Sub UserControl_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'CONTROL MouseUp _
 ======================================================================================
Private Sub UserControl_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Let Me.BackColor = .ReadProperty(Name:="BackColor", DefaultValue:=VBRUN.SystemColorConstants.vbApplicationWorkspace)
        Let Me.ScrollAmountH = .ReadProperty(Name:="ScrollAmountH", DefaultValue:=32)
        Let Me.ScrollAmountV = .ReadProperty(Name:="ScrollAmountV", DefaultValue:=32)
        Let Me.ScrollLineSize = .ReadProperty(Name:="ScrollLineSize", DefaultValue:=16)
        Let Me.ScrollCharSize = .ReadProperty(Name:="ScrollCharSize", DefaultValue:=16)
    End With
    
    'Only subclass if not in VB's design mode
    If blu.UserMode = True Then
        'Attach the mouse events to look for mouse enter / leave / wheel
        Set MouseEvents = New bluMouseEvents
        Call MouseEvents.Attach( _
            UserControl.hWnd, GetParentForm_hWnd(UserControl.Parent) _
        )
        'Subclass the control to listen to scroll bar events
        Set Magic = New bluMagic
        Call Magic.ssc_Subclass(UserControl.hWnd, 0, 1, Me)
        Call Magic.ssc_AddMsg( _
            UserControl.hWnd, MSG_BEFORE, _
            WM_PAINT, WM_ERASEBKGND, WM_HSCROLL, WM_VSCROLL _
        )
    End If
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize()
    'Reconfigure the scroll bar parameters, this will deal with the change in client _
     size as the showing / hiding of the scrollbars changes the client size
    Call InitScrollBars
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
            WM_PAINT, WM_ERASEBKGND, WM_HSCROLL, WM_VSCROLL _
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
        Call .WriteProperty(Name:="ScrollAmountH", Value:=My_ScrollAmount(HORZ), DefaultValue:=32)
        Call .WriteProperty(Name:="ScrollAmountV", Value:=My_ScrollAmount(VERT), DefaultValue:=32)
        Call .WriteProperty(Name:="ScrollLineSize", Value:=My_ScrollLineSize, DefaultValue:=16)
        Call .WriteProperty(Name:="ScrollCharSize", Value:=My_ScrollCharSize, DefaultValue:=16)
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
    RaiseEvent MouseHover(Button, Shift, X, Y)
End Sub

'EVENT MouseEvents MOUSEHSCROLL _
 ======================================================================================
Private Sub MouseEvents_MouseHScroll(ByVal CharsScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single)
    With c.Info(HORZ)
        Let .Mask = SIF_POS
        Let .Pos = Lib.Range( _
            .Pos - (CharsScrolled * My_ScrollCharSize), _
            Me.ScrollMax(HORZ), .Min _
        )
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, HORZ, c.Info(HORZ), API_TRUE)
    Call Me.Refresh
    RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
End Sub

'EVENT MouseEvents MOUSEVSCROLL _
 ======================================================================================
Private Sub MouseEvents_MouseVScroll(ByVal LinesScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single)
    With c.Info(VERT)
        Let .Mask = SIF_POS
        Let .Pos = Lib.Range( _
            .Pos - (LinesScrolled * My_ScrollLineSize), _
            Me.ScrollMax(VERT), .Min _
        )
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, VERT, c.Info(VERT), API_TRUE)
    Call Me.Refresh
    RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
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
    Let c.UserControl_BackColor = WIN32.OLETranslateColor(Color)
    'Apply the colour to the back buffer DC, this will automatically be used at paint
    If Not Buffer Is Nothing Then
        Call WIN32.gdi32_SetDCBrushColor( _
            Buffer.hDC, c.UserControl_BackColor _
        )
    End If
    Call Me.Refresh
    Call UserControl.PropertyChanged("BackColor")
End Property

'PROPERTY hDC : Handle to the device context for the image layer - not the control _
 ======================================================================================
Public Property Get hDC( _
    Optional ByVal Layer As Long = 0 _
) As Long
    If NumberOfLayers <> 0 And Layer < NumberOfLayers Then
        Let hDC = Layers(Layer).Image.hDC
    End If
End Property

'PROPERTY ScrollMax : Return the maximum scroll value _
 ======================================================================================
'You can't set this value as it is automatically managed by the viewport based on the _
 image size and viewport size
Public Property Get ScrollMax(ByVal Bar As bluScrollBar) As Long
Attribute ScrollMax.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Let ScrollMax = Lib.Min( _
        c.Info(Bar).Max - c.Info(Bar).Page - 1 _
    )
End Property

'PROPERTY ScrollAmountH : The amount to scroll when clicking the scroll arrows once _
 ======================================================================================
Public Property Get ScrollAmountH() As Long
Attribute ScrollAmountH.VB_Description = "The amount to scroll (horizontally) clicking the scroll arrow once"
Attribute ScrollAmountH.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Let ScrollAmountH = My_ScrollAmount(HORZ)
End Property
Public Property Let ScrollAmountH(ByVal Value As Long)
    Let My_ScrollAmount(HORZ) = Value
    Call UserControl.PropertyChanged("ScrollAmountH")
End Property

'PROPERTY ScrollAmountV : The amount to scroll when clicking the scroll arrows once _
 ======================================================================================
Public Property Get ScrollAmountV() As Long
Attribute ScrollAmountV.VB_Description = "The amount to scroll (vertically) clicking the scroll arrow once"
Attribute ScrollAmountV.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Let ScrollAmountV = My_ScrollAmount(VERT)
End Property
Public Property Let ScrollAmountV(ByVal Value As Long)
    Let My_ScrollAmount(VERT) = Value
    Call UserControl.PropertyChanged("ScrollAmountV")
End Property

'PROPERTY ScrollLineSize : The size of a "line" for mouse wheel scrolling _
 ======================================================================================
Public Property Get ScrollLineSize() As Long
Attribute ScrollLineSize.VB_Description = "The size (in px) of a ""line"" for vertical mouse wheel scrolling."
Attribute ScrollLineSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Let ScrollLineSize = My_ScrollLineSize
End Property
Public Property Let ScrollLineSize(ByVal Value As Long)
    Let My_ScrollLineSize = Value
    Call UserControl.PropertyChanged("ScrollLineSize")
End Property

'PROPERTY ScrollCharSize : The size of a "char" for horizontal wheel scrolling _
 ======================================================================================
Public Property Get ScrollCharSize() As Long
Attribute ScrollCharSize.VB_Description = "The size (in px) of a ""char"" for horizontal mouse wheel scrolling."
Attribute ScrollCharSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Let ScrollCharSize = My_ScrollCharSize
End Property
Public Property Let ScrollCharSize(ByVal Value As Long)
    Let My_ScrollCharSize = Value
    Call UserControl.PropertyChanged("ScrollCharSize")
End Property

'PROPERTY ScrollX : Scroll the viewport to a specific place horizontally _
 ======================================================================================
Public Property Get ScrollX() As Long: Let ScrollX = c.Info(HORZ).Pos: End Property
Public Property Let ScrollX(ByVal Value As Long)
    With c.Info(HORZ)
        Let .Mask = SIF_POS
        Let .Pos = Lib.Max(.Pos, Me.ScrollMax(HORZ))
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, HORZ, c.Info(HORZ), API_TRUE)
    Call Me.Refresh
    RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
End Property

'PROPERTY ScrollY : Scroll the viewport to a specific place vertically _
 ======================================================================================
Public Property Get ScrollY() As Long: Let ScrollY = c.Info(VERT).Pos: End Property
Public Property Let ScrollY(ByVal Value As Long)
    With c.Info(VERT)
        Let .Mask = SIF_POS
        Let .Pos = Lib.Max(.Pos, Me.ScrollMax(VERT))
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, VERT, c.Info(VERT), API_TRUE)
    Call Me.Refresh
    RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
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
        Set .Image = New bluImage
        Call .Image.Create24Bit( _
            ImageWidth:=c.ImageRECT.Right, ImageHeight:=c.ImageRECT.Bottom, _
            BackgroundColour:=TransparentColour, UseTransparency:=True _
        )
    End With
    
    Let NumberOfLayers = NumberOfLayers + 1
    Let AddLayer = NumberOfLayers
End Function

'Cls : Clears the image -- NOT the viewport _
 ======================================================================================
Public Sub Cls(Optional ByVal Layer As Long = -1)
    'If no layer is specified ("-1"), clear all of them
    Dim i As Long
    For i = _
        IIf(Layer = -1, LBound(Layers), Layer) To _
        IIf(Layer = -1, UBound(Layers), Layer)
        'Paint the layer clear
        Call Layers(i).Image.Cls
    Next i
    Call Me.Refresh
End Sub

'Refresh _
 ======================================================================================
Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
    'Queue a `WM_PAINT` message to repaint the whole viewport area
    Call user32_InvalidateRect(UserControl.hWnd, c.ClientRECT, API_FALSE)
End Sub

'ScrollTo : Scroll to an X and Y location in one call _
 ======================================================================================
Public Sub ScrollTo(ByVal X As Long, ByVal Y As Long)
    'Scroll X
    With c.Info(HORZ)
        Let .Mask = SIF_POS
        Let .Pos = Lib.Range(X, Me.ScrollMax(HORZ), 0)
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, HORZ, c.Info(HORZ), API_TRUE)
    
    'Scroll Y
    With c.Info(VERT)
        Let .Mask = SIF_POS
        Let .Pos = Lib.Range(Y, Me.ScrollMax(VERT), 0)
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, VERT, c.Info(VERT), API_TRUE)
    
    'Repaint to see the new location
    Call Me.Refresh
    
    RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
End Sub

'SetImageProperties : Set the size of the back buffer image to scroll around _
 ======================================================================================
Public Sub SetImageProperties( _
    ByVal Width As Long, ByVal Height As Long _
)
    'Changing the image size will destroy all existing layers
    Erase Layers
    ReDim Layers(0) As Layer
    With Layers(0)
        Set .Image = New bluImage
        Call .Image.Create24Bit(Width, Height, c.UserControl_BackColor)
    End With
    Let NumberOfLayers = 1
    
    'Cache details of the image for faster painting
    Call WIN32.user32_SetRect(c.ImageRECT, 0, 0, Width, Height)
    
    'When the back buffer changes, recalculate the scrollbars
    Call InitScrollBars
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'GetParentForm_hWnd : Recurses through the parent objects until we hit the top form _
 ======================================================================================
Private Function GetParentForm_hWnd(ByRef StartWith As Object) As Long
    Do
        On Error GoTo NowCheckMDI
        Set StartWith = StartWith.Parent
    Loop
NowCheckMDI:
    On Error GoTo Complete
    'There is no built in way to find the MDI parent of a child form, though of _
     course you can only have one MDI form in the app, but I wouldn't want to have to _
     reference that by name here, yours might be named something else. What we do is _
     use Win32 to go up through the "MDIClient" window (that isn't exposed to VB) _
     which acts as the viewport of the MDI form and then up again to hit the MDI form
    If StartWith.MDIChild = True Then
        Let GetParentForm_hWnd = _
            WIN32.user32_GetParent( _
            WIN32.user32_GetParent(StartWith.hWnd) _
        )
        Exit Function
    End If
Complete:
    Let GetParentForm_hWnd = StartWith.hWnd
End Function

'InitScrollBars _
 ======================================================================================
Private Sub InitScrollBars()
    'Get the size of the viewport
    Call WIN32.user32_GetClientRect(UserControl.hWnd, c.ClientRECT)
    
    'Show or hide the scrollbars based on the size of the viewport
    Call user32_ShowScrollBar( _
        UserControl.hWnd, HORZ, Abs(c.ImageRECT.Right > c.ClientRECT.Right) _
    )
    Call user32_ShowScrollBar( _
        UserControl.hWnd, VERT, Abs(c.ImageRECT.Bottom > c.ClientRECT.Bottom) _
    )
    
    'If a scrollbar was visible and gets hidden, it changes the size of the viewport, _
     regrab the size and work from that now on
    Call WIN32.user32_GetClientRect(UserControl.hWnd, c.ClientRECT)
    
    'Recreate the back buffer (the size of the viewport) for flicker-free painting
    Set Buffer = Nothing
    Set Buffer = New bluImage
    Call Buffer.Create24Bit( _
        c.ClientRECT.Right, c.ClientRECT.Bottom, _
        c.UserControl_BackColor _
    )
    
    'Set the background to be used at painting to the back buffer
    Call WIN32.gdi32_SetDCBrushColor( _
        Buffer.hDC, c.UserControl_BackColor _
    )
    
    'If the image is narrower than than the viewport then centre it horizontally
    If c.ImageRECT.Right < c.ClientRECT.Right Then
        Let c.Centre.X = (c.ClientRECT.Right - c.ImageRECT.Right) \ 2
        Let c.Dst.Width = c.ImageRECT.Right
    Else
        Let c.Centre.X = 0
        Let c.Dst.Width = c.ClientRECT.Right
    End If
        
    'If the image is shorter than the viewport then centre it vertically
    If c.ImageRECT.Bottom < c.ClientRECT.Bottom Then
        Let c.Centre.Y = (c.ClientRECT.Bottom - c.ImageRECT.Bottom) \ 2
        Let c.Dst.Height = c.ImageRECT.Bottom
    Else
        Let c.Centre.Y = 0
        Let c.Dst.Height = c.ClientRECT.Bottom
    End If
    
    'Resizing the control might cause the scroll position to change if it's against _
     the ends of the scroll limit, we need to check this and send a Scroll event
    Dim OldHpos As Long, OldVPos As Long
    
    'Recalculate the scroll bars:
    With c.Info(HORZ)
        Let OldHpos = .Pos
        Let .Mask = SIF_PAGE Or SIF_RANGE Or SIF_POS
        Let .Page = c.ClientRECT.Right
        Let .Max = Lib.Min(c.ImageRECT.Right + .Page - c.ClientRECT.Right)
        Let .Pos = Lib.Range(.Pos, Me.ScrollMax(HORZ), .Min)
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, HORZ, c.Info(HORZ), API_TRUE)
    
    With c.Info(VERT)
        Let OldVPos = .Pos
        Let .Mask = SIF_PAGE Or SIF_RANGE Or SIF_POS
        Let .Page = c.ClientRECT.Bottom
        Let .Max = Lib.Min(c.ImageRECT.Bottom + .Page - c.ClientRECT.Bottom)
        Let .Pos = Lib.Range(.Pos, Me.ScrollMax(VERT), .Min)
    End With
    Call user32_SetScrollInfo(UserControl.hWnd, VERT, c.Info(VERT), API_TRUE)
    
    Call Me.Refresh
    
    'Now send a scroll event if the scroll value changed
    If OldHpos <> c.Info(HORZ).Pos _
    Or OldVPos <> c.Info(VERT).Pos _
    Then
        RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
    End If
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
                Call WIN32.user32_FillRect( _
                    Buffer.hDC, c.ClientRECT, c.DC_BRUSH _
                )
                
                'Work downward through the layers:
                Dim i As Long
                For i = 0 To NumberOfLayers - 1
                    'The bottom layer does not need to be painted transparently
                    If i = 0 Then
                        'Paint the visible portion of the image
                        Call WIN32.gdi32_BitBlt( _
                            Buffer.hDC, _
                            c.Centre.X, c.Centre.Y, _
                            c.Dst.Width, c.Dst.Height, _
                            Layers(0).Image.hDC, _
                            c.Info(HORZ).Pos, c.Info(VERT).Pos, _
                            vbSrcCopy _
                        )
                    Else
                        'For the other layers, mask out their background colour
                        If WIN32.gdi32_GdiTransparentBlt( _
                            Buffer.hDC, _
                            c.Centre.X, c.Centre.Y, _
                            c.Dst.Width, c.Dst.Height, _
                            Layers(i).Image.hDC, _
                            c.Info(HORZ).Pos, c.Info(VERT).Pos, _
                            c.Dst.Width, c.Dst.Height, _
                            Layers(i).Image.BackgroundColour _
                        ) = 0 Then Stop
                    End If
                Next i
                
                'Give the controller the opportunity to paint over the final display
                RaiseEvent Paint(Buffer.hDC, c.Info(HORZ).Pos, c.Info(VERT).Pos)
                
                'Copy the back buffer onto the display
                Call WIN32.gdi32_BitBlt( _
                    Paint.hndDC, 0, 0, c.ClientRECT.Right, c.ClientRECT.Bottom, _
                    Buffer.hDC, 0, 0, vbSrcCopy _
                )
                Call user32_EndPaint(UserControl.hWnd, Paint)
                
'                '"validates the update region"
'                Call Magic.ssc_CallOrigWndProc(hndWindow, Message, wParam, lParam)
                
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
                    
                    'Page left
                    Case SB.SB_PAGELEFT, SB.SB_PAGEUP
                        Let .Pos = .Pos - .Page
                        
                    Case SB.SB_PAGERIGHT, SB.SB_PAGEDOWN
                        Let .Pos = .Pos + .Page
                    
                    'Any other kind of interaction doesn't change the position
                    Case Else
                        Let ReturnValue = 0
                        Let Handled = True
                        Exit Sub
                        
                End Select
                'Make sure the new value isn't out of range
                Let .Pos = Lib.Range(.Pos, Me.ScrollMax(Bar), .Min)
                
                'Convert the old position to a relative value (+/-...)
                Let ScrollBy(Bar) = ScrollBy(Bar) - .Pos
            End With
            
            'Scroll the pixels in the window
            Call user32_ScrollWindowEx( _
                UserControl.hWnd, ScrollBy(HORZ), ScrollBy(VERT), _
                0, 0, 0, 0, SW_INVALIDATE _
            )
            'Send `WM_PAINT` to fill in the empty area
            Call user32_UpdateWindow(UserControl.hWnd)
            
            'Update the scroll bar
            Call user32_SetScrollInfo(UserControl.hWnd, Bar, c.Info(Bar), API_TRUE)
            'Alert the owner of the move
            RaiseEvent Scroll(c.Info(HORZ).Pos, c.Info(VERT).Pos)
            
            Let ReturnValue = 0
            Let Handled = True
    End Select

'======================================================================================
'    C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
'--------------------------------------------------------------------------------------
'           DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'======================================================================================
End Sub
