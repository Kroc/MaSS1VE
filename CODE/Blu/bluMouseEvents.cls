VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluMouseEvents"
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
'CLASS :: bluMouseEvents

'Subclass an object to track mouse enter / leave / wheel events. _
 Special thanks to Tanner Helland <tannerhelland.com> for direction and bug-fixes

'Status             Ready to use
'Dependencies       blu.bas, bluMagic.cls
'Last Updated       27-JUN-14
'Last Update        Fixed bug with negative mouse position thanks to Tanner Helland _
                    & PhotoDemon <photodemon.org>

'/// API DEFS /////////////////////////////////////////////////////////////////////////

'A mouse event _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms644968%28v=vs.85%29.aspx>
Private Type MOUSEHOOKSTRUCT
    PT As POINT
    hndWindow As Long
    HitTestCode As Long
    ExtraInfo As Long
End Type

'Informs the mouse tracking to use the system default for hover time-out
Private Const HOVER_DEFAULT As Long = &HFFFFFFFF

'Structure to call the mouse tracking API _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms645604%28v=vs.85%29.aspx>
Private Type TRACKMOUSEEVENT
    SizeOfMe        As Long         'Size of this structure in bytes
    Flags           As TME_FLAGS    'Various settings
    hndTrack        As Long         'Which window to track
    HoverTime       As Long         'Hover timeout
End Type

Private Enum TME_FLAGS
    TME_CANCEL = &H80000000
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_NONCLIENT = &H10&
    TME_QUERY = &H40000000
End Enum

'Tap into the mouse movement _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms646265%28v=vs.85%29.aspx>
Private Declare Function user32_TrackMouseEvent Lib "user32" Alias "TrackMouseEvent" ( _
    ByRef ptrTRACKMOUSEEVENT As TRACKMOUSEEVENT _
) As BOOL

'What virtual keys are pressed, i.e. during `WM_MOUSEHOVER` / `WM_MOUSEMOVE`
Private Enum MK
    MK_LBUTTON = &H1            'Left mouse button
    MK_RBUTTON = &H2            'Right mouse button
    MK_SHIFT = &H4              'SHIFT key
    MK_CONTROL = &H8            'CTRL key
    MK_MBUTTON = &H10           'Middle mouse button
    MK_XBUTTON1 = &H20          'PREV mouse button
    MK_XBUTTON2 = &H40          'NEXT mouse button
End Enum

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'Which mouse pointer to show when the mouse is in the control
Public MousePointer As IDC

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'This will handle the subclassing safely
Private Magic As bluMagic

'The control we are tracking
Private hndTarget As Long
'Optionally, the parent form of the target so as to track mouse wheel events without _
 the control having to be focused. We listen into wheel events across the whole form _
 and test to see if the mouse is within the target control
Private hndParentForm As Long

'This strucutre will be used to start the mouse tracking, we have to keep hold of it _
 as the mouse tracking stops after every event fires
Private TME As TRACKMOUSEEVENT
'Whether the mouse tracking is running
Private IsTracking As Boolean
'Whether the mouse has left the window; _
 we use this to ensure that the MouseIn event only fires once
Private IsOut As Boolean

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'When the mouse leaves the window
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
'Mouse wheel scroll
Event MouseVScroll( _
    ByVal LinesScrolled As Single, _
    ByVal Button As VBRUN.MouseButtonConstants, _
    ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single _
)
Event MouseHScroll( _
    ByVal CharsScrolled As Single, _
    ByVal Button As VBRUN.MouseButtonConstants, _
    ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single _
)
'Mouse X-buttons
Event MouseBackButtonClick( _
    ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single _
)
Event MouseForwardButtonClick( _
    ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single _
)

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Let IsTracking = False
    Let IsOut = False
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Reset state
    Let IsTracking = False
    Let IsOut = False
    'Clean up the subclassing object
    If Not Magic Is Nothing Then
        'If a parent form was specified to subclass, detach the mouse wheel _
         messages from that separately from the main target
        Call Magic.ssc_DelMsg( _
            IIf(hndParentForm = 0, hndTarget, hndParentForm), MSG_BEFORE, _
            WM_MOUSEWHEEL, WM_MOUSEHWHEEL _
        )
        'And unsubclass the parent form
        If hndParentForm <> 0 Then _
            Magic.ssc_UnSubclass (hndParentForm)
        'Unsubclass the main target
        Call Magic.ssc_DelMsg( _
            hndTarget, MSG_BEFORE, _
            WM_SETCURSOR, WM_MOUSEMOVE, WM_MOUSEHOVER, WM_MOUSELEAVE, WM_XBUTTONDOWN _
        )
        Call Magic.ssc_UnSubclass(hndTarget)
        Call Magic.ssc_Terminate
        Set Magic = Nothing
    End If
End Sub

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Attach _
 ======================================================================================
Public Sub Attach( _
    ByVal Target_hWnd As Long, _
    Optional ByVal ParentForm_hWnd As Long = 0 _
)
    'Must attach at least something
    If Target_hWnd = 0 Then
        MsgBox "You must specify a non-zero Target Window Handle (`hWnd`) for " & _
               "`bluMouseEvents.Attach`. Perhaps your code called an API and " & _
               "received an invalid handle which you passed onwards.", _
               vbOKOnly Or vbCritical, "Developer Error"
        Exit Sub
    End If
    
    'Clean up any current tracking
    Call Class_Terminate
    
    'Keep hold of the window handles
    Let hndTarget = Target_hWnd
    Let hndParentForm = ParentForm_hWnd
    
    'Start subclassing (see the subclass section at the bottom of the file)
    Set Magic = New bluMagic
    'First, subclass the main target
    If Magic.ssc_Subclass(hndTarget, , , Me) Then
        'Prepare the mouse tracking API, we want to receive `WM_MOUSEHOVER` and _
         `WM_MOUSELEAVE` messages
        With TME
            Let .SizeOfMe = Len(TME)
            Let .Flags = TME_HOVER Or TME_LEAVE
            Let .HoverTime = HOVER_DEFAULT
            Let .hndTrack = hndTarget
        End With
        'Start the tracking. When the hover or mouse leave events occur the window _
         will receive the window messages which we handle in the subclass procedure
        If user32_TrackMouseEvent(TME) = API_TRUE Then
            Let IsTracking = True
            Let IsOut = False
            'Listen into the mouse events
            Call Magic.ssc_AddMsg( _
                hndTarget, MSG_BEFORE, _
                WM_SETCURSOR, WM_MOUSEMOVE, WM_MOUSEHOVER, WM_MOUSELEAVE, _
                WM_XBUTTONDOWN _
            )
        End If
        
        'If a parent form is given, subclass it. The mouse wheel events require the _
         window to be focused, so it's best to subclass the parent form which will _
         always be focused regardless of the other control's focus
        If hndParentForm <> 0 Then Call Magic.ssc_Subclass(hndParentForm, , , Me)
        'Add the mouse wheel events to the subclassing, either to the main target _
         or the parent form (if given)
        Call Magic.ssc_AddMsg( _
            IIf(hndParentForm = 0, hndTarget, hndParentForm), MSG_BEFORE, _
            WM_MOUSEWHEEL, WM_MOUSEHWHEEL _
        )
    End If
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'GetMouseStateFromParam : Extract the mouse / key state from a message param _
 ======================================================================================
Private Sub GetMouseStateFromParam( _
    ByVal wParam As Long, ByVal lParam As Long, _
    ByRef Button As VBRUN.MouseButtonConstants, _
    ByRef Shift As VBRUN.ShiftConstants, _
    ByRef X As Single, ByRef Y As Single _
)
    'TODO: Alt mask?
    If (wParam And MK.MK_LBUTTON) <> 0 Then Let Button = Button Or vbLeftButton
    If (wParam And MK.MK_MBUTTON) <> 0 Then Let Button = Button Or vbMiddleButton
    If (wParam And MK.MK_RBUTTON) <> 0 Then Let Button = Button Or vbRightButton
    If (wParam And MK.MK_CONTROL) <> 0 Then Let Shift = Shift Or vbCtrlMask
    If (wParam And MK.MK_SHIFT) <> 0 Then Let Shift = Shift Or vbShiftMask
    
    'The X and Y mouse coords are in the Long value as upper and lower 16-bits
    Let X = blu.LoWord(lParam)
    Let Y = blu.HiWord(lParam)
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
    'For some message events we'll provide the mouse and key state _
     (mouse button, ctrl/shift key state and mouse X / Y)
    Dim Button As VBRUN.MouseButtonConstants
    Dim Shift As VBRUN.ShiftConstants
    Dim X As Single, Y As Single
    
    '`WM_MOUSEWHEEL`: _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645617%28v=vs.85%29.aspx>
    '`WM_MOUSEHWHEEL`: _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645614%28v=vs.85%29.aspx>
    If Message = WM_MOUSEWHEEL _
    Or Message = WM_MOUSEHWHEEL Then '-------------------------------------------------
        'Since mouse wheel events are usually only sent to the window in focus, and _
         we don't want the user to have to click the control before being able to _
         scroll, the parent form is (optionally) subclassed so that we can hear _
         mouse wheel events all the time and we only need check if the mouse is _
         within the control
        
        'Get the location of the control in the screen _
         (including the non-client area, such as the scrollbars)
        Dim WindowRECT As RECT
        Call user32_GetWindowRect(hndTarget, WindowRECT)
        'Determine which mouse button / meta keys are pressed
        Call GetMouseStateFromParam(wParam, lParam, Button, Shift, X, Y)
        If user32_PtInRect(WindowRECT, X, Y) = API_FALSE Then Exit Sub
        
        'The mouse co-ords are relative to the screen, so we need to transform these _
         to be local to the target window
        Dim ClientPOINT As POINT
        Let ClientPOINT.X = X
        Let ClientPOINT.Y = Y
        Call user32_ScreenToClient(hndTarget, ClientPOINT)
        
        'The upper bytes of the `wParam` contain the amount the wheel scrolled, _
         which will be some fraction or multiple of 120 _
         <blogs.msdn.com/b/oldnewthing/archive/2013/01/23/10387366.aspx>
        Dim Delta As Long
        Let Delta = wParam \ &H10000
        
        'The horizontal wheel scroll gives a delta value opposite to vertical _
         scrolling (that is, right is positive, whilst up (not down) is positive. _
         We need to normalise this to ensure that real horizontal scroll doesn't act _
         differently from shift+vscroll
        If Message = WM_MOUSEHWHEEL Then Let Delta = -Delta
        
        'WARNING: Most code examples out there integer divide the delta (amount _
         scrolled) by 120 to determine the number of "detents" (notches) of the wheel _
         have been scrolled, but newer wheel mice have no notches and scroll smoothly _
         sending multiple `WM_MOUSEWHEEL` messages with smaller delta values, _
         therefore these examples break for absolute delta values that are not _
         multiples of 120. Instead, we will work on partial or multiple "lines" / _
         "chars" being scrolled. The height of a "line" / "char" is up to the _
         controller to convert into px
        
        'Horizontal or vertical scroll? _
         (holding shift converts a vertical scroll to horizontal)
        If Message = WM_MOUSEHWHEEL Or (Shift And vbShiftMask) > 0 Then
            Dim CharsScrolled As Single
            Let CharsScrolled = WheelScrollChars * (Delta / 120)
            'Alert the controller so it can do the scrolling
            'NOTE: You should *subtract* the `CharsScrolled` value from your scroll _
             position to scroll correctly
            RaiseEvent MouseHScroll( _
                CharsScrolled, Button, Shift, _
                ClientPOINT.X, ClientPOINT.Y _
            )
        Else
            Dim LinesScrolled As Single
            Let LinesScrolled = WheelScrollLines * (Delta / 120)
            'Alert the controller so it can do the scrolling
            'NOTE: Down vertical scrolling is negative, so you should *subtract* the _
             `LinesScrolled` value from your scroll position to scroll correctly
            RaiseEvent MouseVScroll( _
                LinesScrolled, Button, Shift, _
                ClientPOINT.X, ClientPOINT.Y _
            )
        End If
        
        Let ReturnValue = 0
        Let Handled = True
    
    '`WM_MOUSEHOVER` _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645613%28v=vs.85%29.aspx>
    ElseIf Message = WM_MOUSEHOVER Then '----------------------------------------------
        '`TrackMouseEvent` automatically cancels itself each time one of its messages _
         (`WM_MOUSELEAVE` / `WM_MOUSEHOVER`) are sent
        Let IsTracking = False
        
        'Determine which mouse button / meta keys are pressed
        Call GetMouseStateFromParam(wParam, lParam, Button, Shift, X, Y)
        RaiseEvent MouseHover(Button, Shift, X, Y)
        
        Let ReturnValue = 0
        Let Handled = True
        
    '`WM_MOUSELEAVE`: Does not fire until you've used `TrackMouseEvent` API _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645615%28v=vs.85%29.aspx>
    ElseIf Message = WM_MOUSELEAVE Then '----------------------------------------------
        '`TrackMouseEvent` automatically cancels itself each time one of its messages _
         (`WM_MOUSELEAVE` / `WM_MOUSEHOVER`) are sent
        RaiseEvent MouseOut
        Let IsTracking = False
        Let IsOut = True
        
        Let ReturnValue = 0
        Let Handled = True
    
    '`WM_MOUSEMOVE` _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645616%28v=vs.85%29.aspx>
    ElseIf Message = WM_MOUSEMOVE Then '-----------------------------------------------
        'Had the mouse left the window, the mouse tracking would have stopped
        If IsTracking = False Then
            Let IsTracking = True
            'Start the mouse tracking again to get the mouse leave message
            Call user32_TrackMouseEvent(TME)
        End If
        'There is no mouse-enter message, the first time a mouse move message is _
         receieve then the mouse has entered the window
        If IsOut = True Then
            Let IsOut = False
            RaiseEvent MouseIn
        End If
        
        'We don't mark the message as handled because we want the regular `MouseMove` _
         event to fire. With thanks to Tanner Helland for alerting me to this
    
    '`WM_XBUTTONDOWN` _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms646245%28v=vs.85%29.aspx>
    ElseIf Message = WM_XBUTTONDOWN Then '---------------------------------------------
        'NOTE: The correct way to handle the mouse X buttons is to process the _
         `WM_APPCOMMAND` message as this will be sent back / forward commands _
         automatically <msdn.microsoft.com/en-us/library/ms997498.aspx>, however _
         since the back / forward app commands (and a host of others) can be _
         triggered by the keyboard too, that's really the responsibility of a _
         global keyboard hook control and not a per-window tracker like this
        
        'Determine which meta keys are pressed, this doesn't determine _
         the mouse X buttons but fills in everything else we need
        Call GetMouseStateFromParam(wParam, lParam, Button, Shift, X, Y)
        
        'Now test which X button
        If (wParam And MK.MK_XBUTTON1) <> 0 Then
            RaiseEvent MouseBackButtonClick(Shift, X, Y)
                            
        ElseIf (wParam And MK.MK_XBUTTON2) <> 0 Then
            RaiseEvent MouseForwardButtonClick(Shift, X, Y)
            
        End If
        
        Let ReturnValue = 1
        Let Handled = True
    
    '`WM_SETCURSOR` _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms648382%28v=vs.85%29.aspx>
    ElseIf Message = WM_SETCURSOR Then '-----------------------------------------------
        If IsOut = False And MousePointer > 0 Then
            'If we set the cursor here, in `WM_SETCURSOR` it won't be overrided or _
             replaced by some other action in the message queue _
             <blogs.msdn.com/b/oldnewthing/archive/2005/05/25/421707.aspx>
            Call user32_SetCursor( _
                user32_LoadCursor(0, MousePointer) _
            )
            Let ReturnValue = 1
            Let Handled = True
        End If
        
    End If
    
'======================================================================================
'    C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
'--------------------------------------------------------------------------------------
'           DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'======================================================================================
End Sub
