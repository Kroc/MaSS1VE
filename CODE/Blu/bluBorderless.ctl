VERSION 5.00
Begin VB.UserControl bluBorderless 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "bluBorderless.ctx":0000
End
Attribute VB_Name = "bluBorderless"
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
'CONTROL :: bluBorderless

'This control does multiple, complex things for creating a modern metro-esque UI:

'1. Transforms the parent form into borderless Win8/Metro style, whilst still _
    preserving the system-provided window shadow. This is done using APIs available _
    on Vista and above for the Desktop Window Manager

'2. Provides a Minimise / Maximise / Close Window control box for the borderless form _
    that automatically adjusts to the state of the form

'On systems and settings where hardware accelerated DWM compositing is not available _
 bluBorderless disables itself and reverts to the standard window chrome. _
 It also does this in real time should the user disable the DWM by switching to a _
 non-accelerated theme

'Status             INCOMPLETE, DO NOT USE
'Dependencies       blu.bas
'Last Updated       08-MAY-14
'Last Update        Removed lib.bas depedency

'--------------------------------------------------------------------------------------

'A borderless form with system default shadow!? How is this possible!?

'We can't just set the Form's BorderStyle to None or there won't be any shadow and _
 I'm not a fan of drawing a shadow manually as this wastes resources, is hard to _
 manage and is a maintenance problem trying to keep abreast of new Windows versions.

'A form must have a non-client area (titlebar / border) in order to have a shadow. _
 We need to trick the form into having a least a tiny bit of non-client area which _
 can't be seen or blends into the form.

'Windows Vista and above provides an API (`DwmExtendFrameIntoClientArea`) that allows _
 you to make the window border extend _into_ the form area, that is, make the border _
 thicker on some or all sides, allowing you to place controls in the 'glass' area _
 (e.g. Internet Explorer has a transparent glass toolbar on Vista/7)

'Therefore if we extend the border into the form by 1 pixel (it has to be at least _
 one pixel), we can then remove the rest of the borders. We can't just change the _
 Form's BorderStyle at runtime, so instead we have to tap the WM_NCCALCSIZE message _
 which occurs when the window is calculating the size of the non-client (border) _
 area, where upon we can jump in there and feed it some zeroes.

'At this point with .NET you're golden, but there's a few more problems to overcome _
 with VB6. The GDI API that VB6 is based upon works entirely with 24-bit colours _
 (with the exception of the `AlphaBlend` function), meaning that the extra byte in _
 every 32-bit word is always zero (i.e. White = &H00FFFFFF). This is not good for us _
 as the Desktop Window Manager works with fully 32-bit colours and interprets the _
 upper zeroes as meaning transparent. Therefore our form, and any controls on it, _
 will appear discoloured in the glass area (which we've managed to limit to just _
 one pixel row).

'We can solve this this hard way or the easy way.

'The hard way involves subclassing the form and any controls that cover the 1px line _
 and tapping their WM_ERASEBKGND message where in we use `BeginBufferedPaint` and _
 `BufferedPaintSetAlpha` APIs to set a truly 32-bit background for the form/control. _
 Whilst this is doable, it's more complex, more prone to crashing and just doesn't _
 work for some types of controls.

'The easy way is to use `SetLayeredWindowAttributes` with a chroma key. The window _
 will gain proper transparency information with the only drawback that one particular _
 colour (of your choosing) will appear as a "hole" in your form wherever it appears.

'--------------------------------------------------------------------------------------

'Pieced together by Kroc Camen from endless searching and porting of .NET and C _
 samples, inlcuding:
 
':: Metro UI (Zune like) Interface (form) _
    <www.codeproject.com/Articles/138661/Metro-UI-Zune-like-Interface-form> _
    Demostrates the basic method in .NET

':: Windows Vista for Developers – Part 3 – The Desktop Window Manager _
    <weblogs.asp.net/kennykerr/archive/2006/08/10/Windows-Vista-for-Developers-_1320_-Part-3-_1320_-The-Desktop-Window-Manager.aspx> _
    Explains the Desktop Window Manager in detail, with the 24-bit control background _
    problem and how to solve it using layered window attributes

':: Controls and the Desktop Window Manager _
    <weblogs.asp.net/kennykerr/archive/2007/01/23/controls-and-the-desktop-window-manager.aspx> _
    The first blog post I found with the true answer to solving 24-bit control _
    backgrounds using the `BufferedPaint` APIs

'Modifying the window attributes was based upon "FormBdr" by Karl E. Peterson _
 <vb.mvps.org/samples/FormBdr/>

'TODO: Fold the mouse tracking APIs into this class? (Only need MouseOut)

'/// API CALLS ////////////////////////////////////////////////////////////////////////

'Window manipulation: _
 --------------------------------------------------------------------------------------

'Set the window state - show / hide / minimise / maximise &c. _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx>
Private Declare Function user32_ShowWindow Lib "user32" Alias "ShowWindow" ( _
    ByVal hndWindow As Long, _
    ByVal ShowCmd As SW _
) As BOOL

Private Enum SW
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
End Enum

'Position a window _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633545%28v=vs.85%29.aspx>
Private Declare Function user32_SetWindowPos Lib "user32" Alias "SetWindowPos" ( _
    ByVal hndWindow As Long, _
    ByVal hndInsertAfter As hWnd, _
    ByVal Left As Long, _
    ByVal Top As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal Flags As SWP _
) As BOOL

Private Enum hWnd
    HWND_TOP = 0                    'Move window to the top
    HWND_BOTTOM = 1                 'Move window to the bottom
    HWND_TOPMOST = -1               'Keep the window always on top
    HWND_NOTOPMOST = -2             'Undo the previous
End Enum

Private Enum SWP
    SWP_FRAMECHANGED = &H20         'Sends `WM_NCCALCSIZE` to calculate border area
    SWP_HIDEWINDOW = &H80           'Make invisible
    SWP_NOACTIVATE = &H10           'Don't focus the window
    SWP_NOCOPYBITS = &H100          'Don't paint the old contents into the new contents
    SWP_NOMOVE = &H2                'Don't change the Top / Left position
    SWP_NOREDRAW = &H8              'Don't repaint the window
    SWP_NOOWNERZORDER = &H200       'Don't change owner window's z-order
    SWP_NOSENDCHANGING = &H400      'Don't send `WM_WINDOWPOSCHANGING`
    SWP_NOSIZE = &H1                'Don't change window width / height
    SWP_NOZORDER = &H4              'Don't change window's z-order
    SWP_SHOWWINDOW = &H40           'Make visible
End Enum

'The System Menu is what you get when you: _
 1. Click the app-icon on the left of the title bar _
 2. Right-click the title bar, or _
 3. Press Alt+Spc
'<msdn.microsoft.com/en-us/library/windows/desktop/ms647985%28v=vs.85%29.aspx>
Private Declare Function user32_GetSystemMenu Lib "user32" Alias "GetSystemMenu" ( _
    ByVal hWnd As Long, _
    ByVal revert As Long _
) As Long

'The DWM APIs to expand the borders into the client area _
 --------------------------------------------------------------------------------------
'NOTE: Windows Vista+

'Whether Desktop Window Manager (DWM) composition is enabled _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa969518%28v=vs.85%29.aspx>
'TODO: Will need to listen for WM_DWMCOMPOSITIONCHANGED and handle changes
Private Declare Function dwmapi_DwmIsCompositionEnabled Lib "dwmapi" Alias "DwmIsCompositionEnabled" ( _
    ByRef Enabled As BOOL _
) As HRESULT

'Change the amount of non-client (border) area _
 <http://msdn.microsoft.com/en-us/library/windows/desktop/aa969512%28v=vs.85%29.aspx>
Private Declare Function dwmapi_DwmExtendFrameIntoClientArea Lib "dwmapi" Alias "DwmExtendFrameIntoClientArea" ( _
    ByVal hndWindow As Long, _
    ByRef Margin As MARGINS _
) As HRESULT

'Essentially a RECT. Used for expanding the window borders inwards _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb773244%28v=vs.85%29.aspx>
Private Type MARGINS
    Left As Long
    Right As Long
    Top As Long
    Bottom As Long
End Type

'Stuff to set the Layered Window attributes for fixing the form transparency _
 --------------------------------------------------------------------------------------

'Set the transparency information on the Layered Window _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633540%28v=vs.85%29.aspx>
Private Declare Function user32_SetLayeredWindowAttributes Lib "user32" Alias "SetLayeredWindowAttributes" ( _
    ByVal hndWindow As Long, _
    ByVal Colour As Long, _
    ByVal Alpha As Byte, _
    ByVal Flags As LWA _
) As BOOL

Private Enum LWA
    LWA_ALPHA = &H2                     'Set the transparency for the whole window
    LWA_COLORKEY = &H1                  'Make only a particular colour transparent
End Enum

'Subclassing Definitions _
 --------------------------------------------------------------------------------------

'Response codes for `WM_ACTIVATE`
Private Enum WA
    WA_ACTIVE = 1                       'Activated by some method other than mouse
                                         '(e.g. keyboard or `SetActiveWindow` function)
    WA_CLICKACTIVE = 2                  'Activated by a mouse click
    WA_INACTIVE = 0                     'Deactivated
End Enum

'`wParam` values for `WM_SYSCOMMAND` _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms646360(v=vs.85).aspx>
Private Enum SC
    SC_CLOSE = &HF060                   'Closes the window
    SC_DEFAULT = &HF160                 'System menu double-clicked
    SC_MAXIMIZE = &HF030                'Maximise window
    SC_MINIMIZE = &HF020                'Minimise window
    SC_MOVE = &HF010                    'Move the window
    SC_RESTORE = &HF120                 'Un-max/minimise the window
    SC_SIZE = &HF00                     'Resize the window
End Enum

'A `WM_NCCALCSIZE` message occurs when the window wants to calculate the amount of _
 non-client (border) area on the window. Whilst we don't need this to remove the _
 borders, we do need it to correct the window size when maximised _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms632606%28v=vs.85%29.aspx>
Private Type NCCALCSIZE_PARAMS
    Rectangles(0 To 2) As RECT
    ptrWINDOWPOS As Long                'Pointer to a `WINDOWPOS` structure _
                                         (not used in this class)
End Type

'Monitor info for min / max window size: _
 --------------------------------------------------------------------------------------

'Which monitor is a window mostly on? _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145064%28v=vs.85%29.aspx>
Private Declare Function user32_MonitorFromWindow Lib "user32" Alias "MonitorFromWindow" ( _
    ByVal hndWindow As Long, _
    ByVal Flags As MONITOR _
) As Long

Private Enum MONITOR
    MONITOR_DEFAULTTOPRIMARY = &H1      'Always get the primary monitor
    MONITOR_DEFAULTTONEAREST = &H2      'Get the monitor the window is mostly within
End Enum

'Retrieve details about a particular monitor _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd144901%28v=vs.85%29.aspx>
Private Declare Function user32_GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" ( _
    ByVal hndMonitor As Long, _
    ByRef Info As MONITORINFO _
) As BOOL

'A structure for details about a monitor, we're particularly interested in `WorkArea`, _
 this tells the area of the screen available to windows, that is, excluding the task _
 bar (whichever side it is on), docked toolbars and what not _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145065%28v=vs.85%29.aspx>
Private Type MONITORINFO
    SizeOfMe As Long
    MonitorArea As RECT
    WorkArea As RECT
    Flags As Long
End Type

'Define the minimum and maximum size of a window _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms632605%28v=vs.85%29.aspx>
Private Type MINMAXINFO
    Reserved As POINT
    MaxSize As POINT                    'Max width / height
    MaxPosition As POINT                'Top/Left of a maximised window
    MinTrackSize As POINT               'Minimum width / height
    MaxTrackSize As POINT               'Maximum size on the virtual screen?
End Type

'Non-Client Hit Testing: _
 --------------------------------------------------------------------------------------

'This defines the different types of zones in a window that Windows will handle itself _
 (if we choose not to override) _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms645618%28v=vs.85%29.aspx>
Private Enum HT
    HTTRANSPARENT = -1          'Click-through (in the same thread)
    HTNOWHERE = 0               'On screen background or dividing line between windows
    HTCLIENT = 1                'In the client area
    HTCAPTION = 2               'In the title bar
    HTSYSMENU = 3               'In a window menu or Close button in a child window
    HTSIZE = 4                  'In a sizing gripper (bottom-right corner sizing box)
    HTMENU = 5                  'In a menu
    HTHSCROLL = 6               'In a horizontal scroll bar
    HTVSCROLL = 7               'In the vertical scrollbar
    HTMINBUTTON = 8             'In a minimum button
    HTMAXBUTTON = 9             'In a maximize button
    HTLEFT = 10                 'In the left border of a resizable window
    HTRIGHT = 11                'In the right border of a resizable window
    HTTOP = 12                  'In the upper-horizontal border of a window
    HTTOPLEFT = 13              'In the upper-left corner of a window border
    HTTOPRIGHT = 14             'In the upper-right corner of a window border
    HTBOTTOM = 15               'In the lower-horizontal border of a resizable window
    HTBOTTOMLEFT = 16           'In the lower-left corner of a resizable border
    HTBOTTOMRIGHT = 17          'In the lower-right corner of a resizable border
    HTBORDER = 18               'Border of a window that does not have a sizing border
    HTCLOSE = 20                'In a close button
    HTHELP = 21                 'In a help button
End Enum

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'If our magic is working and the form has been made borderless (with shadow). Note _
 that this is regardless of if your form was already borderless (BorderStyle:None); _
 a borderless form will have a shadow added and this flag will be True. However, that _
 also means that if the magic is not working (On Windows XP / Aero is off / high _
 contrast theme) then this flag will be False even if your form was borderless to _
 begin with. Remember this flag is to tell you if bluBorderless has control of the _
 form borders, and not the border state of your form (use `Form.BorderStyle` for that)
Private My_IsBorderless As Boolean

'Which colour should be transparent on the form _
 (you'll want to set this to a colour that won't likely appear in your app)
Private My_ChromaKey As OLE_COLOR

'Button colours
Private My_ActiveColour As OLE_COLOR
Private My_BaseColour As OLE_COLOR
Private My_Style As bluSTYLE

'Minimum and maximum window size
Private My_MinWidth As Long
Private My_MinHeight As Long
Private My_MaxWidth As Long
Private My_MaxHeight As Long

'Keep the window always on top of other windows
Private My_AlwaysOnTop As Boolean

'/// PRIVATE DEFS /////////////////////////////////////////////////////////////////////

'Our subclassing object
Private Magic As bluMagic

'We'll use this to provide MouseIn/Out events
Private WithEvents MouseEvents As bluMouseEvents
Attribute MouseEvents.VB_VarHelpID = -1
'TODO: Replace this with proper `GetAsynckeyState` use
Private IsMouseDown As Boolean
Private WhichButton As bluBorderless_ButtonType

Private Enum bluBorderless_ButtonType
    Quit = 0
    Max = 1
    Min = 2
End Enum

'We need to refer to the parent form's handle a lot and unbound lookups are slow
Private ParentForm_hWnd As Long

'If the form was borderless to begin with. In order to give an already borderless _
 form a shadow we have to add a border before we run our subclassing to remove it. _
 We need to remember if the form was originally borderless so that if DWM switches _
 off whilst the program is running (Theme changed to Aero Basic / High Contrast) _
 we don't want to add the temporary border back on
Private WasBorderless As Boolean

'Here we'll store the thickness of the window borders before we remove them, _
 if they have to be added back on during runtime (i.e. DWM switches off), we need to _
 resize and reposition the form to account for the borders we removed earlier
Private Borders As RECT

'A list of window handles that are also subclassed, which will act as fake title bars _
 (for moving the form) or fake sizer boxes (for resizing the form)
Private NonClientHandlers As New Collection

'An enum for choosing the type of non client handler _
 (see the `RegisterNonClientHandler` procedure)
Private Enum NCH_TYPE
    MoveHandler = 1             'Act as a title bar (move the form)
    SizeHandler = 2             'Act as a sizing box in the bottom-right corner
End Enum

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Alert the parent form when the window compositing switches on / off, _
 they will want to re-juggle the form layout a bit
Event BorderlessStateChange(ByVal Enabled As Boolean)

'Notify the form of a true Activate event (i.e. when the window gains focus from _
 another app, unlike VB's Activate event which is only between VB's own windows)
Event Activate()
'And for the deactivate. In Windows 8 an inactive window has no shadow so the user _
 may need to do something to reduce the jarring effect of the shadowless form
Event Deactivate()

'CONTROL InitProperties : When a new instance of bluBorderless gets plopped on a form _
 ======================================================================================
Private Sub UserControl_InitProperties()
    'Get the handle to the parent form, even if the control is in a container
    Let ParentForm_hWnd = blu.GetParentForm_hWnd( _
        StartWith:=UserControl.Parent, GetMDIParent:=False _
    )
    'Initialise some defaults since `ReadProperties` is not called after this!
    Let Me.BaseColour = blu.BaseColour
    Let Me.ActiveColour = blu.ActiveColour
    Let Me.Style = Normal
    Let Me.ChromaKey = &H123456
End Sub

'CONTROL MouseDown _
 ======================================================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Remember when the mouse is down so that any calls to repaint keep the clicked _
     effect in place
    If IsMouseDown = False Then
        Let IsMouseDown = True
        Call Refresh
    End If
End Sub

'CONTROL MouseMove _
 =====================================================================================
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Refresh
End Sub

'CONTROL MouseUp _
 ======================================================================================
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'When letting go of the mouse, refresh, removing the click effect
    Let IsMouseDown = False
    Call Refresh
    
    'If you hold the mouse button down inside the control but release the button _
     outside then it doesn't count (allows you to escape from an accidental close)
    Dim ClientRECT As RECT
    Call blu.user32_GetClientRect(UserControl.hWnd, ClientRECT)
    If blu.user32_PtInRect(ClientRECT, X, Y) = API_FALSE Then Exit Sub
    
    'Only left button applies to action
    If Button <> VBRUN.MouseButtonConstants.vbLeftButton Then Exit Sub
    
    'Determine which button was clicked
    If WhichButton = Quit Then
        Unload blu.GetParentForm(UserControl.Parent, False)
    ElseIf WhichButton = Max Then
        'Flip between maximised / restore
        Call user32_ShowWindow( _
            ParentForm_hWnd, _
            IIf( _
                blu.user32_IsZoomed(ParentForm_hWnd) = API_TRUE, _
                SW_RESTORE, SW.SW_SHOWMAXIMIZED _
            ) _
        )
    ElseIf WhichButton = Min Then
        'Just minimise
        Call user32_ShowWindow(ParentForm_hWnd, SW_MINIMIZE)
    End If
End Sub

'CONTROL Paint _
 ======================================================================================
Private Sub UserControl_Paint()
    'During run-time the control is painted via `WM_PAINT` subclass, but to make the _
     buttons visible during design-time we call the shared paint routine here
    If blu.UserMode = False Then Call PaintButtons
End Sub

'CONTROL ReadPropertes : The ActiveX control is being loaded _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Get the handle to the parent form, even if the control is in a container
    Let ParentForm_hWnd = blu.GetParentForm_hWnd( _
        StartWith:=UserControl.Parent, GetMDIParent:=False _
    )
    
    'Proceed only if code is running (don't run in IDE's design mode)
    If blu.UserMode = True Then
        
        'Attach the mouse tracking
        Set MouseEvents = New bluMouseEvents
        'Show the hand-pointer on the control
        Let MouseEvents.MousePointer = IDC_HAND
        Call MouseEvents.Attach(UserControl.hWnd)
        
        'Subclass this user control to do flicker-free painting
        Set Magic = New bluMagic
        Call Magic.ssc_Subclass(UserControl.hWnd, 0, 1, Me)
        Call Magic.ssc_AddMsg( _
            UserControl.hWnd, MSG_BEFORE, _
            WM_PAINT, WM_ERASEBKGND _
        )
        
        '------------------------------------------------------------------------------
        
        'Determine if the form is borderless to begin with; we will need to add a _
         temporary border in order to add a shadow to the form:
        
        'Check if there's not any kind of window border
        If Not CBool(WindowStyle And (WS.WS_BORDER Or WS.WS_DLGFRAME Or WS.WS_THICKFRAME)) _
        And Not CBool(WindowStyleEx And WS_EX.WS_EX_TOOLWINDOW) _
        Then
            'Mark as originally borderless. Should the DWM switch off we normally _
             restore the border, but we can remember to leave it off
            Let WasBorderless = True
        End If
        
        'Set the extended window style to allow the chroma key transparency
        Call user32_SetWindowLongEx( _
            ParentForm_hWnd, GWL_EXSTYLE, WindowStyleEx Or WS_EX_LAYERED _
        )
        
        'Check if we will be able to (at the moment) do the borderless trick. _
         This is only possible on Vista and above, as long as the Desktop Window _
         Manager is enabled (hardware acceleration) and the theme is not Aero Basic / _
         Classic or High Contrast. This class will watch out for theme changes and _
         remove the form borders if it becomes possible whilst running
        Let My_IsBorderless = IsDWMAvailable()
        
        'Subclass the parent form and begin listening into the message stream; _
         see the subclass section at the bottom of this file
        '(We pass the user param as `HTCAPTION` so you can drag the form from anywhere)
        Call Magic.ssc_Subclass(ParentForm_hWnd, HT.HTCAPTION, 1, Me)
        Call Magic.ssc_AddMsg( _
            ParentForm_hWnd, MSG_BEFORE, _
            WM_NCCALCSIZE, WM_GETMINMAXINFO, _
            WM_DWMCOMPOSITIONCHANGED, WM_THEMECHANGED, _
            WM_ACTIVATE, WM_NCLBUTTONDOWN, WM_LBUTTONDOWN, WM_LBUTTONDBLCLK _
        )
        
        'When we remove the borders what we're really doing is expanding the form into _
         the borders so the form is bigger than it was before. This might be a problem _
         for you in some instances where you expect your form to be the same size with _
         a border, as without (e.g. tool windows, about screens). To solve this we _
         shrink the form by the size of the borders to make it the intended size again
        Let Borders = GetNonClientSize()
        If My_IsBorderless = True And WasBorderless = False Then
            Call RepositionForm
        Else
            'Basically force `WM_NCCALCSIZE` to fire (removes the borders). _
             `RepositionForm` above does that, but here we need to do so ourselves
            Call user32_SetWindowPos( _
                ParentForm_hWnd, 0, 0, 0, 0, 0, _
                SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE _
            )
        End If
    End If
    
    'Read and set up the ActiveX properties
    With PropBag
        Let Me.ActiveColour = .ReadProperty(Name:="ActiveColour", DefaultValue:=blu.ActiveColour)
        Let Me.BaseColour = .ReadProperty(Name:="BaseColour", DefaultValue:=blu.BaseColour)
        Let Me.Style = .ReadProperty(Name:="Style", DefaultValue:=bluSTYLE.Normal)
        Let Me.AlwaysOnTop = .ReadProperty(Name:="AlwaysOnTop", DefaultValue:=False)
        'Apply the chroma key, fixing the transparent pixel row
        Let Me.ChromaKey = .ReadProperty(Name:="ChromaKey", DefaultValue:=&H123456)
        Let My_MinWidth = .ReadProperty(Name:="MinWidth", DefaultValue:=0)
        Let My_MinHeight = .ReadProperty(Name:="MinHeight", DefaultValue:=0)
        Let My_MaxWidth = .ReadProperty(Name:="MaxWidth", DefaultValue:=0)
        Let My_MaxHeight = .ReadProperty(Name:="MaxHeight", DefaultValue:=0)
    End With
End Sub

'CONTROL Resize : The user is trying to resize the control on the form design _
 ======================================================================================
Private Sub UserControl_Resize()
    'This control can't be resized
    Let UserControl.Height = blu.Ypx(blu.Metric)
    Let UserControl.Width = blu.Xpx(blu.Metric * 3)
    Call UserControl_Paint
End Sub

'CONTROL Terminate : Clean up _
 ======================================================================================
Private Sub UserControl_Terminate()
    'The object won't have been initialised in the IDE at design time
    If Not Magic Is Nothing Then
        'Detach subclassing from the control for flicker-free painting
        Call Magic.ssc_DelMsg( _
            UserControl.hWnd, MSG_BEFORE, _
            WM_PAINT, WM_ERASEBKGND _
        )
        Call Magic.ssc_UnSubclass(UserControl.hWnd)
        
        'Detatch the window messages
        Call Magic.ssc_DelMsg( _
            ParentForm_hWnd, MSG_BEFORE, _
            WM_NCCALCSIZE, WM_GETMINMAXINFO, _
            WM_DWMCOMPOSITIONCHANGED, WM_THEMECHANGED, _
            WM_ACTIVATE, WM_NCLBUTTONDOWN, WM_LBUTTONDOWN, WM_LBUTTONDBLCLK _
        )
        Call Magic.ssc_UnSubclass(ParentForm_hWnd)
        
        'Detatch the subclassing from the controls registered as non-client handlers _
         (i.e. move / resize boxes). There might be an error here if you destroyed _
         these controls programatically before we get to this point. (`bluMagic` will _
         unsubclass them itself, so it won't crash, but may still warn you)
        Dim i As Long
        For i = 1 To NonClientHandlers.Count
            Call Magic.ssc_DelMsg( _
                NonClientHandlers.Item(i), MSG_BEFORE, _
                WM_LBUTTONDOWN, WM_LBUTTONDBLCLK _
            )
            Call Magic.ssc_UnSubclass(NonClientHandlers.Item(i))
        Next
        Call Magic.ssc_Terminate
        
        Set Magic = Nothing
        Set NonClientHandlers = Nothing
    End If
    'Detatch the mouse tracking subclassing
    Set MouseEvents = Nothing
End Sub

'CONTROL WriteProperties _
 ======================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        'Common blu properties
        Call .WriteProperty(Name:="ActiveColour", Value:=My_ActiveColour, DefaultValue:=blu.ActiveColour)
        Call .WriteProperty(Name:="BaseColour", Value:=My_BaseColour, DefaultValue:=blu.BaseColour)
        Call .WriteProperty(Name:="Style", Value:=My_Style, DefaultValue:=bluSTYLE.Normal)
        'Window properties
        Call .WriteProperty(Name:="AlwaysOnTop", Value:=My_AlwaysOnTop, DefaultValue:=False)
        Call .WriteProperty(Name:="ChromaKey", Value:=My_ChromaKey, DefaultValue:=&H123456)
        Call .WriteProperty(Name:="MinWidth", Value:=My_MinWidth, DefaultValue:=0)
        Call .WriteProperty(Name:="MinHeight", Value:=My_MinHeight, DefaultValue:=0)
        Call .WriteProperty(Name:="MaxWidth", Value:=My_MaxWidth, DefaultValue:=0)
        Call .WriteProperty(Name:="MaxHeight", Value:=My_MaxHeight, DefaultValue:=0)
    End With
End Sub

'EVENT <MouseEvents> MouseIn : The mouse has entered the control _
 ======================================================================================
Private Sub MouseEvents_MouseIn()
    Let IsMouseDown = False
    Call Refresh
End Sub

'EVENT <MouseEvents> MouseOut : The mouse has gone out of the control _
 ======================================================================================
Private Sub MouseEvents_MouseOut()
    Let IsMouseDown = False
    Call Refresh
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY ActiveColour _
 ======================================================================================
Public Property Get ActiveColour() As OLE_COLOR
    Let ActiveColour = My_ActiveColour
End Property
Public Property Let ActiveColour(ByVal NewColour As OLE_COLOR)
    Let My_ActiveColour = NewColour
    Call Refresh
    Call UserControl.PropertyChanged("ActiveColour")
End Property

'PROPERTY BaseColour _
 ======================================================================================
Public Property Get BaseColour() As OLE_COLOR
    Let BaseColour = My_BaseColour
End Property
Public Property Let BaseColour(ByVal NewColour As OLE_COLOR)
    Let My_BaseColour = NewColour
    Call Refresh
    Call UserControl.PropertyChanged("BaseColour")
End Property

'PROPERTY Style _
 ======================================================================================
Public Property Get Style() As bluSTYLE
    Let Style = My_Style
End Property
Public Property Let Style(ByVal NewStyle As bluSTYLE)
    Let My_Style = NewStyle
    Call Refresh
    Call UserControl.PropertyChanged("Style")
End Property

'PROPERTY AlwaysOnTop _
 ======================================================================================
Public Property Get AlwaysOnTop() As Boolean: Let AlwaysOnTop = My_AlwaysOnTop: End Property
Public Property Let AlwaysOnTop(ByVal State As Boolean)
    Let My_AlwaysOnTop = State
    'We don't want to do this during design-time
    If blu.UserMode = True Then
        'You can't set this with `WS_EX_TOPMOST`, use `SetWindowPos` instead, _
         Thanks goes to Karl E. Petterson's CFormBorder class for alerting me to this _
         <vb.mvps.org/samples/FormBdr/>
        'NOTE: The form won't always stay on top when running from the IDE _
         <support.microsoft.com/kb/192254>
        Call user32_SetWindowPos( _
            ParentForm_hWnd, IIf(State = True, HWND_TOPMOST, HWND_NOTOPMOST), _
            0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE _
        )
    End If
    Call UserControl.PropertyChanged("AlwaysOnTop")
End Property

'PROPERTY ChromaKey : Set what colour is transparent in the form _
 ======================================================================================
Public Property Get ChromaKey() As OLE_COLOR: Let ChromaKey = My_ChromaKey: End Property
Public Property Let ChromaKey(ByVal Colour As OLE_COLOR)
    Let My_ChromaKey = Colour
    'We don't need to actually do this in design time
    If blu.UserMode = True Then
        'Update the transparent colour used on the form
        Call user32_SetLayeredWindowAttributes( _
            ParentForm_hWnd, blu.OLETranslateColor(My_ChromaKey), 0, LWA.LWA_COLORKEY _
        )
    End If
    Call UserControl.PropertyChanged("ChromaKey")
End Property

'PROPERTY CloseButton : If the form has a close button _
 ======================================================================================
Public Property Get CloseButton() As Boolean
Attribute CloseButton.VB_MemberFlags = "400"
    Let CloseButton = CBool(WindowStyle And WS_SYSMENU)
End Property
Public Property Let CloseButton(ByVal State As Boolean)
    'NOTE: Min/Max buttons will not show if there is no close button
    'TODO
End Property

'PROPERTY IsBorderless : If the form has been made borderless, with shadow _
 ======================================================================================
Public Property Get IsBorderless() As Boolean
Attribute IsBorderless.VB_MemberFlags = "400"
    Let IsBorderless = My_IsBorderless
End Property

'PROPERTY MinWidth _
 ======================================================================================
Public Property Get MinWidth() As Long: Let MinWidth = My_MinWidth: End Property
Public Property Let MinWidth(ByVal Width As Long)
    Let My_MinWidth = Width
    Call UserControl.PropertyChanged("MinWidth")
End Property

'PROPERTY MinHeight _
 ======================================================================================
Public Property Get MinHeight() As Long: Let MinHeight = My_MinHeight: End Property
Public Property Let MinHeight(ByVal Height As Long)
    Let My_MinHeight = Height
    Call UserControl.PropertyChanged("MinHeight")
End Property

'PROPERTY MaxWidth _
 ======================================================================================
Public Property Get MaxWidth() As Long: Let MaxWidth = My_MaxWidth: End Property
Public Property Let MaxWidth(ByVal Width As Long)
    Let My_MaxWidth = Width
    Call UserControl.PropertyChanged("MaxWidth")
End Property

'PROPERTY MinHeight _
 ======================================================================================
Public Property Get MaxHeight() As Long: Let MaxHeight = My_MaxHeight: End Property
Public Property Let MaxHeight(ByVal Height As Long)
    Let My_MaxHeight = Height
    Call UserControl.PropertyChanged("MaxHeight")
End Property

'PROPERTY MinButton : If the form has a minimise button or not _
 ======================================================================================
Public Property Get MinButton() As Boolean
Attribute MinButton.VB_MemberFlags = "400"
    Let MinButton = CBool(WindowStyle And WS_MINIMIZEBOX)
End Property
Public Property Let MinButton(ByVal State As Boolean)
    'Remove the minimise button from the form
    Call ToggleWindowBits(WS_MINIMIZEBOX, State)
    'TODO: Remove the minimise menu item in the system menu (alt+spc)
    'TODO: Update our control box display
End Property

'PROPERTY MaxButton : If the form has a maximise button or not _
 ======================================================================================
Public Property Get MaxButton() As Boolean
Attribute MaxButton.VB_MemberFlags = "400"
    Let MaxButton = CBool(WindowStyle And WS_MAXIMIZEBOX)
End Property
Public Property Let MaxButton(ByVal State As Boolean)
    'Remove the maximise button from the form
    Call ToggleWindowBits(WS_MAXIMIZEBOX, State)
    'TODO: Remove the maximise menu item in the system menu (alt+spc)
    'TODO: Update our control box display
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'RegisterMoveHandler : Set a control to act as a title bar, moving the form _
 ======================================================================================
Public Sub RegisterMoveHandler(ByVal Target As Object)
    Call RegisterNonClientHandler(Target.hWnd, MoveHandler)
End Sub

'RegisterSizeHandler : Set a control to act as a corner sizing box _
 ======================================================================================
Public Sub RegisterSizeHandler(ByVal Target As Object)
    Call RegisterNonClientHandler(Target.hWnd, SizeHandler)
End Sub

'/// PRIVATE PROPERTIES ///////////////////////////////////////////////////////////////

'PROPERTY WindowStyle _
 ======================================================================================
Private Property Get WindowStyle() As WS
    'Retrieve the window style bits
    Let WindowStyle = blu.user32_GetWindowLong(ParentForm_hWnd, GWL_STYLE)
End Property
Private Property Let WindowStyle(ByVal NewStyle As WS)
    'Set the new window bits
    Call blu.user32_SetWindowLong(ParentForm_hWnd, GWL_STYLE, NewStyle)
End Property

'PROPERTY WindowStyleEx _
 ======================================================================================
Private Property Get WindowStyleEx() As WS_EX
    'Retrieve the extended window style bits
    Let WindowStyleEx = blu.user32_GetWindowLong(ParentForm_hWnd, GWL_EXSTYLE)
End Property
Private Property Let WindowStyleEx(ByVal NewStyle As WS_EX)
    'Set the new extended window bits
    Call blu.user32_SetWindowLongEx(ParentForm_hWnd, GWL_EXSTYLE, NewStyle)
End Property

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Exists : Check if an item exists in a Collection object _
 ======================================================================================
'<stackoverflow.com/questions/40651/check-if-a-record-exists-in-a-vb6-collection/9535221#9535221>
Private Function Exists(ByVal Key As String, ByRef Col As Collection) As Boolean
    Dim var As Variant
    
TryObject:
    On Error GoTo ExistsTryObject
    Set var = Col(Key)
    Let Exists = True
    Exit Function

TryNonObject:
    On Error GoTo ExistsTryNonObject
    Let var = Col(Key)
    Let Exists = True
    Exit Function

ExistsTryObject:
    'This will reset your Err Handler
    Resume TryNonObject

ExistsTryNonObject:
    Let Exists = False
End Function

'IsDLLAndProcedureAvailable : Check if we'll be able to make a DLL call _
 ======================================================================================
Private Function IsDLLAndProcedureAvailable(ByVal DLL As String, ByVal Procedure As String) As Boolean
    Dim hndModule As Long
    
    'Try the first method (apparently buggy sometimes)
    Let hndModule = kernel32_GetModuleHandle(DLL)
    If hndModule = 0 Then
        'If that failed, try the alternative
        Let hndModule = kernel32_LoadLibrary(DLL)
        'If that failed the DLL doesn't exist on the system
        If hndModule = 0 Then Exit Function
        'Test if the procedure exists in the DLL
        Let IsDLLAndProcedureAvailable = _
            (kernel32_GetProcAddress(hndModule, Procedure) <> 0)
        'The alternative method has to be freed
        Call kernel32_FreeLibrary(hndModule)
    Else
        'Test if the procedure exists in the DLL, _
         we don't need to free the handle with the first method
        Let IsDLLAndProcedureAvailable = _
            (kernel32_GetProcAddress(hndModule, Procedure) <> 0)
    End If
End Function

'IsDWMAvailable : If Vista and above, desktop composition is available _
 ======================================================================================
'The `DwmExtendFrameIntoClientArea` API we use for removing the borders is available _
 on Vista and above
Private Function IsDWMAvailable() As Boolean
    'If the "Show shadows under windows" option is off we will want to skip going _
     borderless. We could put a sinle pixel border around the window, but I leave _
     that to perhaps sometime in the future
    If blu.DropShadows = False Then Exit Function
    
    'We don't want to go borderless if the high contrast mode is on _
     (Windows 8 will return false for `DwmIsCompositionEnabled` anyway)
    If blu.IsHighContrastMode = True Then Exit Function
    
    'Check if the DWM APIs we want to use are available on the system
    If IsDLLAndProcedureAvailable("dwmapi", "DwmIsCompositionEnabled") = True Then
        'Now check if the hardware compositing is enabled: _
         Windows Vista Basic / 7 Starter do not have it all and in other versions _
         it can always be switched off. On Windows 8 it's always on all the time with _
         no off switch -- though it will report as off when the high-contrast theme _
         is selected! Note that on Vista and above selecting a high contrast theme _
         automatically enables the high contrast mode (something entirely different). _
         This is not the case on XP, but that doesn't concern us as the DWM APIs are _
         Vista+ only <blogs.msdn.com/b/oldnewthing/archive/2008/12/03/9167477.aspx>
        Dim Result As BOOL
        If dwmapi_DwmIsCompositionEnabled(Result) = S_OK Then
            'If the API call succeded, `Result` will give us our yes/no
            Let IsDWMAvailable = (Result = API_TRUE)
        End If
    End If
End Function

'GetNonClientSize : Retrives the size of the window title / borders _
 ======================================================================================
Private Function GetNonClientSize() As RECT
    'TODO: Different frame sizes (dialog / tool window)
    With GetNonClientSize
        Let .Bottom = blu.GetSystemMetric(SM_CYSIZEFRAME)
        Let .Top = .Bottom + blu.GetSystemMetric(SM_CYCAPTION)
        Let .Left = blu.GetSystemMetric(SM_CXSIZEFRAME)
        Let .Right = blu.GetSystemMetric(SM_CXSIZEFRAME)
    End With
    
'    Debug.Print "Borders: Top " & GetNonClientSize.Top & _
'                " Bottom " & GetNonClientSize.Bottom & _
'                " Left " & GetNonClientSize.Left & _
'                " Right " & GetNonClientSize.Right
'    Debug.Print "SysMetrics: " & _
'                "CXBORDER " & blu.GetSystemMetric(SM_CXBORDER) & _
'                " CXFIXEDFRAME " & blu.GetSystemMetric(SM_CXFIXEDFRAME) & _
'                " CXPADDEDBORDER " & blu.GetSystemMetric(SM_CXPADDEDBORDER) & _
'                " CXSIZEFRAME " & blu.GetSystemMetric(SM_CXSIZEFRAME) & _
'                " CYCAPTION " & blu.GetSystemMetric(SM_CYCAPTION) & _
'                " CYEDGE " & blu.GetSystemMetric(SM_CYEDGE)
'    Debug.Print "CXPADDEDBORDER " & blu.GetSystemMetric(SM_CXPADDEDBORDER)
End Function

'PaintButtons : Shared routine to paint all three buttons (for subclassed/IDE display) _
 ======================================================================================
Private Sub PaintButtons()
    'Clear the background
    
    'Set up the bounding box
    Dim Box As blu.RECT
    Call blu.user32_GetClientRect(UserControl.hWnd, Box)
    'Fill the background colour
    Call blu.gdi32_SetDCBrushColor( _
        UserControl.hDC, blu.OLETranslateColor(My_BaseColour) _
    )
    Call blu.user32_FillRect( _
        UserControl.hDC, Box, blu.gdi32_GetStockObject(DC_BRUSH) _
    )
    
    'Any buttons at all?
    If CloseButton = False Then Exit Sub
    
    Dim Pos As Long
    Let Pos = 1
    
    Call PaintButton(Quit, UserControl.ScaleWidth - (32 * Pos))
    
    If MaxButton Then
        Let Pos = Pos + 1
        Call PaintButton(Max, UserControl.ScaleWidth - (32 * Pos))
    End If
    
    If MinButton Then
        Let Pos = Pos + 1
        Call PaintButton(Min, UserControl.ScaleWidth - (32 * Pos))
    End If
End Sub

'PaintButton : Draw a min/max/close button _
 ======================================================================================
Private Sub PaintButton( _
    ByVal ButtonType As bluBorderless_ButtonType, ByVal Left As Long _
)
    Dim BackColour As OLE_COLOR
    Dim ForeColour As OLE_COLOR
    
    'Set up the bounding box
    Dim Box As blu.RECT
    Call blu.user32_SetRect(RECTToSet:=Box, _
        Left:=Left, Top:=0, Right:=Left + 32, Bottom:=UserControl.ScaleHeight _
    )
    
    'Is the mouse in it?
    Dim MousePos As blu.POINT
    Let MousePos = blu.GetMousePos_Window(UserControl.hWnd)
    
    If blu.user32_PtInRect(Box, MousePos.X, MousePos.Y) = API_TRUE Then
        'Mark which button is hovered to act upon mouse-up
        Let WhichButton = ButtonType
        
        If IsMouseDown Then
            If ButtonType = Quit Then
                Let BackColour = blu.ClosePressColour
            Else
                Let BackColour = My_ActiveColour
            End If
            Let ForeColour = My_BaseColour
        Else
            If ButtonType = Quit Then
                Let ForeColour = My_BaseColour
                Let BackColour = blu.CloseHoverColour
            Else
                Let ForeColour = blu.TextHoverColour
                Let BackColour = My_BaseColour
            End If
        End If
        
    ElseIf My_Style = Normal Then
        Let BackColour = My_BaseColour
        Let ForeColour = blu.TextColour
    Else
        Let BackColour = My_ActiveColour
        Let ForeColour = My_BaseColour
    End If
    
    'Draw the button background
    Call blu.gdi32_SetDCBrushColor( _
        UserControl.hDC, blu.OLETranslateColor(BackColour) _
    )
    Call blu.user32_FillRect( _
        UserControl.hDC, Box, blu.gdi32_GetStockObject(DC_BRUSH) _
    )
    
    'Draw the glyph
    Dim Letter As String
    Select Case ButtonType
        Case bluBorderless_ButtonType.Quit: Let Letter = "r"
        Case bluBorderless_ButtonType.Min: Let Letter = "0"
        Case bluBorderless_ButtonType.Max
            'Use maximise or restore glyph?
            If blu.user32_IsZoomed(ParentForm_hWnd) = API_TRUE Then _
                Let Letter = "2" Else Let Letter = "1"
    End Select
    
    Call blu.DrawText( _
        hndDeviceContext:=UserControl.hDC, BoundingBox:=Box, _
        Text:=Letter, Colour:=ForeColour, Alignment:=vbCenter, _
        Orientation:=Horizontal, FontName:="Marlett", FontSizePx:=14 _
    )
End Sub

'Refresh _
 ======================================================================================
Private Sub Refresh()
    'Queue a `WM_PAINT` message to repaint the control box
    Dim ClientRECT As blu.RECT
    Call blu.user32_GetClientRect(UserControl.hWnd, ClientRECT)
    Call user32_InvalidateRect(UserControl.hWnd, ClientRECT, API_FALSE)
End Sub

'RegisterNonClientHandler _
 ======================================================================================
Private Sub RegisterNonClientHandler( _
    ByVal hndWindow As Long, _
    ByVal HandlerType As NCH_TYPE _
)
    If Exists(Key:=hndWindow, Col:=NonClientHandlers) = False Then
        
        Call NonClientHandlers.Add(Item:=hndWindow, Key:=CStr(hndWindow))
        
        Call Magic.ssc_Subclass( _
            hndWindow, _
            Choose(HandlerType, HT.HTCAPTION, HT.HTBOTTOMRIGHT), _
            1, Me _
        )
        Call Magic.ssc_AddMsg( _
            hndWindow, MSG_BEFORE, _
            WM_LBUTTONDOWN, WM_LBUTTONDBLCLK _
        )
    End If
End Sub

'RepositionForm : Shift/resize the form to allow for the borders being added/removed _
 ======================================================================================
Private Function RepositionForm(Optional ByVal DoRemove As Boolean = True)
    'Because what's actually happening when we "remove" the form borders is that the _
     internal ("client") space is being expanded outwards _into_ the borders, the _
     form actually gets bigger than it was before. We need to move and resize the _
     form slightly to account for this (as well as when the borders get added back on)
    
    'Get the internal size of the window (sans-borders), on a window that has been _
     made borderless this will of course be the size of the whole window
    Dim WindowRECT As RECT
    Call blu.user32_GetWindowRect(ParentForm_hWnd, WindowRECT)
    
    'Are the borders being added or removed?
    If DoRemove = True Then
        'The borders are in the process of being removed, shrink the form
        With WindowRECT
            Let .Left = .Left + Borders.Left
            Let .Right = .Right - Borders.Right
            Let .Top = .Top + Borders.Top
            Let .Bottom = .Bottom - Borders.Bottom
        End With
    Else
        'The borders are being added back on, grow the form
        With WindowRECT
            Let .Left = .Left - Borders.Left
            Let .Right = .Right + Borders.Right
            Let .Top = .Top - Borders.Top
            Let .Bottom = .Bottom + Borders.Bottom
        End With
    End If
    
    'Move / resize the window and fire `WM_NCCALCSIZE` to ensure the borders are _
     added / removed accordingly
    Call user32_SetWindowPos( _
        ParentForm_hWnd, 0, _
        WindowRECT.Left, WindowRECT.Top, _
        (WindowRECT.Right - WindowRECT.Left), _
        (WindowRECT.Bottom - WindowRECT.Top), _
        SWP_NOSENDCHANGING Or SWP_NOACTIVATE Or SWP_FRAMECHANGED _
    )
End Function

'ToggleWindowBits : Turn Window features on/off _
 ======================================================================================
Private Function ToggleWindowBits( _
    ByVal WindowBit As WS, ByVal State As Boolean _
) As Long
    If State Then
        Let WindowStyle = WindowStyle Or WindowBit
    Else
        Let WindowStyle = WindowStyle And Not WindowBit
    End If
End Function

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
    ByRef UserParam As HT _
)
    'This will be used to extend the form into its borders
    Dim Margin As MARGINS
    
    'WM_NCCALCSIZE : Define the non-client (border) area _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms632634%28v=vs.85%29.aspx>
    If Message = WM_NCCALCSIZE And wParam = 1 Then '-----------------------------------
        '`WM_NCCALCSIZE` is "sent when the size and position of a window's client area _
         must be calculated". Microsoft have provided a very easy way to remove the _
         window borders: "When wParam is TRUE, simply returning 0 without processing _
         the NCCALCSIZE_PARAMS rectangles will cause the client area to resize to the _
         size of the window, including the window frame. This will remove the window _
         frame and caption items from your window, leaving only the client area _
         displayed."
     
        'If the form was originally borderless we have to manage adding and removing _
         a temporary border (which then gets covered over) to retain the window shadow
        If WasBorderless = True Then
            'Get the current border style on the window
            Dim WStyle As Long
            Let WStyle = blu.user32_GetWindowLong(ParentForm_hWnd, GWL_STYLE)
            
            'Are we adding or removing the temporary border? (when DWM state changes)
            If My_IsBorderless = True Then
                'Add the temporary border
                Let WStyle = WStyle Or WS.WS_CAPTION
            Else
                'Remove the temporary border returning the form to a regular _
                 borderless form without shadow
                Let WStyle = WStyle And Not WS.WS_CAPTION
            End If
            'Apply the new border style
            Call user32_SetWindowLong(ParentForm_hWnd, GWL_STYLE, WStyle)
        End If
        
        'At this point, if DWM is off we can't expand the form into the borders, _
         just leave the form as is (for example, WindowsXP)
        If My_IsBorderless = False Then Exit Sub
        
        'Extend the frame into the client area by one pixel row so that the window _
         shadow remains even though the form appears borderless. That one pixel row _
         appears transparent, which is fixed by use of `SetLayeredWindowAttributes`
        Let Margin.Bottom = 1
        Call dwmapi_DwmExtendFrameIntoClientArea( _
            ParentForm_hWnd, Margin _
        )
        
        'There's an issue with maximising the form -- maximised forms are actually _
         bigger than the screen, to account for pushing the window border out of the _
         edges of the screen. We need to check if the form has been maximised and _
         adjust the borders specifcally
        '<blogs.msdn.com/b/oldnewthing/archive/2012/03/26/10287385.aspx>
        If blu.user32_IsZoomed(ParentForm_hWnd) = API_TRUE Then
            'Coerce the lParam value into a structure
            Dim Params As NCCALCSIZE_PARAMS
            Call blu.kernel32_RtlMoveMemory(Params, ByVal lParam, Len(Params))
            'Remove the borders when maximised
            With Params.Rectangles(0)
                'When maximised the title bar is still visible, so we only need to _
                 remove the frame thickness, which we'll borrow from the bottom
                Let .Top = .Top + Borders.Bottom
                Let .Bottom = .Bottom - Borders.Bottom
                Let .Left = .Left + Borders.Left
                Let .Right = .Right - Borders.Right
            End With
            'Return our changes into the pointer provided to us
            Call blu.kernel32_RtlMoveMemory(ByVal lParam, Params, Len(Params))
        End If
        'We've handled this ourselves, don't allow Windows to further process this
        Let Handled = True
    
    '<blogs.msdn.com/b/llobo/archive/2006/08/01/maximizing-window-_2800_with-windowstyle_3d00_none_2900_-considering-taskbar.aspx>
    '<msdn.microsoft.com/en-us/library/windows/desktop/ms632626%28v=vs.85%29.aspx>
    ElseIf Message = WM_GETMINMAXINFO Then '-------------------------------------------
        'TODO: Must listen for work area change?
        Dim MinMax As MINMAXINFO
        Call blu.kernel32_RtlMoveMemory(MinMax, ByVal lParam, Len(MinMax))

        Dim hndMonitor As Long
        Let hndMonitor = user32_MonitorFromWindow( _
            ParentForm_hWnd, MONITOR_DEFAULTTONEAREST _
        )
        If hndMonitor <> 0 Then
            Dim Info As MONITORINFO
            Let Info.SizeOfMe = Len(Info)
            If user32_GetMonitorInfo(hndMonitor, Info) = API_TRUE Then
                With Info
                    If My_MinWidth > 0 Then Let MinMax.MinTrackSize.X = My_MinWidth
                    If My_MinHeight > 0 Then Let MinMax.MinTrackSize.Y = My_MinHeight
                    
                    Let MinMax.MaxPosition.X = Abs(.WorkArea.Left - .MonitorArea.Left)
                    Let MinMax.MaxPosition.Y = Abs(.WorkArea.Top - .MonitorArea.Top)
                    Let MinMax.MaxSize.X = IIf( _
                        Expression:=My_MaxWidth = 0, _
                        TruePart:=Abs(.WorkArea.Right - .WorkArea.Left), _
                        FalsePart:=My_MaxWidth _
                    )
                    Let MinMax.MaxSize.Y = IIf( _
                        Expression:=My_MaxHeight = 0, _
                        TruePart:=Abs(.WorkArea.Bottom - .WorkArea.Top), _
                        FalsePart:=My_MaxHeight _
                    )
                End With
            End If
        End If

        Call blu.kernel32_RtlMoveMemory(ByVal lParam, MinMax, Len(MinMax))
        Let Handled = True
    
    ElseIf Message = WM_THEMECHANGED _
        Or Message = WM_DWMCOMPOSITIONCHANGED Then '------------------------------------
        'Windows 8 does not send `WM_DWMCOMPOSITIONCHANGED` messages (DWM is always on); _
         even though it will report as *OFF* when using high contrast mode. Therefore _
         for Windows 8 we need to listen to `WM_THEMECHANGED` messages to spot when _
         the user changed to a high contrast theme. Windows Vista & 7 will send *BOTH* _
         messages, so we need to ignore one of them to avoid changing the borders _
         twice in one theme change
        If blu.WindowsVersion < 6.2 And Message = WM_THEMECHANGED Then Exit Sub
        
        'Is DWM switching on or off?
        Dim Old As Boolean
        Let Old = My_IsBorderless
        Let My_IsBorderless = IsDWMAvailable()
        
        'If form had borders and we're removing them:
        If Old = False And My_IsBorderless = True Then
            Let Borders = GetNonClientSize()
            If WasBorderless = False Then Call RepositionForm(True)
        End If
        'If form had been made borderless and we're putting them back on:
        If Old = True And My_IsBorderless = False Then
            If WasBorderless = False Then Call RepositionForm(False)
            
            Let Margin.Bottom = 1
            Call dwmapi_DwmExtendFrameIntoClientArea( _
                ParentForm_hWnd, Margin _
            )
        End If
        
        If My_IsBorderless <> Old Then
            'Notify the form of the change, it may need to rearrange some controls. _
             Unfortunately, we can't just fire the parent's `Resize` event ourselves: _
             <support.microsoft.com/kb/187740>
            RaiseEvent BorderlessStateChange(My_IsBorderless)
            
        End If
    
    '`WM_ACTIVATE` documentation: _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms646274%28v=vs.85%29.aspx>
    ElseIf Message = WM_ACTIVATE Then '------------------------------------------------
        'Visual Basic's Activate and Deactivate events only occur between the windows _
         in your app, not when your form gets or loses focus from another app. _
         By subclassing we can trap the true activate/deactivate events and raise _
         events ourselves. On Windows 8 inactive windows have no shadow so you may _
         want to dull the colours or add a thin border when the window deactivates
        
        'The window can be activated either by click or by other means such as _
         keyboard or API calls to bring the window to the front
        If wParam = WA.WA_ACTIVE Or wParam = WA.WA_CLICKACTIVE Then
            RaiseEvent Activate
        ElseIf wParam = WA.WA_INACTIVE Then
            RaiseEvent Deactivate
        End If
    
    '`WM_LBUTTONDOWN` : Left mouse button down _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645607%28v=vs.85%29.aspx>
    ElseIf Message = WM_LBUTTONDOWN Then '---------------------------------------------
        'This message is also processed by the non client handlers that act as title _
         bars / size-boxes for the parent form. If the left mouse button is down on _
         the registered control, then we pass through a message to the form to act _
         out the necessary action
        'WARNING: This causes the `Click` event of the form to no longer fire for the _
         left mouse button, but will for the right mouse button!
        Call blu.user32_SendMessage(ParentForm_hWnd, WM.WM_NCLBUTTONDOWN, UserParam, 0)
    
    '`WM_LBUTTONDBLCLK` : Left mouse button double click -- maximise / restore form _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645606%28v=vs.85%29.aspx>
    ElseIf Message = WM_LBUTTONDBLCLK Then '-------------------------------------------
        'Is the form currently maximised?
        Call user32_ShowWindow( _
            ParentForm_hWnd, _
            IIf( _
                blu.user32_IsZoomed(ParentForm_hWnd) = API_TRUE, _
                SW_RESTORE, SW.SW_SHOWMAXIMIZED _
            ) _
        )
    
    '`WM_NCLBUTTONDOWN` : Left mouse button down in the non-client (border) area _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645620%28v=vs.85%29.aspx>
    ElseIf Message = WM_NCLBUTTONDOWN Then '-------------------------------------------
        'When sending the window message to fake clicking the title bar or size box _
         Windows repositions the mouse! To stop this we combine the action (`SC_MOVE`) _
         with the non-client area (`HTCAPTION`). _
         This criticaly important discovery due to this page, and its project: _
         <www.codeproject.com/script/Content/ViewAssociatedFile.aspx?rzp=%2FKB%2Fvbscript%2Flavolpecw32%2Flvcw32h.zip&zep=DLLclasses%2FclsCustomWindow.cls&obid=11916&obtid=2&ovid=4> _
         <www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62605&lngWId=1>
        If wParam = HT.HTCAPTION Then
            Call blu.user32_PostMessage(ParentForm_hWnd, WM_SYSCOMMAND, SC_MOVE Or HT.HTCAPTION, lParam)
        ElseIf wParam = HT.HTBOTTOMRIGHT Then
            Call blu.user32_PostMessage(ParentForm_hWnd, WM_SYSCOMMAND, SC_SIZE Or (HT.HTBOTTOMRIGHT - 9), lParam)
        End If
    
    '`WM_SETCURSOR` : Windows is asking which mouse pointer to show _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms648382%28v=vs.85%29.aspx>
    ElseIf Message = WM_SETCURSOR Then '-----------------------------------------------
    
    '`WM_ERASEBKGND` _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms648055%28v=vs.85%29.aspx>
    ElseIf Message = WM_ERASEBKGND Then '----------------------------------------------
        'Don't paint the background so as to avoid flicker, _
         all painting will be done in `WM_PAINT`
        Let ReturnValue = 1
        Let Handled = True
     
    '`WM_PAINT` _
     <msdn.microsoft.com/en-us/library/windows/desktop/dd145213%28v=vs.85%29.aspx>
    ElseIf Message = WM_PAINT Then '---------------------------------------------------
        Call PaintButtons
        
    End If
    
'======================================================================================
'    C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
'--------------------------------------------------------------------------------------
'           DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'======================================================================================
End Sub
