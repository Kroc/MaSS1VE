Attribute VB_Name = "blu"
Option Explicit
'======================================================================================
'blu : A Modern Metro-esque graphical toolkit; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: blu

'Shared APIs and routines

'/// API //////////////////////////////////////////////////////////////////////////////

'COMMON _
 --------------------------------------------------------------------------------------

'In VB6 True is -1 and False is 0, but in the Win32 API it's 1 for True
Public Enum BOOL
    API_TRUE = 1
    API_FALSE = 0
End Enum

'Some of the more modern WIN32 APIs return 0 for success instead of 1, it varies _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa378137%28v=vs.85%29.aspx>
Public Enum HRESULT
    S_OK = 0
    S_FALSE = 1
End Enum

'A point _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162805%28v=vs.85%29.aspx>
Public Type POINT
   X As Long
   Y As Long
End Type

'Effectively the same as POINT, but used for better readability
Public Type Size
    Width As Long
    Height As Long
End Type

'MOUSE & KEYBOARD _
 --------------------------------------------------------------------------------------

'Get mouse position in screen coordinates _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms648390%28v=vs.85%29.aspx>
Private Declare Function user32_GetCursorPos Lib "user32" Alias "GetCursorPos" ( _
    ByRef Pos As POINT _
) As BOOL

'Convert an X/Y point on the screen to local coordinates of a window area _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162952%28v=vs.85%29.aspx>
Public Declare Function user32_ScreenToClient Lib "user32" Alias "ScreenToClient" ( _
    ByVal hndWindow As Long, _
    ByRef ScreenPoint As POINT _
) As BOOL

'Load a mouse cursor _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms648391%28v=vs.85%29.aspx>
Public Declare Function user32_LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hndInstance As Long, _
    ByVal CursorName As IDC _
) As Long

'Mouse cursor selection for the `MousePointer` property
Public Enum IDC
    'This is our own addition to tell us to not change it one way or another
    vbDefault = 0
    
    IDC_APPSTARTING = 32650&
    IDC_ARROW = 32512&
    IDC_CROSS = 32515&
    IDC_HAND = 32649&
    IDC_HELP = 32651&
    IDC_IBEAM = 32513&
    IDC_ICON = 32641&
    IDC_NO = 32648&
    IDC_SIZEALL = 32646&
    IDC_SIZENESW = 32643&
    IDC_SIZENS = 32645&
    IDC_SIZENWSE = 32642&
    IDC_SIZEWE = 32644&
    IDC_UPARROW = 32516&
    IDC_WAIT = 32514&
End Enum

'Sets the screen cursor _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms648393%28v=vs.85%29.aspx>
Public Declare Function user32_SetCursor Lib "user32" Alias "SetCursor" ( _
    ByVal hndCursor As Long _
) As Long

'TODO: GetAsyncKeyState
'<msdn.microsoft.com/en-us/library/ms646293%28VS.85%29.aspx>

'RECTANGLES _
 --------------------------------------------------------------------------------------

'A rectangle _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162897%28v=vs.85%29.aspx>
Public Type RECT
    Left                As Long
    Top                 As Long
    'It's important to note that the Right and Bottom edges are _exclusive_, that is, _
     the right-most and bottom-most pixel are not part of the overall width / height _
     <blogs.msdn.com/b/oldnewthing/archive/2004/02/18/75652.aspx>
    Right               As Long
    Bottom              As Long
End Type

'Populate a RECT structure _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145085%28v=vs.85%29.aspx>
Public Declare Function user32_SetRect Lib "user32" Alias "SetRect" ( _
    ByRef RECTToSet As RECT, _
    ByVal Left As Long, _
    ByVal Top As Long, _
    ByVal Right As Long, _
    ByVal Bottom As Long _
) As Long

'Shift a rectangle's position _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162746%28v=vs.85%29.aspx>
Public Declare Function user32_OffsetRect Lib "user32" Alias "OffsetRect" ( _
    ByRef RectToMove As RECT, _
    ByVal X As Long, _
    ByVal Y As Long _
) As BOOL

'Is a point in the rectangle? e.g. check if mouse is within a window _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162882%28v=vs.85%29.aspx>
Public Declare Function user32_PtInRect Lib "user32" Alias "PtInRect" ( _
    ByRef InRect As RECT, _
    ByVal X As Long, _
    ByVal Y As Long _
) As BOOL

'MEMORY _
 --------------------------------------------------------------------------------------

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Public Declare Sub kernel32_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef ptrDestination As Any, _
    ByRef ptrSource As Any, _
    ByVal Length As Long _
)

'Fill memory with zeroes (used to erase a BITMAPINFO structure) _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366920%28v=vs.85%29.aspx>
Public Declare Sub kernel32_RtlZeroMemory Lib "kernel32" Alias "RtlZeroMemory" ( _
    ByRef ptrDestination As Any, _
    ByVal Length As Long _
)

'DLL LOADING _
 --------------------------------------------------------------------------------------

'Try load a DLL _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms6831   99%28v=vs.85%29.aspx>
'"The GetModuleHandle function returns a handle to a mapped module without _
 incrementing its reference count. Therefore, do not pass a handle returned by _
 GetModuleHandle to the FreeLibrary function"
Public Declare Function kernel32_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
    ByVal ModuleName As String _
) As Long

'The above can apparently be buggy so this is used as a fallback _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms684175%28v=vs.85%29.aspx>
Public Declare Function kernel32_LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal Filename As String _
) As Long

'Free the resource associated with the above call _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms683152%28v=vs.85%29.aspx>
Public Declare Function kernel32_FreeLibrary Lib "kernel32" Alias "FreeLibrary" ( _
    ByVal hndModule As Long _
) As BOOL

'Get the address of a DLL procedure _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms683212%28v=vs.85%29.aspx/html>
Public Declare Function kernel32_GetProcAddress Lib "kernel32" Alias "GetProcAddress" ( _
    ByVal hndModule As Long, _
    ByVal ProcedureName As String _
) As Long

'INIT COMMON CONTROLS _
 --------------------------------------------------------------------------------------

'Get VB's controls to be themed by Windows _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb775697%28v=vs.85%29.aspx>
Private Declare Function comctl32_InitCommonControlsEx Lib "comctl32" Alias "InitCommonControlsEx" ( _
    ByRef Struct As INITCOMMONCONTROLSEX _
) As BOOL

'Used for the above to specify what control sets to theme _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb775507%28v=vs.85%29.aspx>
Private Type INITCOMMONCONTROLSEX
    SizeOfMe As Long
    Flags As ICC
End Type

Public Enum ICC
    ICC_ANIMATE_CLASS = &H80&           'Animation control
    ICC_BAR_CLASSES = &H4&              'Toolbar, status bar, trackbar, & tooltip
    ICC_COOL_CLASSES = &H400&           'Rebar
    ICC_DATE_CLASSES = &H100&           'Date and time picker
    ICC_HOTKEY_CLASS = &H40&            'Hot key control
    ICC_INTERNET_CLASSES = &H800&       'Web control
    ICC_LINK_CLASS = &H8000&            'Hyperlink control
    ICC_LISTVIEW_CLASSES = &H1&         'List view / header
    ICC_NATIVEFNTCTL_CLASS = &H2000&    'Native font control
    ICC_PAGESCROLLER_CLASS = &H1000&    'Pager control
    ICC_PROGRESS_CLASS = &H20&          'Progress bar
    ICC_TAB_CLASSES = &H8&              'Tab and tooltip
    ICC_TREEVIEW_CLASSES = &H2&         'Tree-view and tooltip
    ICC_UPDOWN_CLASS = &H10&            'Up-down control
    ICC_USEREX_CLASSES = &H200&         'ComboBoxEx
    ICC_STANDARD_CLASSES = &H4000&      'button, edit, listbox, combobox, & scroll bar
    ICC_WIN95_CLASSES = &HFF&           'Animate control, header, hot key, list-view,
                                         'progress bar, status bar, tab, tooltip,
                                         'toolbar, trackbar, tree-view, and up-down
    ICC_ALL_CLASSES = &HFDFF&           'All of the above
End Enum

'RESOURCES _
 --------------------------------------------------------------------------------------

'This will allow us to load the icons embedded in the EXE _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms648045%28v=vs.85%29.aspx>
Private Declare Function user32_LoadImage Lib "user32" Alias "LoadImageA" ( _
    ByVal hndInstance As Long, _
    ByVal ImageName As String, _
    ByVal ImageType As IMAGE, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal LoadFlag As LR _
) As Long

Private Enum IMAGE
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
End Enum

Private Enum LR
    LR_SHARED = &H8000&                 'Re-uses a resource. The system will unload it
End Enum

'We'll use this in `SetIcon` to find VB6's hidden top-level window _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633515%28v=vs.85%29.aspx>
Private Declare Function user32_GetWindow Lib "user32" Alias "GetWindow" ( _
    ByVal hndWindow As Long, _
    ByVal Command As GW _
) As Long

Private Enum GW
    GW_OWNER = 4
End Enum

'SYSTEM INFORMATION _
 --------------------------------------------------------------------------------------

'Structure for obtaining the Windows version _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724834%28v=vs.85%29.aspx>
Private Type OSVERSIONINFO
    SizeOfMe As Long
    MajorVersion As Long
    MinorVersion As Long
    BuildNumber As Long
    PlatformID As Long
    ServicePack As String * 128
End Type

'Get the windows version _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724451%28v=vs.85%29.aspx>
Private Declare Function kernel32_GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
    ByRef VersionInfo As OSVERSIONINFO _
) As BOOL

'Get/set various system configuration info _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx>
Private Declare Function user32_SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
    ByVal Action As SPI, _
    ByVal Param As Long, _
    ByRef ParamAny As Any, _
    ByVal WinIni As Long _
) As BOOL

Private Enum SPI
    'If the high contrast mode is enabled
    'NOTE: This is not the same thing as the high contrast theme -- on Windows XP
     'the user might use a high contrast theme without having high contrast mode on.
     'On Vista and above the high contrast mode is automatically enabled when a high
     'contrast theme is selected: <blogs.msdn.com/b/oldnewthing/archive/2008/12/03/9167477.aspx>
    SPI_GETHIGHCONTRAST = &H42
    
    'Number of "lines" to scroll with the mouse wheel
    SPI_GETWHEELSCROLLLINES = &H68
    'Number of "chars" to scroll with a horizontal mouse wheel
    SPI_GETWHEELSCROLLCHARS = &H6C
    
    'Determines whether the drop shadow effect is enabled.
    SPI_GETDROPSHADOW = &H1024
End Enum

'Used with `SystemParametersInfo` and `SPI_GETHIGHCONTRAST` to get info about the _
 high-contrast theme _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318112%28v=vs.85%29.aspx>
Private Type HIGHCONTRAST
    SizeOfMe As Long
    Flags As HCF
    ptrDefaultScheme As Long
End Type

'HIGHCONTRAST flags
Private Enum HCF
    HCF_HIGHCONTRASTON = &H1
End Enum

'<msdn.microsoft.com/en-us/library/windows/desktop/ms724385%28v=vs.85%29.aspx>
Private Declare Function user32_GetSystemMetrics Lib "user32" Alias "GetSystemMetrics" ( _
    ByVal Index As Long _
) As Long

Public Enum SM
    SM_CXVSCROLL = 2            'Width of vertical scroll bar
    SM_CYHSCROLL = 3            'Height of horizontal scroll bar
    
    SM_CYCAPTION = 4            'Title bar height
    SM_CXBORDER = 5             'Border width. Equivalent to SM_CXEDGE for windows
                                 'with the 3-D look
    SM_CYBORDER = 6             'Border width. Equivalent to the SM_CYEDGE for windows
                                 'with the 3-D look
    SM_CXFIXEDFRAME = 7         'Thickness of the frame around a window that has a
                                 'caption but is not sizable
    SM_CYFIXEDFRAME = 8         'Border height
    SM_CXSIZEFRAME = 32         'Resizable border horizontal thickness
    SM_CYSIZEFRAME = 33         'Resizable border vertical thickness
    SM_CYEDGE = 46              'The height of a 3-D border
    SM_CYSMCAPTION = 51         'Tool window title bar height
    SM_CXPADDEDBORDER = 92      'The amount of border padding for captioned windows
                                 'Not supported on XP
    
    SM_SWAPBUTTON = 23          'Mouse buttons are swapped
    SM_MOUSEHORIZONTALWHEELPRESENT = 91
    SM_MOUSEWHEELPRESENT = 75
    
    SM_CXICON = 11              'Default width of an icon (Usually 32 or 48)
    SM_CYICON = 12              'Default height of an icon
    SM_CXSMICON = 49            'Width of small icons (Usually 16)
    SM_CYSMICON = 50            'Height of small icons
End Enum

'WINDOWS _
 --------------------------------------------------------------------------------------

'A surpirsingly simple way of telling if a window is maximised _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633531%28v=vs.85%29.aspx>
Public Declare Function user32_IsZoomed Lib "user32" Alias "IsZoomed" ( _
    ByVal hndWindow As Long _
) As BOOL

'Get the parent window of a window _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633510%28v=vs.85%29.aspx>
Public Declare Function user32_GetParent Lib "user32" Alias "GetParent" ( _
    ByVal hndWindow As Long _
) As Long

'Retrieve current window attributes _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633584%28v=vs.85%29.aspx>
Public Declare Function user32_GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hndWindow As Long, _
    ByVal Index As GWL _
) As Long

Public Enum GWL
    GWL_STYLE = -16                 'Standard window styles
    GWL_EXSTYLE = -20               'Extended window styles
End Enum

'Set window attributes _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633591%28v=vs.85%29.aspx>
Public Declare Function user32_SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hndWindow As Long, _
    ByVal Index As GWL, _
    ByVal NewLong As WS _
) As Long
'This one is just for IDE-friendliness
Public Declare Function user32_SetWindowLongEx Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hndWindow As Long, _
    ByVal Index As GWL, _
    ByVal NewLong As WS_EX _
) As Long

'Window styles _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms632600%28v=vs.85%29.aspx>
Public Enum WS
    'Standard window styles (via `GWL_STYLE`)
    WS_BORDER = &H800000            'Thin-line border
    WS_CAPTION = &HC00000           'Title bar (includes WS_BORDER)
    WS_CHILD = &H40000000           'Is a child window
    WS_CLIPCHILDREN = &H2000000     'Don't paint in the area of child windows
    WS_CLIPSIBLINGS = &H4000000     'Clip sibling windows (to deal with overlap)
    WS_DISABLED = &H8000000         'Window is initially disabled
    WS_DLGFRAME = &H400000          'Dialog box style border, cannot have title bar
    WS_GROUP = &H20000              'Part of a group - i.e. radio buttons
    WS_HSCROLL = &H100000           'Has a horizontal scroll bar
    WS_MAXIMIZE = &H1000000         'Window is initially maximised
    WS_MAXIMIZEBOX = &H10000        'Has maximize button
    WS_MINIMIZE = &H20000000        'Window is initally minimised
    WS_MINIMIZEBOX = &H20000        'Has minimize button
    WS_POPUP = &H80000000           'Is a popup window (cannot be WS_CHILD too)
    WS_SYSMENU = &H80000            'Has system menu (ALT+SPACE)
    WS_TABSTOP = &H10000            'Receives focus with the tab key
    WS_THICKFRAME = &H40000         'Has resizing borders
    WS_VISIBLE = &H10000000         'Is initially visible
    WS_VSCROLL = &H200000           'Has a vertical scroll bar
End Enum

'<msdn.microsoft.com/en-us/library/windows/desktop/ff700543%28v=vs.85%29.aspx>
Public Enum WS_EX
    'Extended window styles (via `GWL_EXSTYLE`)
    WS_EX_APPWINDOW = &H40000       'Show in taskbar
    WS_EX_CLIENTEDGE = &H200        'Sunken border
    WS_EX_DLGMODALFRAME = &H1       'Double border
    WS_EX_LAYERED = &H80000         'Layered, that is, can be translucent
    WS_EX_STATICEDGE = &H20000      '3D border for items that do not accept user input
    WS_EX_TOOLWINDOW = &H80
    WS_EX_WINDOWEDGE = &H100        'Border with raised edge
End Enum

'Get the dimensions of the whole window, including the border area _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633519%28v=vs.85%29.aspx>
Public Declare Function user32_GetWindowRect Lib "user32" Alias "GetWindowRect" ( _
    ByVal hndWindow As Long, _
    ByRef IntoRECT As RECT _
) As BOOL

'Get the size of the inside of a window (excluding the titlebar / borders) _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633503%28v=vs.85%29.aspx>
Public Declare Function user32_GetClientRect Lib "user32" Alias "GetClientRect" ( _
    ByVal hndWindow As Long, _
    ByRef ClientRECT As RECT _
) As BOOL

'<msdn.microsoft.com/en-us/library/windows/desktop/dd145002%28v=vs.85%29.aspx>
Public Declare Function user32_InvalidateRect Lib "user32" Alias "InvalidateRect" ( _
    ByVal hndWindow As Long, _
    ByRef InvalidRECT As RECT, _
    ByVal EraseBG As BOOL _
) As BOOL

'Send a message from one window to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms644950%28v=vs.85%29.aspx>
Public Declare Function user32_SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hndWindow As Long, _
    ByVal Message As WM, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

'Sends a message to another window, but doesn't wait for a return value _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms644944%28v=vs.85%29.aspx>
Public Declare Function user32_PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hndWindow As Long, _
    ByVal Message As WM, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As BOOL

Public Enum WM
    WM_PAINT = &HF
    WM_ERASEBKGND = &H14                'Clearing background before drawing
    
    WM_ACTIVATE = &H6                   'Window got / lost focus
    WM_GETMINMAXINFO = &H24             'Determine min/max allowed window size
    WM_NCCALCSIZE = &H83                'Calculate non-client (border) area
    
    WM_SETCURSOR = &H20                 'Which cursor should the mouse have?
    WM_MOUSEMOVE = &H200
    WM_MOUSEWHEEL = &H20A               'Mouse wheel scrolled
    WM_LBUTTONDOWN = &H201              'Left mouse button is down
    WM_LBUTTONDBLCLK = &H203            'Left double-click
    WM_XBUTTONDOWN = &H20B              'Mouse X button pressed (Back / Forward)
    WM_MOUSEHWHEEL = &H20E              'Horizontal mouse wheel scrolled
    WM_MOUSEHOVER = &H2A1
    WM_MOUSELEAVE = &H2A3
    WM_NCLBUTTONDOWN = &HA1             'Left mouse button is down in a non-client area
    WM_NCRBUTTONDOWN = &HA4             'Right button, as above
    WM_NCMBUTTONDOWN = &HA7             'Middle button, as above
    
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
    
    WM_SYSCOMMAND = &H112               'System menu interaction (move / size &c.)
    WM_THEMECHANGED = &H31A             'Windows theme changed
    WM_DWMCOMPOSITIONCHANGED = &H31E    'DWM was enabled / disabled
End Enum

'BITMAPS: _
 --------------------------------------------------------------------------------------
 
'The header used on a .BMP file _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd183374%28v=vs.85%29.aspx>
Public Type BITMAPFILEHEADER
    Type As Integer
    Size As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffsetToBits As Long
End Type

'The header that describes a bitmap _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd183376%28v=vs.85%29.aspx>
Public Type BITMAPINFOHEADER
    SizeOfMe            As Long     'Number of bytes required by the structure
    Width               As Long     'Width of the bitmap, in pixels
    Height              As Long     'Height of the bitmap, in pixels. If negative the _
                                     bitmap is top-down instead of upside-down
    BitPlanes           As Integer  'Number of bit-planes; keep this as 1
    Depth               As Integer  'Bit-depth
    Compression         As Long     '=`gdi32_Compression`
    DataSize            As Long     'Length of the data for the bits (4-Byte aligned)
    pxpmX               As Long     'Pixels Per Metre on the X-Axis
    pxpmY               As Long     'Pixles Per Metre on the Y-Axis
    UsedColors          As Long     'Number of (palette) colours used in the image
    ImportantColors     As Long     'Number of palette colours deemed critical
End Type

'Compression descriptor for the DIB image in memory, we don't use this beyond the _
 `DI_RGB` value as we want our images to be manipulable
Public Enum gdi32_Compression
    BI_RGB = 0                      'Uncompressed
    BI_RLE8 = 1                     'Run-Length-Encoding designed for 8-Bit images
    BI_RLE4 = 2                     'Run-Length-Encoding designed for 4-Bit images
    BI_BITFIELDS = 3                'Allows use of 5-5-5 and 5-6-5 bit colours
    BI_JPEG = 4                     'It's a JPEG!
    BI_PNG = 5                      'It's a PNG!
End Enum

'A palette entry in the BITMAPINFO structure _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162938%28v=vs.85%29.aspx>
Public Type RGBQUAD
    Blue                As Byte
    Green               As Byte
    Red                 As Byte
    Reserved            As Byte
End Type

'A bitmap combines the header and the palette _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd183375%28v=vs.85%29.aspx/html>
Public Type BITMAPINFO
    Header              As BITMAPINFOHEADER
    Colors(0 To 255)    As RGBQUAD
End Type

'Create a memory Device Context compatible with an existing device (screen by default) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd183489%28v=vs.85%29.aspx>
Public Declare Function gdi32_CreateCompatibleDC Lib "gdi32" Alias "CreateCompatibleDC" ( _
    ByVal hndDeviceContext As Long _
) As Long

'Creates the DIB based on the `BITMAPINFO` structure passed _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd183494%28v=vs.85%29.aspx>
Public Declare Function gdi32_CreateDIBSection Lib "gdi32" Alias "CreateDIBSection" ( _
    ByVal hndDeviceContext As Long, _
    ByRef ptrBITMAPINFO As BITMAPINFO, _
    ByVal Usage As gdi32_Usage, _
    ByRef ptrBits As Long, _
    ByVal hndFileMappingObject As Long, _
    ByVal Offset As Long _
) As Long

Public Enum gdi32_Usage
    DIB_RGB_COLORS = 0
    DIB_PAL_COLORS = 1
End Enum

'Delete a Device Context _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd183533%28v=vs.85%29.aspx>
Public Declare Function gdi32_DeleteDC Lib "gdi32" Alias "DeleteDC" ( _
    ByVal hndDeviceContext As Long _
) As Long

'Get the raw data stream of the image (can manipulate as a byte array) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd144879%28v=vs.85%29.aspx>
Public Declare Function gdi32_GetDIBits Lib "gdi32" Alias "GetDIBits" ( _
    ByVal hndDeviceContext As Long, _
    ByVal hndDIB As Long, _
    ByVal StartScan As Long, _
    ByVal NumberOfScans As Long, _
    ByRef ptrBits As Any, _
    ByRef ptrBITMAPINFO As BITMAPINFO, _
    ByVal Usage As gdi32_Usage _
) As Long

'Set the image pixel data from a byte array _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162973%28v=vs.85%29.aspx>
Public Declare Function gdi32_SetDIBits Lib "gdi32" Alias "SetDIBits" ( _
    ByVal hndDeviceContext As Long, _
    ByVal hndDIB As Long, _
    ByVal StartScan As Long, _
    ByVal NumberOfScans As Long, _
    ByRef ptrBits As Any, _
    ByRef ptrBITMAPINFO As BITMAPINFO, _
    ByVal Usage As gdi32_Usage _
) As Long

'Get palette colour(s) on a DIB _
 <msdn.microsoft.com/en-us/library/dd144878(v=vs.85).aspx>
Public Declare Function gdi32_GetDIBColorTable Lib "gdi32" Alias "GetDIBColorTable" ( _
    ByVal hndDeviceContext As Long, _
    ByVal StartIndex As Long, _
    ByVal Count As Long, _
    ByRef ptrRGBQUAD As Any _
) As Long

'Set palette colour(s) on a DIB _
 <msdn.microsoft.com/en-us/library/dd162972(v=vs.85).aspx>
Public Declare Function gdi32_SetDIBColorTable Lib "gdi32" Alias "SetDIBColorTable" ( _
    ByVal hndDeviceContext As Long, _
    ByVal StartIndex As Long, _
    ByVal Count As Long, _
    ByRef ptrRGBQUAD As Any _
) As Long

'DRAWING _
 --------------------------------------------------------------------------------------

'Convert a system color (such as "button face" or "inactive window") to a RGB value _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms694353%28v=vs.85%29.aspx>
Private Declare Function olepro32_OleTranslateColor Lib "olepro32" Alias "OleTranslateColor" ( _
    ByVal OLEColour As OLE_COLOR, _
    ByVal hndPalette As Long, _
    ByRef ptrColour As Long _
) As Long

'Select a GDI object into a Device Context _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162957%28v=vs.85%29.aspx>
Public Declare Function gdi32_SelectObject Lib "gdi32" Alias "SelectObject" ( _
    ByVal hndDeviceContext As Long, _
    ByVal hndGdiObject As Long _
) As Long

'Delete a GDI object we created (the DIB) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd183539%28v=vs.85%29.aspx>
Public Declare Function gdi32_DeleteObject Lib "gdi32" Alias "DeleteObject" ( _
    ByVal hndGdiObject As Long _
) As BOOL

'Some handy pens / brushes already available that we don't have to create / destroy _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd144925%28v=vs.85%29.aspx>
Public Declare Function gdi32_GetStockObject Lib "gdi32" Alias "GetStockObject" ( _
    ByVal Which As STOCKOBJECT _
) As Long

Public Enum STOCKOBJECT
    WHITE_BRUSH = 0
    LTGRAY_BRUSH = 1
    GRAY_BRUSH = 2
    DKGRAY_BRUSH = 3
    BLACK_BRUSH = 4
    NULL_BRUSH = 5
    DC_BRUSH = 18
    
    WHITE_PEN = 6
    BLACK_PEN = 7
    NULL_PEN = 8
    DC_PEN = 19
    
    DEFAULT_PALETTE = 15
End Enum

'Set a colour to use for painting, without having to create / destroy a resource! _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162969%28v=vs.85%29.aspx>
Public Declare Function gdi32_SetDCBrushColor Lib "gdi32" Alias "SetDCBrushColor" ( _
    ByVal hndDeviceContext As Long, _
    ByVal Color As Long _
) As Long

'Move the origin point (0,0) used for painting. This is partciularly important when _
 rotating so as to ensure you rotate around the centrepoint of the shape / text _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145099%28v=vs.85%29.aspx>
Public Declare Function gdi32_SetViewportOrgEx Lib "gdi32" Alias "SetViewportOrgEx" ( _
    ByVal hndDeviceContext As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByRef PreviousOrigin As POINT _
) As BOOL

'Enable access to world transformations (that is, scaling and rotating) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162977%28v=vs.85%29.aspx>
Public Declare Function gdi32_SetGraphicsMode Lib "gdi32" Alias "SetGraphicsMode" ( _
    ByVal hndDeviceContext As Long, _
    ByVal Mode As GM _
) As Long

Public Enum GM
    GM_COMPATIBLE = 1
    GM_ADVANCED = 2
End Enum

'A transformation matrix, used by `Get/SetWorldTransform` to apply scaling & rotation _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145228%28v=vs.85%29.aspx>
Public Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

'Retrieve any current world transform (i.e. scaling and rotation) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd144953%28v=vs.85%29.aspx>
Public Declare Function gdi32_GetWorldTransform Lib "gdi32" Alias "GetWorldTransform" ( _
    ByVal hndDeviceContext As Long, _
    ByRef Transform As XFORM _
) As BOOL

'Set the world transform _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145104%28v=vs.85%29.aspx>
Public Declare Function gdi32_SetWorldTransform Lib "gdi32" Alias "SetWorldTransform" ( _
    ByVal hndDeviceContext As Long, _
    ByRef Transform As XFORM _
) As BOOL

'Paint an area of an image one colour _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162719%28v=vs.85%29.aspx>
Public Declare Function user32_FillRect Lib "user32" Alias "FillRect" ( _
    ByVal hndDeviceContext As Long, _
    ByRef Rectangle As RECT, _
    ByVal hndBrush As Long _
) As Long

'Paint a square box (without fill) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd144838%28v=vs.85%29.aspx>
Public Declare Function user32_FrameRect Lib "user32" Alias "FrameRect" ( _
    ByVal hndDeviceContext As Long, _
    ByRef Rectangle As RECT, _
    ByVal hndBrush As Long _
) As Long

'Copy an image or portion thereof to somewhere else _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd183370%28v=vs.85%29.aspx>
Public Declare Function gdi32_BitBlt Lib "gdi32" Alias "BitBlt" ( _
    ByVal hndDestDC As Long, _
    ByVal DestLeft As Long, _
    ByVal DestTop As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal hndSrcDC As Long, _
    ByVal SrcLeft As Long, _
    ByVal SrcTop As Long, _
    ByVal RasterOperation As VBRUN.RasterOpConstants _
) As Long

'Copy an image or portion thereof to somewhere else, stretching if necessary _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145120%28v=vs.85%29.aspx>
Public Declare Function gdi32_StretchBlt Lib "gdi32" Alias "StretchBlt" ( _
    ByVal hndDestDC As Long, _
    ByVal DestLeft As Long, _
    ByVal DestTop As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal hndSrcDC As Long, _
    ByVal SrcLeft As Long, _
    ByVal SrcTop As Long, _
    ByVal SrcWidth As Long, _
    ByVal SrcHeight As Long, _
    ByVal RasterOperation As VBRUN.RasterOpConstants _
) As Long

'Copy and optionally stretch an image, making one colour transparent _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd373586%28v=VS.85%29.aspx>
Public Declare Function gdi32_GdiTransparentBlt Lib "gdi32" Alias "GdiTransparentBlt" ( _
    ByVal hndDestDC As Long, _
    ByVal DestLeft As Long, _
    ByVal DestTop As Long, _
    ByVal DestWidth As Long, _
    ByVal DestHeight As Long, _
    ByVal hndSrcDC As Long, _
    ByVal SrcLeft As Long, _
    ByVal SrcTop As Long, _
    ByVal SrcWidth As Long, _
    ByVal SrcHeight As Long, _
    ByVal TransparentColour As Long _
) As Long

'TEXT: _
 --------------------------------------------------------------------------------------

'Create a font object for writing text with GDI _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd183499%28v=vs.85%29.aspx>
Public Declare Function gdi32_CreateFont Lib "gdi32" Alias "CreateFontA" ( _
    ByVal Height As Long, _
    ByVal Width As Long, _
    ByVal Escapement As Long, _
    ByVal Orientation As Long, _
    ByVal Weight As FW, _
    ByVal Italic As BOOL, _
    ByVal Underline As BOOL, _
    ByVal StrikeOut As BOOL, _
    ByVal CharSet As FDW_CHARSET, _
    ByVal OutputPrecision As FDW_OUT, _
    ByVal ClipPrecision As FDW_CLIP, _
    ByVal Quality As FDW_QUALITY, _
    ByVal PitchAndFamily As FDW_PITCHANDFAMILY, _
    ByVal Face As String _
) As Long

'Font weight: _
 "The weight of the font in the range 0 through 1000. For example, 400 is normal and _
  700 is bold. If this value is zero, a default weight is used. _
  The following values are defined for convenience:"
Public Enum FW
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_ULTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_REGULAR = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_DEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_ULTRABOLD = 800
    FW_HEAVY = 900
    FW_BLACK = 900
End Enum

'Font character set:
Public Enum FDW_CHARSET
    ANSI_CHARSET = 0
    ARABIC_CHARSET = 178        'Middle East language edition of Windows
    BALTIC_CHARSET = 186
    CHINESEBIG5_CHARSET = 136
    DEFAULT_CHARSET = 1         'Use system locale to determine character set
    EASTEUROPE_CHARSET = 238
    GB2312_CHARSET = 134
    GREEK_CHARSET = 161
    HANGEUL_CHARSET = 129
    HEBREW_CHARSET = 177        'Middle East language edition of Windows
    JOHAB_CHARSET = 130         'Korean language edition of Windows
    MAC_CHARSET = 77
    OEM_CHARSET = 255           'Operating system dependent character set
    RUSSIAN_CHARSET = 204
    SHIFTJIS_CHARSET = 128
    SYMBOL_CHARSET = 2
    THAI_CHARSET = 222          'Thai language edition of Windows
    TURKISH_CHARSET = 162
End Enum

'Font output precision: _
 "The output precision defines how closely the output must match the requested font's _
  height, width, character orientation, escapement, pitch, and font type. It can be _
  one of the following values:"
Public Enum FDW_OUT
    OUT_DEFAULT_PRECIS = 0      'The default font mapper behaviour
    OUT_DEVICE_PRECIS = 5       'Choose a Device font when the system contains
                                 'multiple fonts with the same name
    OUT_OUTLINE_PRECIS = 8      'Choose from TrueType and other outline-based fonts
    OUT_RASTER_PRECIS = 6       'Choose a raster font when the system contains
                                 'multiple fonts with the same name
    OUT_STRING_PRECIS = 1       'This value is not used by the font mapper,
                                 'but it is returned when raster fonts are enumerated
    OUT_STROKE_PRECIS = 3       'This value is not used by the font mapper, but it is
                                 'returned when TrueType, other outline-based fonts,
                                 'and vector fonts are enumerated
    OUT_TT_ONLY_PRECIS = 7      'Choose from only TrueType fonts
    OUT_TT_PRECIS = 4           'Choose a TrueType font when the system contains
                                 'multiple fonts with the same name
End Enum

'The clipping precision: _
 "The clipping precision defines how to clip characters that are partially outside the _
  clipping region. It can be one or more of the following values:"
Public Enum FDW_CLIP
    CLIP_DEFAULT_PRECIS = 0     'Specifies default clipping behavior
    CLIP_EMBEDDED = 128         'Use an embedded read-only font
    CLIP_LH_ANGLES = 16         'When this value is used, the rotation for all fonts
                                 'depends on whether the orientation of the coordinate
                                 'system is left-handed or right-handed
                                'If not used, device fonts always rotate counter-
                                 'clockwise, but the rotation of other fonts is
                                 'dependent on the orientation of the coordinate system
    CLIP_STROKE_PRECIS = 2      'Not used by the font mapper, but is returned when
                                 'raster, vector, or TrueType fonts are enumerated
                                'For compatibility, this value is always returned
                                 'when enumerating fonts
End Enum

'The output quality: _
 "The output quality defines how carefully GDI must attempt to match the logical-font _
  attributes to those of an actual physical font. It can be one of the following _
  values:"
Public Enum FDW_QUALITY
    ANTIALIASED_QUALITY = 4     'Font is antialiased if the font supports it and the
                                 'size is not too small or too large
    CLEARTYPE_QUALITY = 5       'Use ClearType (when possible) antialiasing method
    DEFAULT_QUALITY = 0         'Appearance of the font does not matter
    DRAFT_QUALITY = 1           'Appearance of the font is less important than when
                                 'the PROOF_QUALITY value is used. For GDI raster
                                 'fonts, scaling is enabled, which means that more
                                 'font sizes are available, but the quality may be
                                 'lower. Bold, italic, underline, and strikeout fonts
                                 'are synthesized, if necessary
    NONANTIALIASED_QUALITY = 3  'Font is never antialiased
    PROOF_QUALITY = 2           'Character quality of the font is more important than
                                 'exact matching of the logical-font attributes.
                                 'For GDI raster fonts, scaling is disabled and the
                                 'font closest in size is chosen. Although the chosen
                                 'font size may not be mapped exactly when
                                 'PROOF_QUALITY is used, the quality of the font is
                                 'high and there is no distortion of appearance.
                                 'Bold, italic, underline, and strikeout fonts are
                                 'synthesized, if necessary
End Enum

'The pitch and family of the font:
Public Enum FDW_PITCHANDFAMILY
    '"The two low-order bits specify the pitch of the font and can be one of the
     'following values:"
    DEFAULT_PITCH = 0
    FIXED_PITCH = 1
    VARIABLE_PITCH = 2
    '"The four high-order bits specify the font family and can be one of the
     'following values:"
    FF_DECORATIVE = 80          'Novelty fonts. Old English is an example
    FF_DONTCARE = 0             'Use default font
    FF_MODERN = 48              'Fonts with constant stroke width, with or without
                                 'serifs. Pica, Elite, and Courier New are examples
    FF_ROMAN = 16               'Fonts with variable stroke width and with serifs,
                                 'MS Serif is an example
    FF_SCRIPT = 64              'Fonts designed to look like handwriting,
                                 'Script and Cursive are examples
    FF_SWISS = 32               'Fonts with variable stroke width and without serifs,
                                 'MS Sans Serif is an example
End Enum

'Does what it says on the tin _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145093%28v=vs.85%29.aspx>
Public Declare Function gdi32_SetTextColor Lib "gdi32" Alias "SetTextColor" ( _
    ByVal hndDeviceContext As Long, _
    ByVal Color As Long _
) As Long

'Set the horizontal / vertical text alignment _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145091%28v=vs.85%29.aspx>
Public Declare Function gdi32_SetTextAlign Lib "gdi32" Alias "SetTextAlign" ( _
    ByVal hndDeviceContext As Long, _
    ByVal Flags As TA) As Long

'The text alignment by using a mask of the values in the following list. _
 Only one flag can be chosen from those that affect horizontal and vertical alignment. _
 In addition, only one of the two flags that alter the current position can be chosen
Public Enum TA
    TA_BASELINE = 24    'Align to the baseline of the text
    TA_BOTTOM = 8       'Align to the bottom edge of the bounding rectangle
    TA_CENTER = 6       'Align horizontally centered along the bounding rectangle
    TA_LEFT = 0         'Align to the left edge of the bounding rectangle
    TA_NOUPDATECP = 0   'Do not set the current point to the reference point
    TA_RIGHT = 2        'Align to the right edge of the bounding rectangle
    TA_TOP = 0          'Align to the top edge of the bounding rectangle
    TA_UPDATECP = 1     'Set the current point to the reference point
    TA_TOPCENTER = 6
End Enum

'Set transparent background for drawing text _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162965%28v=vs.85%29.aspx>
Public Declare Function gdi32_SetBkMode Lib "gdi32" Alias "SetBkMode" ( _
    ByVal hndDeviceContext As Long, _
    ByVal Mode As BKMODE _
) As Long

Public Enum BKMODE
    TRANSPARENT = 1
    OPAQUE = 2
End Enum

'Draw some text _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd145133%28v=vs.85%29.aspx>
Public Declare Function gdi32_TextOut Lib "gdi32" Alias "TextOutA" ( _
    ByVal hndDeviceContext As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Text As String, _
    ByVal Length As Long _
) As BOOL

'With this you can specify a bounding RECT so as to truncate text, i.e. "..." _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162498%28v=vs.85%29.aspx>
Public Declare Function user32_DrawText Lib "user32" Alias "DrawTextA" ( _
    ByVal hndDeviceContext As Long, _
    ByVal Text As String, _
    ByVal Length As Long, _
    ByRef BoundingBox As RECT, _
    ByVal Format As DT _
) As Long

Public Enum DT
    DT_TOP = &H0                    'Top align text
    DT_LEFT = &H0                   'Left align text
    DT_CENTER = &H1                 'Centre text horziontally
    DT_RIGHT = &H2                  'Right align center
    DT_VCENTER = &H4                'Centre text vertically
    DT_BOTTOM = &H8                 'Bottom align the text; `DT_SINGLELINE` only
    DT_WORDBREAK = &H10             'Word-wrap
    DT_SINGLELINE = &H20            'Single line only
    DT_EXPANDTABS = &H40            'Display tab characters
    DT_TABSTOP = &H80               'Set the tab size (see the MSDN documentation)
    DT_NOCLIP = &H100               'Don't clip the text outside the bounding box
    DT_EXTERNALLEADING = &H200      'Include the font's leading in the line height
    DT_CALCRECT = &H400             'Update the RECT to fit the bounds of the text,
                                     'but does not actually draw the text
    DT_NOPREFIX = &H800             'Do not render "&" as underscore (accelerator)
    DT_INTERNAL = &H1000            'Use the system font to calculate metrics
    DT_EDITCONTROL = &H2000         'Behave as a text-box control, clips any partially
                                     'visible line at the bottom
    DT_PATH_ELLIPSIS = &H4000       'Truncate in the middle (e.g. file paths)
    DT_END_ELLIPSIS = &H8000        'Truncate the text with "..."
    DT_MODIFYSTRING = &H10000       'Change the string to match the truncation
    DT_WORD_ELLIPSIS = &H40000      'Truncate any word outside the bounding box
    DT_HIDEPREFIX = &H100000        'Process accelerators, but hide the underline
End Enum

'/// PUBLIC VARS //////////////////////////////////////////////////////////////////////

'When a user control is nested in another user control, the `Ambient.UserMode` _
 property returns the incorrect value of True when the control is being run in _
 "Design Mode" (on the form editor). This would cause the design mode controls _
 to be subclassed and crashes the IDE. To stop this, the variable below will _
 always be False when the controls are running in Design Mode. Set `UserMode` _
 to True in your `Sub Main()` to tell the controls it's okay to subclass. _
 (`Sub Main()` will only be run when you your app runs, not during design time)
Public UserMode As Boolean

'The default measurement (px) to base control layout around. _
 Use `blu.Xpx(blu.Metric)` / `blu.Ypx(blu.Metric)` to get it in Twips
Public Const Metric As Long = 32

'The default colour palette for our controls
Public Const BaseColour As Long = vbWhite
Public Const BaseHoverColour As Long = &HEEEEEE
Public Const TextColour As Long = &H999999
Public Const TextHoverColour As Long = &H666666
Public Const ActiveColour As Long = &HFFAF00
Public Const InertColour As Long = &HFFEABA

'The close control box button is red unlike the others
Public Const CloseHoverColour As Long = &H4343E0
Public Const ClosePressColour As Long = &H5050C7 '&H3D3D99

'Used in converting colours to Hue / Saturation / Lightness
Public Type HSL
    Hue As Long
    Saturation As Long
    Luminance As Long
End Type

'Public Enums _
 --------------------------------------------------------------------------------------
'The Blu ActiveX controls use these to define friendly names for some properties

Public Enum bluORIENTATION
    Horizontal = 0
    VerticalUp = 1
    VerticalDown = 2
End Enum

Public Enum bluSTATE
    Inactive = 0
    Active = 1
End Enum

Public Enum bluSTYLE
    Normal = 0
    Invert = 1
End Enum

Public Enum bluTRUNCATE
    NoTruncation = 0
    EndElipsis = 1
    MiddleElipsis = 2
End Enum

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////
'Yes, you can actually place properties in a module! Why would you want to do this? _
 Saves having to store a global variable and use a function to init the value

'PROPERTY DropShadows : If the "Show shadows under windows" option is on _
 ======================================================================================
Public Property Get DropShadows() As Boolean
    Dim Result As BOOL
    Call user32_SystemParametersInfo(SPI_GETDROPSHADOW, 0, Result, 0)
    Let DropShadows = (Result = API_TRUE)
End Property

'PROPERTY InIDE : Are we running the code from the Visual Basic IDE? _
 ======================================================================================
Public Property Get InIDE() As Boolean
    On Error GoTo Err_True
    
    'Do something that only faults in the IDE
    Debug.Print 1 \ 0
    InIDE = False
    Exit Property

Err_True:
    InIDE = True
End Property

'PROPERTY IsHighContrastMode : If high contrast mode is on _
 ======================================================================================
'NOTE: This is not the same thing as the high contrast theme -- on Windows XP. _
 The user might use a high contrast theme without having high contrast mode on. _
 On Vista and above the high contrast mode is automatically enabled when a high _
 contrast theme is selected: _
 <blogs.msdn.com/b/oldnewthing/archive/2008/12/03/9167477.aspx>
Public Property Get IsHighContrastMode() As Boolean
    'prepare the structure to hold the information about high contrast mode
    Dim Info As HIGHCONTRAST
    Let Info.SizeOfMe = Len(Info)
    'Get the information, passing our structure in
    If user32_SystemParametersInfo( _
        SPI_GETHIGHCONTRAST, Info.SizeOfMe, Info, 0 _
    ) = API_TRUE Then
        'Determine if the bit is set for high contrast mode on/off
        Let IsHighContrastMode = (Info.Flags And HCF_HIGHCONTRASTON) <> 0
    End If
End Property

'PROPERTY WheelScrollLines : The number of lines to scroll when the mouse wheel rolls _
 ======================================================================================
Public Property Get WheelScrollLines() As Long
    Call user32_SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, WheelScrollLines, 0)
    If WheelScrollLines <= 0 Then WheelScrollLines = 3
End Property

'PROPERTY WheelScrollChars : The number of characters to scroll with horizontal wheel _
 ======================================================================================
Public Property Get WheelScrollChars() As Long
    Call user32_SystemParametersInfo(SPI_GETWHEELSCROLLCHARS, 0, WheelScrollChars, 0)
    If WheelScrollChars <= 0 Then WheelScrollChars = 3
End Property

'PROPERTY WindowsVersion : As a Kernel number, i.e. 6.0 = Vista, 6.1 = "7", 6.2 = "8" _
 ======================================================================================
Public Property Get WindowsVersion() As Single
    'NOTE: If the app is in compatibility mode, this will return the compatible _
     Windows version, not the actual version; but that's fine with me
    Dim VersionInfo As OSVERSIONINFO
    Let VersionInfo.SizeOfMe = Len(VersionInfo)
    If kernel32_GetVersionEx(VersionInfo) = API_TRUE Then
        Let WindowsVersion = _
            CSng(VersionInfo.MajorVersion & "." & VersionInfo.MinorVersion)
    End If
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'DrawText : Shared routine for drawing text, used by bluLabel/Button/Tab &c. _
 ======================================================================================
Public Sub DrawText( _
    ByVal hndDeviceContext As Long, _
    ByRef BoundingBox As RECT, _
    ByVal Text As String, _
    ByVal Colour As OLE_COLOR, _
    Optional ByVal Alignment As VBRUN.AlignmentConstants = vbLeftJustify, _
    Optional ByVal Orientation As bluORIENTATION = Horizontal, _
    Optional ByVal FontName As String = "Arial", _
    Optional ByVal FontSizePx As Long = 15 _
)
    'Create and set the font: _
     ----------------------------------------------------------------------------------
    'Create the GDI font object that describes our font properties
    Dim hndFont As Long
    Let hndFont = gdi32_CreateFont( _
        Height:=FontSizePx, Width:=0, _
        Escapement:=0, Orientation:=0, _
        Weight:=FW_NORMAL, Italic:=API_FALSE, Underline:=API_FALSE, _
        StrikeOut:=API_FALSE, CharSet:=DEFAULT_CHARSET, _
        OutputPrecision:=OUT_DEFAULT_PRECIS, ClipPrecision:=CLIP_DEFAULT_PRECIS, _
        Quality:=DEFAULT_QUALITY, PitchAndFamily:=VARIABLE_PITCH Or FF_DONTCARE, _
        Face:=FontName _
    )

    'Select the font (remembering the previous object selected to clean up later)
    Dim hndOld As Long
    Let hndOld = gdi32_SelectObject(hndDeviceContext, hndFont)
    
    'The `DrawText` API doesn't work with the position set by `SetTextAlign`, _
     so we ensure it's set to a safe, non-interfering value
    Call gdi32_SetTextAlign( _
        hndDeviceContext, TA_TOP Or TA_LEFT Or TA_NOUPDATECP _
    )
    
    'Rotate the text? _
     ----------------------------------------------------------------------------------
    'Made possible with directions from: <edais.mvps.org/Tutorials/GDI/DC/DCch8.html>
    If Orientation <> Horizontal Then
        'Determine the centre point of the bounding box before we begin to modify it
        Dim Centre As POINT
        Let Centre.X = BoundingBox.Left + (BoundingBox.Right - BoundingBox.Left) \ 2
        Let Centre.Y = BoundingBox.Top + (BoundingBox.Bottom - BoundingBox.Top) \ 2
        
        'The button is already in a vertical shape, but we want to rotate a horizontal _
         piece of text, so we have to swap the dimensions of the button to begin with
        Call user32_SetRect( _
            BoundingBox, _
            BoundingBox.Left, BoundingBox.Top, BoundingBox.Bottom, BoundingBox.Right _
        )
        'In addition to that, we also need to position our text with its center at _
         0,0, instead of the top-left corner, so that when we rotate, the text stays _
         centered and doesn't swing off out of place
        Call user32_OffsetRect( _
            BoundingBox, -BoundingBox.Right \ 2, -BoundingBox.Bottom \ 2 _
        )
        
        'Now we need to move the origin point (0,0) to the centre of the button _
         so that the rotated text obviously appears in the center of the button _
         whilst the rotation occurs around the centrepoint of the text
        Dim Org As POINT
        Call gdi32_SetViewportOrgEx(hndDeviceContext, Centre.X, Centre.Y, Org)
        
        'In order to use Get/SetWorldTransform we have to make this call
        Dim OldGM As Long
        Let OldGM = gdi32_SetGraphicsMode(hndDeviceContext, GM_ADVANCED)
        
        'Now calculate the rotation
        Const Pi As Single = 3.14159
        Dim RotAng As Single
        Let RotAng = IIf(Orientation = VerticalDown, -90, 90)
        Dim RotRad As Single
        Let RotRad = (RotAng / 180) * Pi
        
        'Read any current transform from the device context
        Dim OldXForm As XFORM, RotXForm As XFORM
        Call gdi32_GetWorldTransform(hndDeviceContext, OldXForm)
        
        'Define our rotation matrix
        With RotXForm
            Let .eM11 = Cos(RotRad)
            Let .eM21 = Sin(RotRad)
            Let .eM12 = -.eM21
            Let .eM22 = .eM11
        End With
        
        'Apply the matrix -- rotate the world!
        Call gdi32_SetWorldTransform(hndDeviceContext, RotXForm)
    End If
    
    'Draw the text! _
     ----------------------------------------------------------------------------------
    'Select the colour of the text
    Call gdi32_SetTextColor( _
        hndDeviceContext, OleTranslateColor(Colour) _
    )
    
    'Add a little padding either side
    With BoundingBox
        Let .Left = .Left + 8
        Let .Right = .Right - 8
    End With

    'Now just paint the text
    Call user32_DrawText( _
        hndDeviceContext:=hndDeviceContext, _
        Text:=Text, Length:=Len(Text), _
        BoundingBox:=BoundingBox, _
        Format:=Choose(Alignment + 1, DT.DT_LEFT, DT.DT_RIGHT, DT.DT_CENTER) _
                Or DT_VCENTER Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP _
    )
    
    'Clean up: _
     ----------------------------------------------------------------------------------
    'If we rotated the text, we need to do some additional clean up
    If Orientation <> Horizontal Then
        'Restore the previous world transform
        Call gdi32_SetWorldTransform(hndDeviceContext, OldXForm)
        'Switch back to the previous graphics mode
        Call gdi32_SetGraphicsMode(hndDeviceContext, OldGM)
        'Return the origin point (0,0) back to the upper-left corner
        Call gdi32_SetViewportOrgEx(hndDeviceContext, Org.X, Org.Y, Org)
    End If
    
    'Select the previous object into the DC (i.e. unselect the font)
    Call gdi32_SelectObject(hndDeviceContext, hndOld)
    Call gdi32_DeleteObject(hndFont)
End Sub

'GetMousePos_Screen : Get the mouse coordinates on the screen _
 ======================================================================================
Public Function GetMousePos_Screen() As POINT
    Call user32_GetCursorPos(GetMousePos_Screen)
End Function

'GetMousePos_Window : Get the mouse position within (relative) a window _
 ======================================================================================
Public Function GetMousePos_Window(ByVal hWnd As Long) As POINT
    Let GetMousePos_Window = GetMousePos_Screen()
    Call user32_ScreenToClient(hWnd, GetMousePos_Window)
End Function

'GetParentForm : Recurses through the parent objects until we hit the top form _
 ======================================================================================
Public Function GetParentForm( _
    ByRef StartWith As Object, _
    Optional ByVal GetMDIParent As Boolean = False _
) As Object
    'Begin with the provided starting object
    Set GetParentForm = StartWith
    'Walk up the parent tree as far as we can
    Do
        On Error GoTo NowCheckMDI
        Set GetParentForm = GetParentForm.Parent
    Loop
NowCheckMDI:
    On Error GoTo Complete
    'Have been asked to negotiate from the MDI child into the MDI parent?
    If GetMDIParent = False Then Exit Function
    
    'There is no built in way to find the MDI parent of a child form, though of _
     course you can only have one MDI form in the app, but I wouldn't want to have to _
     reference that by name here, yours might be named something else. What we do is _
     use Win32 to go up through the "MDIClient" window (that isn't exposed to VB) _
     which acts as the viewport of the MDI form and then up again to hit the MDI form
    If Not TypeOf GetParentForm Is MDIForm Then
        Dim MDIParent_hWnd As Long
        Let MDIParent_hWnd = user32_GetParent( _
            user32_GetParent(GetParentForm.hWnd) _
        )
        'Once we have the handle, check the list of loaded VB forms to find the _
         MDI form it belongs to
        Dim Frm As Object
        For Each Frm In VB.Forms
            If Frm.hWnd = MDIParent_hWnd Then
                Set GetParentForm = Frm
                Exit Function
            End If
        Next
    End If
Complete:
End Function

'GetParentForm_hWnd : Get the window handle of the parent form (or MDI-parent) _
 ======================================================================================
Public Function GetParentForm_hWnd( _
    ByRef StartWith As Object, _
    Optional ByVal GetMDIParent As Boolean = False _
) As Long
    'Acts as a wrapper to the function above so that the callee doesn't have to hold _
     an object reference
    Dim ParentForm As Object
    Set ParentForm = GetParentForm(StartWith, GetMDIParent)
    Let GetParentForm_hWnd = ParentForm.hWnd
End Function

'GetSystemMetric : Sizes for window borders, menus, scroll bars &c. _
 ======================================================================================
Public Function GetSystemMetric(ByVal Metric As SM) As Long
    Let GetSystemMetric = user32_GetSystemMetrics(Metric)
End Function

'HiWord : Get the high-Word (top 16-bits) from a Long (32-bits) _
 ======================================================================================
Public Function HiWord(ByVal Value As Long) As Integer
    'Special thanks to Tanner Helland & PhotoDemon <photodemon.org> _
     for making this negative-safe
    If Value And &H80000000 Then
        Let HiWord = (Value \ 65535) - 1
    Else
        Let HiWord = Value \ 65535
    End If
End Function

'HSLToRGB : Convert Hue, Saturation, Ligthness to (roughly) Red, Green, Blue _
 ======================================================================================
'<www.xbeat.net/vbspeed/c_HSLToRGB.htm>
Public Function HSLToRGB( _
    ByVal Hue As Long, _
    ByVal Saturation As Long, _
    ByVal Luminance As Long _
) As Long
    'by Donald (Sterex 1996), donald@xbeat.net, 20011124
    Dim r As Long, G As Long, b As Long
    Dim lMax As Long, lMid As Long, lMin As Long
    Dim q As Single

    lMax = (Luminance * 255) / 100
  
    If Saturation > 0 Then

        lMin = (100 - Saturation) * lMax / 100
        q = (lMax - lMin) / 60
        
        Select Case Hue
            Case 0 To 60
                lMid = (Hue - 0) * q + lMin
                r = lMax: G = lMid: b = lMin
            Case 60 To 120
                lMid = -(Hue - 120) * q + lMin
                r = lMid: G = lMax: b = lMin
            Case 120 To 180
                lMid = (Hue - 120) * q + lMin
                r = lMin: G = lMax: b = lMid
            Case 180 To 240
                lMid = -(Hue - 240) * q + lMin
                r = lMin: G = lMid: b = lMax
            Case 240 To 300
                lMid = (Hue - 240) * q + lMin
                r = lMid: G = lMin: b = lMax
            Case 300 To 360
                lMid = -(Hue - 360) * q + lMin
                r = lMax: G = lMin: b = lMid
        End Select
        HSLToRGB = b * &H10000 + G * &H100& + r
    Else
        HSLToRGB = lMax * &H10101
    End If
End Function

'InitCommonControls : Enable Windows themeing on controls (application wide) _
 ======================================================================================
Public Function InitCommonControls(Optional ByVal Types As ICC = ICC_STANDARD_CLASSES) As Boolean
    'NOTE: Call this procedure from your `Sub Main` before loading any forms
    
    'NOTE: Your app must have a manifest file (either internal or external) in order _
     for this to work, see the web page below for instructions
    'Thanks goes to LaVolpe and his manifest creator for the help _
     <www.vbforums.com/showthread.php?606736-VB6-XP-Vista-Win7-Manifest-Creator>
    
    'WARNING: If your app never displays any common controls (a form containing them _
     doesn't get loaded by the user), then YOUR EXE WILL CRASH ON EXIT. If there is _
     any chance a user can start your app and close it before any common controls _
     have been loaded then to prevent crashing you MUST either:
    '1. Place a hidden ComboBox on any form that has no other common controls
    '2. Delay calling this function until before a form containing common controls _
        is loaded
    
    'Prepare the structure used for `InitCommonControlsEx`
    Dim ControlTypes As INITCOMMONCONTROLSEX
    Let ControlTypes.SizeOfMe = Len(ControlTypes)
    Let ControlTypes.Flags = Types
    
    On Error Resume Next
    Dim hndModule As Long
    'LaVolpe tells us that XP can crash if we have user controls when we call _
     `InitCommonControlsEx` unless we pre-emptively connect to Shell32
    Let hndModule = kernel32_LoadLibrary("shell32.dll")
    'Return whether control initialisation was successful or not
    Let InitCommonControls = (comctl32_InitCommonControlsEx(ControlTypes) = API_TRUE)
    'Free the reference to Shell32
    If hndModule <> 0 Then Call kernel32_FreeLibrary(hndModule)
End Function

'LoWord : Get the low-Word (bottom 16-bits) from a Long (32-bits) _
 ======================================================================================
Public Function LoWord(ByVal Value As Long) As Integer
    'Special thanks to Tanner Helland & PhotoDemon <photodemon.org> _
     for making this negative-safe
    If Value And &H8000& Then
        Let LoWord = &H8000 Or (Value And &H7FFF&)
    Else
        Let LoWord = Value And &HFFFF&
    End If
End Function

'Max : Limit a number to a maximum value _
 ======================================================================================
Public Function Max(ByVal InputNumber As Long, Optional ByVal Maximum As Long = 2147483647) As Long
    Let Max = IIf(InputNumber > Maximum, Maximum, InputNumber)
End Function

'Min : Limit a number to a minimum value _
 ======================================================================================
Public Function Min(ByVal InputNumber As Long, Optional ByVal Minimum As Long = 0) As Long
    Let Min = IIf(InputNumber < Minimum, Minimum, InputNumber)
End Function

'NotZero : Ensure a number is not zero. Useful when dividing by an unknown factor _
 ======================================================================================
Public Function NotZero(ByVal InputNumber As Long, Optional ByVal AtLeast As Long = 1) As Long
    'This is different from Min / Max because it allows you to handle +/- numbers
    If InputNumber = 0 Then Let NotZero = AtLeast Else Let NotZero = InputNumber
End Function

'OLETranslate : Translate an OLE color to an RGB Long _
 ======================================================================================
Public Function OleTranslateColor(ByVal Colour As OLE_COLOR) As Long
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If olepro32_OleTranslateColor( _
        OLEColour:=Colour, hndPalette:=0, ptrColour:=OleTranslateColor _
    ) Then Let OleTranslateColor = vbWhite
End Function

'Range : Limit a number to a minimum and maximum value _
 ======================================================================================
Public Function Range( _
    ByVal InputNumber As Long, _
    Optional ByVal Maximum As Long = 2147483647, _
    Optional ByVal Minimum As Long = -2147483648# _
) As Long
    Let Range = Max(Min(InputNumber, Minimum), Maximum)
End Function

'RGBToHSL : Convert Red, Green, Blue to Hue, Saturation, Lightness _
 ======================================================================================
'<www.xbeat.net/vbspeed/c_RGBToHSL.htm>
Public Function RGBToHSL(ByVal RGBValue As Long) As HSL
    'by Paul - wpsjr1@syix.com, 20011120
    Dim r As Long, G As Long, b As Long
    Dim lMax As Long, lMin As Long
    Dim q As Single
    Dim lDifference As Long
    Static Lum(255) As Long
    Static QTab(255) As Single
    Static init As Long
    
    If init = 0 Then
        For init = 2 To 255 ' 0 and 1 are both 0
            Lum(init) = init * 100 / 255
        Next
        For init = 1 To 255
            QTab(init) = 60 / init
        Next
    End If
    
    r = RGBValue And &HFF
    G = (RGBValue And &HFF00&) \ &H100&
    b = (RGBValue And &HFF0000) \ &H10000
    
    If r > G Then
        lMax = r: lMin = G
    Else
        lMax = G: lMin = r
    End If
    If b > lMax Then
        lMax = b
    ElseIf b < lMin Then
        lMin = b
    End If
    
    RGBToHSL.Luminance = Lum(lMax)
    
    lDifference = lMax - lMin
    If lDifference Then
        'Do a 65K 2D lookup table here for more speed if needed
        RGBToHSL.Saturation = (lDifference) * 100 / lMax
        q = QTab(lDifference)
        Select Case lMax
            Case r
                If b > G Then
                    RGBToHSL.Hue = q * (G - b) + 360
                Else
                    RGBToHSL.Hue = q * (G - b)
                End If
            Case G
                RGBToHSL.Hue = q * (b - r) + 120
            Case b
                RGBToHSL.Hue = q * (r - G) + 240
        End Select
    End If
End Function

'RoundUp : Always round a number upwards _
 ======================================================================================
Public Function RoundUp(ByVal InputNumber As Double) As Double
    If Int(InputNumber) = InputNumber _
        Then Let RoundUp = InputNumber _
        Else Let RoundUp = Int(InputNumber) + 1
End Function

'SetIcon : Use a 32-bit icon from the compiled in resource file _
 ======================================================================================
'This function has been adapted from this article & code by Steve McMahon: _
 <www.vbaccelerator.com/home/VB/Tips/Setting_the_App_Icon_Correctly/article.asp>
'It relies upon the icon being compiled into the EXE using the .res file. _
 See the RES folder for scripts and files that compile the icons into the .res file
Public Sub SetIcon( _
    ByVal hndWindow As Long, _
    ByVal IconResName As String, _
    Optional ByVal SetAsAppIcon As Boolean = True _
)
    Const WM_SETICON As Long = &H80
    Const ICON_SMALL As Long = 0
    Const ICON_BIG As Long = 1
    
    'We can't load icons from the EXE when running from the IDE, obviously! _
     As a cheap fall-back we'll load the icon from the RES file, but you won't get _
     the quality 32-bit icons, so expect some roughness
    If InIDE = True Then
        'Find which form this handle belongs to
        Dim VBForm As VB.Form
        For Each VBForm In VB.Forms
            If VBForm.hWnd = hndWindow Then
                'Set the icon from the resource file, though this will likely load _
                 the 16x16 256-colour icon
                Set VBForm.Icon = VB.LoadResPicture( _
                    IconResName, VBRUN.LoadResConstants.vbResIcon _
                )
                Exit For
            End If
        Next
        Exit Sub
    End If
    
    'VB6 has a hidden window that acts as some kind of persistence / controller for _
     the whole app. We need to find this to set an app-wide icon (e.g. Alt+Tab window)
    If SetAsAppIcon = True Then
        Dim hndParent As Long
        Let hndParent = hndWindow
        Dim hndVB6 As Long
        Let hndVB6 = hndWindow
        Do While Not (hndParent = 0)
            Let hndParent = user32_GetWindow(hndParent, GW_OWNER)
            If Not (hndParent = 0) Then hndVB6 = hndParent
        Loop
    End If
    
    'Do the actual loading of the icon, large size (usually 32 or 48px)
    Dim hndIconLarge As Long
    Let hndIconLarge = user32_LoadImage( _
        App.hInstance, IconResName, IMAGE_ICON, _
        GetSystemMetric(SM_CXICON), GetSystemMetric(SM_CYICON), _
        LR_SHARED _
    )
    'Assign the large icon:
    If SetAsAppIcon = True Then
        Call user32_SendMessage(hndVB6, WM_SETICON, ICON_BIG, hndIconLarge)
    End If
    Call user32_SendMessage(hndWindow, WM_SETICON, ICON_BIG, hndIconLarge)
    
    'Load the small icon size (usually 16px)
    Dim hndIconSmall As Long
    Let hndIconSmall = user32_LoadImage( _
        App.hInstance, IconResName, IMAGE_ICON, _
        GetSystemMetric(SM_CXSMICON), GetSystemMetric(SM_CYSMICON), _
        LR_SHARED _
    )
    'Assign the small icon:
    If SetAsAppIcon = True Then
        Call user32_SendMessage(hndVB6, WM_SETICON, ICON_SMALL, hndIconSmall)
    End If
    Call user32_SendMessage(hndWindow, WM_SETICON, ICON_SMALL, hndIconSmall)
End Sub

'Xpx : Shorthand for a number of horizontal pixels converted to twips _
 ======================================================================================
Public Function Xpx(Optional ByVal px As Long = 1) As Long
    'Yes, we could use `Form.ScaleX (...)` but this doesn't require a form and is _
     shorter to write
    Let Xpx = Screen.TwipsPerPixelX * px
End Function

'Ypx : Shorthand for a number of vertical pixels converted to twips _
 ======================================================================================
Public Function Ypx(Optional ByVal px As Long = 1) As Long
    Let Ypx = Screen.TwipsPerPixelY * px
End Function
