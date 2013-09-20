Attribute VB_Name = "WIN32"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: WIN32

'API calls into the guts of Windows

'Status             In Flux
'Dependencies       Lib.bas
'Last Updated       20-SEP-13
'Last Update        Added `SetIcon` to load 32-bit icons from the EXE

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
Public Type SIZE
    Width As Long
    Height As Long
End Type

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

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Public Declare Sub kernel32_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef ptrDestination As Any, _
    ByRef ptrSource As Any, _
    ByVal Length As Long _
)

'DLL LOADING _
 --------------------------------------------------------------------------------------

'The above can apparently be buggy so this is used as a fallback _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms684175%28v=vs.85%29.aspx>
Private Declare Function kernel32_LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal FileName As String _
) As Long

'Free the resource associated with the above call _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms683152%28v=vs.85%29.aspx>
Private Declare Function kernel32_FreeLibrary Lib "kernel32" Alias "FreeLibrary" ( _
    ByVal hndModule As Long _
) As BOOL

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

'WINDOWS SYSTEM INFORMATION _
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

'Get the location of a special folder, e.g. "My Documents", "System32" &c. _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb762181%28v=vs.85%29.aspx>
Private Declare Function shfolder_SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" ( _
    ByVal hWndOwner As Long, _
    ByVal Folder As CSIDL, _
    ByVal Token As Long, _
    ByVal Flags As SHGFP, _
    ByVal Path As String _
) As HRESULT

'Full list with descriptions here: _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb762494%28v=vs.85%29.aspx>
Public Enum CSIDL
    CSIDL_APPDATA = &H1A&       'Application data (roaming), intended for app data
                                 'that should persist with the user between machines
    CSIDL_LOCAL_APPDATA = &H1C& 'Application data specific to the PC (e.g. cache)
    CSIDL_COMMON_APPDATA = &H23 'Application data shared between all users
    
    
    CSIDL_FLAG_CREATE = &H8000& 'OR this with any of the above to create the folder
                                 'if it doesn't exist (e.g. user deleted My Pictures)
End Enum

Private Enum SHGFP
    SHGFP_TYPE_CURRENT = 0      'Retrieve the folder's current path (it may have moved)
    SHGFP_TYPE_DEFAULT = 1      'Get the default path
End Enum

'Get the location of the temporary files folder _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa364992%28v=vs.85%29.aspx>
Private Declare Function kernel32_GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
    ByVal BufferLength As Long, _
    ByVal Buffer As String _
) As Long

'Convert a system color (such as "button face" or "inactive window") to a RGB value _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms694353%28v=vs.85%29.aspx>
Private Declare Function olepro32_OleTranslateColor Lib "olepro32" Alias "OleTranslateColor" ( _
    ByVal OLEColour As OLE_COLOR, _
    ByVal hndPalette As Long, _
    ByRef ptrColour As Long _
) As Long

'WINDOW MANIPULATION _
 --------------------------------------------------------------------------------------

'<msdn.microsoft.com/en-us/library/windows/desktop/ms633510%28v=vs.85%29.aspx>
Public Declare Function user32_GetParent Lib "user32" Alias "GetParent" ( _
    ByVal hndWindow As Long _
) As Long

'We'll use this in `SetIcon` to find VB6's hidden top-level window _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633515%28v=vs.85%29.aspx>
Private Declare Function user32_GetWindow Lib "user32" Alias "GetWindow" ( _
    ByVal hndWindow As Long, _
    ByVal Command As GW _
) As Long

Private Enum GW
    GW_OWNER = 4
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

'Shift a rectangle's position _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162746%28v=vs.85%29.aspx>
Public Declare Function user32_OffsetRect Lib "user32" Alias "OffsetRect" ( _
    ByRef RectToMove As RECT, _
    ByVal X As Long, _
    ByVal Y As Long _
) As BOOL

'Get mouse position in screen coordinates _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms648390%28v=vs.85%29.aspx>
Public Declare Function user32_GetCursorPos Lib "user32" Alias "GetCursorPos" ( _
    ByRef Pos As POINT _
) As BOOL

'Convert an X/Y point on the screen to local coordinates of a window area _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162952%28v=vs.85%29.aspx>
Public Declare Function user32_ScreenToClient Lib "user32" Alias "ScreenToClient" ( _
    ByVal hndWindow As Long, _
    ByRef ScreenPoint As POINT _
) As BOOL

'Is a point in the rectangle? e.g. check if mouse is within a window _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162882%28v=vs.85%29.aspx>
Public Declare Function user32_PtInRect Lib "user32" Alias "PtInRect" ( _
    ByRef InRect As RECT, _
    ByVal X As Long, _
    ByVal Y As Long _
) As BOOL

'Send a message from one window to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms644950%28v=vs.85%29.aspx>
Public Declare Function user32_SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hndWindow As Long, _
    ByVal Message As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

'Sends a message to another window, but doesn't wait for a return value _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms644944%28v=vs.85%29.aspx>
Public Declare Function user32_PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hndWindow As Long, _
    ByVal Message As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As BOOL

'GDI DRAWING: _
 --------------------------------------------------------------------------------------

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
  values:
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
    ByVal Mode As Long _
) As Long

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

'I need to investigate the actual effectiveness of this lot (preventing repaints to _
 reduce flicker). I've fixed flicker during resizing, but there are instances - mostly _
 when switching level, that several things have to repaint close to each other and I'd _
 like to hold off redrawing the window entirely until the whole process is complete
'--------------------------------------------------------------------------------------
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
'Private Const WM_SETREDRAW as Long = &HB
'Private Const RDW_INVALIDATE as Long = &H1
'Private Const RDW_INTERNALPAINT as Long = &H2
'Private Const RDW_UPDATENOW as Long = &H100
'Private Const RDW_ALLCHILDREN as Long = &H80
'
'Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
'
'Public Function LockRedraw(ByVal hWnd As Long)
'    Call SendMessage(hWnd, WM_SETREDRAW, 0&, 0&)
'End Function
'
'Public Function UnlockRedraw(ByVal hWnd As Long)
'    Dim r As RECT
'    Call SendMessage(hWnd, WM_SETREDRAW, 1, 0&)
'    Call user32_GetClientRect(hWnd, r)
'    'http://www.xtremevbtalk.com/showthread.php?t=189480
'    Call RedrawWindow(hWnd, r, 0&, RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_UPDATENOW Or RDW_ALLCHILDREN)
'    Call InvalidateRect(hWnd, 0&, 0)
'End Function

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

'GetSpecialFolder : Get the path to a system folder, e.g. AppData _
 ======================================================================================
Public Function GetSpecialFolder(ByVal Folder As CSIDL) As String
    'Return null should this fail
    Let GetSpecialFolder = vbNullString
    
    'Fill a buffer to receive the path
    Dim Result As String
    Let Result = String$(260, " ")
    'Attempt to get the special folder path, creating it if it doesn't exist _
     (e.g. the user deleted the "My Pictures" folder)
    If shfolder_SHGetFolderPath( _
        0&, Folder Or CSIDL_FLAG_CREATE, 0&, SHGFP_TYPE_CURRENT, Result _
    ) = S_OK Then
        'The string will be null-terminated; find the end and trim, _
         also ensure it always ends in a slash (this can be inconsistent)
        Let GetSpecialFolder = Lib.EndSlash(Left$( _
            Result, InStr(1, Result, vbNullChar) - 1 _
        ))
    End If
End Function

'GetSystemMetric : Sizes for window borders, menus, scroll bars &c. _
 ======================================================================================
Public Function GetSystemMetric(ByVal Metric As SM) As Long
    Let GetSystemMetric = user32_GetSystemMetrics(Metric)
End Function

'GetTemporaryFolder : Get the path to the temporary files folder _
 ======================================================================================
Public Function GetTemporaryFolder() As String
    'Return null should this fail
    Let GetTemporaryFolder = vbNullString
    
    'Fill a buffer to receive the path
    Dim Result As String
    Let Result = String$(260, " ")
    
    If kernel32_GetTempPath(Len(Result), Result) > 0 Then
        'The string will be null-terminated; find the end and trim, _
         also ensure it always ends in a slash (this can be inconsistent)
        Let GetTemporaryFolder = Lib.EndSlash(Left$( _
            Result, InStr(1, Result, vbNullChar) - 1 _
        ))
    End If
End Function

'GetTemporaryFile : Get a unique file name in the temporary files folder _
 ======================================================================================
Public Function GetTemporaryFile() As String
    'The Windows `GetTempFileName` API is not reliable, it has a limit of 65'535 files _
     which could be hit if we generate a lot and the user doesn't clear their cache. _
     Instead we'll use a timestamp that should be sufficient enough
        
    'Generate a unique file name
    Let GetTemporaryFile = WIN32.GetTemporaryFolder _
        & App.EXEName & "_" _
        & Year(Now) _
        & Right("0" & Month(Now), 2) _
        & Right("0" & Day(Now), 2) _
        & Right("0" & Hour(Now), 2) _
        & Right("0" & Minute(Now), 2) _
        & Right("0" & Second(Now), 2) _
        & "_" & Timer _
        & ".tmp"
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

'OLETranslate : Translate an OLE color to an RGB Long _
 ======================================================================================
Public Function OLETranslateColor(ByVal Colour As OLE_COLOR) As Long
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If olepro32_OleTranslateColor( _
        OLEColour:=Colour, hndPalette:=0, ptrColour:=OLETranslateColor _
    ) Then Let OLETranslateColor = vbWhite
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
    Const WM_SETICON = &H80
    Const ICON_SMALL = 0
    Const ICON_BIG = 1
    
    'We can't load icons from the EXE when running from the IDE, obviously! _
     As a cheap fall-back we'll load the icon from the RES file, but you won't get _
     the quality 32-bit icons, so expect some roughness
    If Run.InIDE = True Then
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
        Next VBForm
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

