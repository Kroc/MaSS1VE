VERSION 5.00
Begin VB.UserControl bluWebView 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "bluWebView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CONTROL :: bluWebView

'The VB6 WebBrowser control can be a problem when distributing the source code to a _
 project. Differing IE versions and Windows versions can cause the object reference _
 to break and the whole project to fall to pieces. This class avoids having to use _
 an OCX reference by creating the ActiveX control on-demand when the user control is _
 instantiated.

'This was extemely hard to develop!

'Status             INCOMPLETE, DO NOT USE
'Dependencies       None
'Last Updated       06-FEB-14
'Last Update        Fixed the problem with links not working in the browser due to _
                    method of instantiation

'TODO: _
'*  Sink events _
    <www.binaryworld.net/Main/CodeDetail.aspx?CodeId=3682&atlanta=software%20development>
'   -   Option to hide the control until first navigate complete _
        (stop white flicker on blue background)
        
'*  Forward mouse wheel events to the ActiveX control

'/// API DEFS /////////////////////////////////////////////////////////////////////////

'Create a Windows control _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms632680%28v=vs.85%29.aspx>
Private Declare Function user32_CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
    ByVal ExStyle As WS_EX, _
    ByVal ClassName As String, _
    ByVal WindowName As String, _
    ByVal Style As WS, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal hWndParent As Long, _
    ByVal hndMenu As Long, _
    ByVal AppInstance As Long, _
    ByRef Param As Any _
) As Long

'Window styles _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms632600%28v=vs.85%29.aspx>
Private Enum WS
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

'Extended Window Styles _
 <msdn.microsoft.com/en-us/library/windows/desktop/ff700543%28v=vs.85%29.aspx>
Private Enum WS_EX
    'Extended window styles (via `GWL_EXSTYLE`)
    WS_EX_APPWINDOW = &H40000       'Show in taskbar
    WS_EX_CLIENTEDGE = &H200        'Sunken border
    WS_EX_DLGMODALFRAME = &H1       'Double border
    WS_EX_LAYERED = &H80000         'Layered, that is, can be translucent
    WS_EX_STATICEDGE = &H20000      '3D border for items that do not accept user input
    WS_EX_TOOLWINDOW = &H80
    WS_EX_WINDOWEDGE = &H100        'Border with raised edge
End Enum

'<msdn.microsoft.com/en-us/library/windows/desktop/ms632682%28v=vs.85%29.aspx>
Private Declare Function user32_DestroyWindow Lib "user32" Alias "DestroyWindow" ( _
    ByVal hndWindow As Long _
) As BOOL

'--------------------------------------------------------------------------------------

'"Initializes ATL's control hosting code" _
 <msdn.microsoft.com/en-us/library/d5f8cs41.aspx>
Private Declare Function atl_AtlAxWinInit Lib "atl" Alias "AtlAxWinInit" () As BOOL

'Create an ActiveX control in a container _
 <msdn.microsoft.com/en-us/library/da181h29%28v=vs.90%29.aspx>
Private Declare Function atl_AtlAxCreateControl Lib "atl" Alias "AtlAxCreateControl" ( _
    ByVal strPtrControlClass As Long, _
    ByVal hndWindow As Long, _
    ByVal ptrStream As Long, _
    ByRef ptrIUnknown As Long _
) As Long

'Attach an ActiveX control container to an existing window _
 <msdn.microsoft.com/en-us/library/d91055eh.aspx>
Private Declare Function atl_AtlAxAttachControl Lib "atl" Alias "AtlAxAttachControl" ( _
    ByVal ptrControl As Long, _
    ByVal hWnd As Long, _
    ByRef ptrIUnknownContainer As Long _
) As HRESULT

'Get the pointer to the IUnknown (a COM parent of VB's Object type), we can use this _
 to cast the IUnknown pointer to a VB Object to get our late-bound reference to the _
 web-browser's IWebBrowser2 interface <msdn.microsoft.com/en-us/library/c9c7hb2f.aspx>
Private Declare Function atl_AtlAxGetControl Lib "atl" Alias "AtlAxGetControl" ( _
    ByVal hWnd As Long, _
    ByRef ptrIUnknown As Long _
) As Long

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Private Declare Sub kernel32_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef ptrDestination As Any, _
    ByRef ptrSource As Any, _
    ByVal Length As Long _
)

'--------------------------------------------------------------------------------------

'Position a window -- used to fit the web browser control to the user control _
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

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'All COM objects implement an "IUnknown" interface, and since VB6 is COM-based, _
 we can coerce the IUnknown pointer we get from `AtlAxGetControl` into a VB Object _
 that implicitly acts as a late-bound web browser control! (`IWebBrowser2`)
Private ptrIUnknown As Long

Private ptrIUnknownContainer As Long

'This will hold our reference to the web browser control's interface. _
 This class will implement what it can of `IWebBrowser2` for convenience
Private IWebBrowser2 As Object

''<msdn.microsoft.com/en-us/library/aa768283%28v=vs.85%29.aspx>
'Private WithEvents DWebBrowserEvents2 As VBControlExtender

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'This will be the handle to the Internet Explorer control, _
 the old web browser control doesn't give you this!
Private My_hWnd As Long

'What page to load as soon as the web browser is instantiated
Private My_StartURL As String

'You can set a background colour on the control to reduce flicker between _
 instantiation and navigation to the start URL
Private My_BackColor As OLE_COLOR

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY BackColor _
 ======================================================================================
Public Property Get BackColor() As OLE_COLOR: Let BackColor = My_BackColor: End Property
Public Property Let BackColor(ByVal Color As OLE_COLOR)
    Let My_BackColor = Color
    Let UserControl.BackColor = My_BackColor
    Call UserControl.PropertyChanged("BackColor")
End Property

'PROPERTY StartURL _
 ======================================================================================
Public Property Get StartURL() As String: Let StartURL = My_StartURL: End Property
Public Property Let StartURL(ByVal URL As String)
    Let My_StartURL = URL
    Call UserControl.PropertyChanged("StartURL")
End Property

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Let My_StartURL = PropBag.ReadProperty(Name:="StartURL", DefaultValue:="http://about:blank")
    Let Me.BackColor = PropBag.ReadProperty(Name:="BackColor", DefaultValue:=vbWhite)
    
    'DON'T try to create the web browser control in design-time or when compiling
    If blu.UserMode = False Then Exit Sub
    
    'Clean up should this called twice
    Call UserControl_Terminate
    
    
    Call atl_AtlAxWinInit
    
    'Create the ActiveX control container. We use this particular method, instead of _
     `AtlAxCreateControl` because for some reason that causes hyperlinks to not work
    'Thanks goes to <www.aivosto.com/vbtips/stringopt2.html#API> for the `StrPtr` trick
    Let My_hWnd = user32_CreateWindowEx( _
        0, "AtlAxWin", StrPtr(My_StartURL), WS_CHILD, _
        0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        UserControl.hWnd, ByVal 0, App.hInstance, ByVal 0 _
    )
    
    'Get the pointer to the ActiveX control created
    Call atl_AtlAxGetControl(My_hWnd, ptrIUnknown)
    'Convert the pointer to a VB Object
    Set IWebBrowser2 = ObjectFromIUnknownPtr(ptrIUnknown)
    
    'Configure the browser view; in all likelihood this is unnecessary
    '"In theater mode, the object's main window fills the entire screen and displays _
      a toolbar that has a minimal set of navigational buttons. A status bar is also _
      provided in the upper-right corner of the screen. Explorer bars, such as History _
      and Favorites , are displayed as an autohide pane on the left edge of the screen _
      in theater mode." <msdn.microsoft.com/en-us/library/aa768273%28v=vs.85%29.aspx>
    Let IWebBrowser2.TheaterMode = False
    Let IWebBrowser2.toolbar = False
    Let IWebBrowser2.AddressBar = False
    Let IWebBrowser2.MenuBar = False
    Let IWebBrowser2.Resizable = False
    Let IWebBrowser2.Silent = True
    Let IWebBrowser2.StatusBar = False
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize()
    'Fit the ActiveX control to the user control
    Call user32_SetWindowPos( _
        My_hWnd, HWND_TOP, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
        SWP_SHOWWINDOW _
    )
End Sub

'CONTROL Terminate _
 ======================================================================================
Private Sub UserControl_Terminate()
    'Detatch our object reference to the interface
    Set IWebBrowser2 = Nothing
'    Set DWebBrowserEvents2 = Nothing
    
    'Get rid of the ActiveX control
    If My_hWnd <> 0 Then Call user32_DestroyWindow(My_hWnd)
End Sub

'CONTROL WriteProperties _
 ======================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(Name:="StartURL", Value:=My_StartURL, DefaultValue:="http://about:blank")
    Call PropBag.WriteProperty(Name:="BackColor", Value:=My_BackColor, DefaultValue:=vbWhite)
End Sub

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

Public Function Navigate(ByVal URL As String)
    Call IWebBrowser2.Navigate(URL)
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

Private Function ObjectFromIUnknownPtr(ByVal ptr As Long) As Object
    Dim objUnknown As IUnknown
    Call kernel32_RtlMoveMemory(objUnknown, ptr, 4&)
    Set ObjectFromIUnknownPtr = objUnknown
    Call kernel32_RtlMoveMemory(objUnknown, 0&, 4&)
End Function

'Public Function GetExtendedControl(oCtl As IUnknown) As VBControlExtender
'    Dim pOleObject      As IOleObject
'    Dim pOleControlSite As IOleControlSite
'
'    On Error Resume Next
'    Set pOleObject = oCtl
'    Set pOleControlSite = pOleObject.GetClientSite
'    Set GetExtendedControl = pOleControlSite.GetExtendedControl
'    On Error GoTo 0
'End Function
