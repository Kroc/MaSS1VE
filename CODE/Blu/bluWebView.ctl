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
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CONTROL :: bluWebView

'Embeds an Internet Explorer web browser control -- without a project reference!

'TODO: _
'*  Sink events
'*  Add property for default URL to load on instantiation (i.e. "about:blank")

'/// API DEFS /////////////////////////////////////////////////////////////////////////

'Create an ActiveX control in a container _
 <msdn.microsoft.com/en-us/library/da181h29%28v=vs.90%29.aspx>
Private Declare Function atl_AtlAxCreateControl Lib "atl" Alias "AtlAxCreateControl" ( _
    ByVal strPtrControlClass As Long, _
    ByVal hndWindow As Long, _
    ByVal ptrStream As Long, _
    ByRef ptrIUnknown As Long _
) As Long

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

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'All COM objects implement an "IUnknown" interface, and since VB6 is COM-based, _
 we can coerce the IUnknown pointer we get from `AtlAxGetControl` into a VB Object _
 that implicitly acts as a late-bound web browser control! (`IWebBrowser2`)
Private ptrIUnknown As Long

'This will hold our reference to the web browser control. _
 This class will implement what it can of `IWebBrowser2` for convenience
Private IWebBrowser2 As Object

''<msdn.microsoft.com/en-us/library/aa768283%28v=vs.85%29.aspx>
'Private WithEvents DWebBrowserEvents2 As VBControlExtender

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'This will be the handle to the Internet Explorer control, _
 the old web browser control doesn't give you this!
Private My_hWnd As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If blu.UserMode = False Then Exit Sub
    
    'Use this for events?
    'http://www.binaryworld.net/Main/CodeDetail.aspx?CodeId=3682&atlanta=software%20development
    
    Dim ptrIUnknownControl As Long
    
    'http://www.aivosto.com/vbtips/stringopt2.html#API
    Let My_hWnd = atl_AtlAxCreateControl( _
        StrPtr("http://about:blank"), UserControl.hWnd, 0, ptrIUnknownControl _
    )
    
    Call atl_AtlAxGetControl( _
        UserControl.hWnd, ptrIUnknown _
    )
    
    Set IWebBrowser2 = ObjectFromIUnknownPtr(ptrIUnknown)
End Sub

Private Sub UserControl_Terminate()
    Set IWebBrowser2 = Nothing
'    Set DWebBrowserEvents2 = Nothing
End Sub

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

Public Function Navigate(ByVal URL As String)
    Call IWebBrowser2.Navigate(URL)
End Function

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
