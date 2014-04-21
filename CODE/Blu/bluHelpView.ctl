VERSION 5.00
Begin VB.UserControl bluHelpView 
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
Attribute VB_Name = "bluHelpView"
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
'CONTROL :: bluHelpView

'A control intended for viewing local embedded help in HTML format

'The VB6 WebBrowser control can be a problem when distributing the source code to a _
 project. Differing IE versions and Windows versions can cause the object reference _
 to break and the whole project to fall to pieces. This class avoids having to use _
 an OCX reference by creating the ActiveX control on-demand when the user control is _
 instantiated

'Status             INCOMPLETE, DO NOT USE
'Dependencies       None
'Last Updated       09-FEB-14
'Last Update        Moved to an entirely simpler instantiation method :| _
                    Hide browser until first DocumentReady

'TODO: _
'*  Forward mouse wheel events to the ActiveX control

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'<msdn.microsoft.com/en-us/library/aa768283%28v=vs.85%29.aspx>
Private WithEvents WebBrowser As VBControlExtender
Attribute WebBrowser.VB_VarHelpID = -1

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

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
    Let My_StartURL = PropBag.ReadProperty(Name:="StartURL", DefaultValue:="about:blank")
    Let Me.BackColor = PropBag.ReadProperty(Name:="BackColor", DefaultValue:=vbWhite)
    
    'DON'T try to create the web browser control in design-time or when compiling
    If blu.UserMode = False Then Exit Sub
    
    'Clean up should this called twice
    Call UserControl_Terminate
    
    'In order to avoid flicker (especially when a background colour is set) _
     we hide this user control until the starting URL is loaded _
     -- we cannot hide the web browser control because it will not fire events _
        correctly when hidden: <support.microsoft.com/kb/259935>
'    Let UserControl.Extender.Visible = False
    
    'Spawn the web browser ActiveX control
    Set WebBrowser = UserControl.Controls.Add("Shell.Explorer.2", "WebBrowser")
    
    'Configure the browser view; _
     in all likelihood this is unnecessary
    With WebBrowser.Object
        On Error Resume Next
        '"In theater mode, the object's main window fills the entire screen and _
          displays a toolbar that has a minimal set of navigational buttons. _
          A status bar is also provided in the upper-right corner of the screen. _
          Explorer bars, such as History and Favorites, are displayed as an autohide _
          pane on the left edge of the screen in theater mode." _
         "Setting TheaterMode (even to VARIANT_FALSE) resets the values of the _
          IWebBrowser2::AddressBar and IWebBrowser2::ToolBar properties to _
          VARIANT_TRUE. Disable the address bar and toolbars after you set the _
          TheaterMode property." _
          <msdn.microsoft.com/en-us/library/aa768273%28v=vs.85%29.aspx>
        Let .TheaterMode = False
        Let .AddressBar = False
        Let .FullScreen = False
        Let .MenuBar = False
        Let .RegisterAsDropTarget = False
        Let .Resizable = False
        Let .Silent = True
        Let .statusbar = False
        Let .toolbar = False
        On Error GoTo 0
        
        Call .Navigate2(My_StartURL)
    End With
    
    Let WebBrowser.Visible = True
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize()
    'There won't be a control to resize in design mode / when compiling
    If blu.UserMode = False Then Exit Sub
    
    'Fit the WebBrowser control to the user control
'    Call WebBrowser.Move( _
'        0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight _
'    )
End Sub

'CONTROL WriteProperties _
 ======================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(Name:="StartURL", Value:=My_StartURL, DefaultValue:="about:blank")
    Call PropBag.WriteProperty(Name:="BackColor", Value:=My_BackColor, DefaultValue:=vbWhite)
End Sub

'CONTROL Terminate _
 ======================================================================================
Private Sub UserControl_Terminate()
    'Detatch our object reference to the interface / event sink
    Set WebBrowser = Nothing
    
    'Get rid of the WebBrowser control
    On Error Resume Next
    Call UserControl.Controls.Remove("WebBrowser")
    On Error GoTo 0
End Sub

'WEBBROWSER ObjectEvent : When an event occurs from the web browser we created _
 ======================================================================================
Private Sub WebBrowser_ObjectEvent(Info As EventInfo)
    On Error Resume Next
    Debug.Print "* " & Info.Name
    Dim P As EventParameter
    For Each P In Info.EventParameters
        Debug.Print "- " & vbTab & P.Name & ": " & P.Value
    Next
    
    Select Case Info.Name
        Case "DocumentComplete"
'            Let UserControl.Extender.Visible = True
        
        'When the user clicks a link, interject -- we want to send all external links _
         to the default web browser instead
        Case "BeforeNavigate2"
'            Let Info.EventParameters.Item("Cancel") = True
'            Stop
            
    End Select
End Sub

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Navigate : Load a web page _
 ======================================================================================
Public Function Navigate(ByVal URL As String)
    Call WebBrowser.Object.Navigate(URL)
End Function
