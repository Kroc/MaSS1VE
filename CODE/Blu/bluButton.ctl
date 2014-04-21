VERSION 5.00
Begin VB.UserControl bluButton 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFAF00&
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   97
   ToolboxBitmap   =   "bluButton.ctx":0000
End
Attribute VB_Name = "bluButton"
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
'CONTROL :: BluButton

'Status             Ready to use
'Dependencies       blu.bas, bluMouseEvents.cls (bluMagic.cls)
'Last Updated       31-AUG-13
'Last Update        Colour not being set on new controls

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'Colours
Private My_BaseColour As OLE_COLOR
Private My_ActiveColour As OLE_COLOR

'Appearance
Private My_Orientation As bluORIENTATION        'Horizontal / Vertical Up or Down
Private My_State As bluSTATE                    'If Active, locks the hover effect
Private My_Style As bluSTYLE                    'Back/Fore or Fore/Back colour scheme

'Text
Private My_Caption As String
Private My_Alignment As VBRUN.AlignmentConstants

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'We'll use this to provide MouseIn/Out events
Private WithEvents MouseEvents As bluMouseEvents
Attribute MouseEvents.VB_VarHelpID = -1

'If the button is a hovered state
Private IsHovered As Boolean

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'The label in the button is already subclassed and will provide MouseIn/Out events _
 which we can then expose to the button controller
Event Click()
Event MouseIn()
Event MouseOut()

'Control CLICK: Expose the button click _
 ======================================================================================
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'CONTROL InitProperties _
 ======================================================================================
Private Sub UserControl_InitProperties()
    Let Me.BaseColour = blu.BaseColour
    Let Me.ActiveColour = blu.ActiveColour
    Let Me.Caption = "bluButton"
    Let Me.Orientation = bluORIENTATION.Horizontal
    Let Me.Style = bluSTYLE.Normal
    Let Me.State = bluSTATE.Inactive
End Sub

'CONTROL Paint _
 ======================================================================================
Private Sub UserControl_Paint()
    'Select the background colour
    Call blu.gdi32_SetDCBrushColor( _
        UserControl.hDC, UserControl.BackColor _
    )
    'Get the dimensions of the button
    Dim ClientRECT As RECT
    Call blu.user32_GetClientRect(UserControl.hWnd, ClientRECT)
    'Then use those to fill with the selected background colour
    Call blu.user32_FillRect( _
        UserControl.hDC, ClientRECT, _
        blu.gdi32_GetStockObject(DC_BRUSH) _
    )
    'All the text drawing is shared
    Call blu.DrawText( _
        UserControl.hDC, ClientRECT, My_Caption, UserControl.ForeColor, _
        My_Alignment, My_Orientation _
    )
End Sub

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Let My_ActiveColour = .ReadProperty(Name:="ActiveColour", DefaultValue:=blu.ActiveColour)
        Let My_Alignment = .ReadProperty(Name:="Alignment", DefaultValue:=VBRUN.AlignmentConstants.vbCenter)
        Let My_BaseColour = .ReadProperty(Name:="BaseColour", DefaultValue:=blu.BaseColour)
        Let My_Caption = .ReadProperty(Name:="Caption", DefaultValue:="bluButton")
        Let My_Orientation = .ReadProperty(Name:="Orientation", DefaultValue:=bluORIENTATION.Horizontal)
        Let My_State = .ReadProperty(Name:="State", DefaultValue:=bluSTATE.Inactive)
        Let My_Style = .ReadProperty(Name:="Style", DefaultValue:=bluSTYLE.Normal)
    End With
    Call SetForeBackColours
    
    'Only subclass when the code is actually running (not in design time / compiling)
    If blu.UserMode = True Then
        'Attach the mouse tracking
        Set MouseEvents = New bluMouseEvents
        Let MouseEvents.MousePointer = IDC.IDC_HAND
        Call MouseEvents.Attach(UserControl.hWnd)
    End If
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize(): Call Me.Refresh: End Sub

'CONTROL Terminate _
 ======================================================================================
Private Sub UserControl_Terminate()
    'Detatch the mouse tracking subclassing
    Set MouseEvents = Nothing
End Sub

'CONTROL WriteProperties _
 ======================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(Name:="ActiveColour", Value:=My_ActiveColour, DefaultValue:=blu.ActiveColour)
        Call .WriteProperty(Name:="Alignment", Value:=My_Alignment, DefaultValue:=VBRUN.AlignmentConstants.vbCenter)
        Call .WriteProperty(Name:="BaseColour", Value:=My_BaseColour, DefaultValue:=blu.BaseColour)
        Call .WriteProperty(Name:="Caption", Value:=My_Caption, DefaultValue:="Blu Button")
        Call .WriteProperty(Name:="Orientation", Value:=My_Orientation, DefaultValue:=bluORIENTATION.Horizontal)
        Call .WriteProperty(Name:="State", Value:=My_State, DefaultValue:=bluSTATE.Inactive)
        Call .WriteProperty(Name:="Style", Value:=My_Style, DefaultValue:=bluSTYLE.Normal)
    End With
End Sub

'EVENT MouseEvents MOUSEIN : The mouse has entered the control _
 ======================================================================================
Private Sub MouseEvents_MouseIn()
    'The mouse is in the button, we'll cause a hover effect as long as the button is _
     not locked into an active state
    Let IsHovered = True
    Call SetForeBackColours
    Call Me.Refresh
    RaiseEvent MouseIn
End Sub

'EVENT MouseEvents MOUSEOUT : The mouse has gone out of the control _
 ======================================================================================
Private Sub MouseEvents_MouseOut()
    'The mouse has left the button, we need to undo the hover effect, as long as the _
     button is not locked into an active state
    Let IsHovered = False
    Call SetForeBackColours
    Call Me.Refresh
    RaiseEvent MouseOut
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY ActiveColour _
 ======================================================================================
Public Property Get ActiveColour() As OLE_COLOR: Let ActiveColour = My_ActiveColour: End Property
Public Property Let ActiveColour(ByVal NewColour As OLE_COLOR)
    Let My_ActiveColour = NewColour
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("ActiveColour")
End Property

'PROPERTY Alignment : Text alignment (left / center / right) _
 ======================================================================================
Public Property Get Alignment() As VBRUN.AlignmentConstants: Let Alignment = My_Alignment: End Property
Public Property Let Alignment(ByVal NewAlignment As VBRUN.AlignmentConstants)
    Let My_Alignment = NewAlignment
    Call Me.Refresh
    Call UserControl.PropertyChanged("Alignment")
End Property

'PROPERTY BaseColour _
 ======================================================================================
Public Property Get BaseColour() As OLE_COLOR: Let BaseColour = My_BaseColour: End Property
Public Property Let BaseColour(ByVal NewColour As OLE_COLOR)
    Let My_BaseColour = NewColour
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("BaseColour")
End Property

'PROPERTY Caption _
 ======================================================================================
Public Property Get Caption() As String: Let Caption = My_Caption: End Property
Public Property Let Caption(ByVal NewCaption As String)
    Let My_Caption = NewCaption
    Call Me.Refresh
    Call UserControl.PropertyChanged("Caption")
End Property

'PROPERTY Orientation _
 ======================================================================================
Public Property Get Orientation() As bluORIENTATION: Let Orientation = My_Orientation: End Property
Public Property Let Orientation(ByVal NewOrientation As bluORIENTATION)
    'If switching between horizontal / vertical (or vice-versa) then rotate the control
    If ( _
        Me.Orientation = bluORIENTATION.Horizontal And _
        (NewOrientation = bluORIENTATION.VerticalDown Or NewOrientation = bluORIENTATION.VerticalUp) And _
        UserControl.Width > UserControl.Height _
    ) Or ( _
        NewOrientation = bluORIENTATION.Horizontal And _
        (Me.Orientation = bluORIENTATION.VerticalDown Or Me.Orientation = bluORIENTATION.VerticalUp) And _
        UserControl.Height > UserControl.Width _
    ) Then
        Dim Swap As Long
        Let Swap = UserControl.Width
        Let UserControl.Width = UserControl.Height
        Let UserControl.Height = Swap
    End If
    
    Let My_Orientation = NewOrientation
    Call Me.Refresh
    Call UserControl.PropertyChanged("Orientation")
End Property

'PROPERTY State : Can be used to lock the button into a hovered state _
 ======================================================================================
Public Property Get State() As bluSTATE: Let State = My_State: End Property
Public Property Let State(ByVal NewState As bluSTATE)
    Let My_State = NewState
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("State")
End Property

'PROPERTY Style _
 ======================================================================================
Public Property Get Style() As bluSTYLE: Let Style = My_Style: End Property
Public Property Let Style(ByVal NewStyle As bluSTYLE)
    Let My_Style = NewStyle
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("Style")
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Refresh : Force a repaint _
 ======================================================================================
Public Sub Refresh()
    Call UserControl_Paint
    Call UserControl.Refresh
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'SetForeBackColours : Based on the state of the button set the fore/back colours _
 ======================================================================================
Private Sub SetForeBackColours()
    'The fore/back colours will be swapped depending on a number of factors
    Dim SwapColours As Boolean
    'If the style is Normal, the button will be default colours. _
     If the style is Invert, the colours will be swapped
    Let SwapColours = (My_Style <> Normal)
    'If the button state is Active, the colours will swap. _
     This cannot be overrided by hover
    If My_State = Active Then Let SwapColours = Not SwapColours
    'If the button is hovered over, the colours will swap (except when Active)
    If My_State = Inactive And IsHovered = True Then Let SwapColours = Not SwapColours
    
    'Set background colour
    Let UserControl.BackColor = blu.OLETranslateColor( _
        IIf(SwapColours = False, My_BaseColour, My_ActiveColour) _
    )
    'Set the foreground (text) colour
    Let UserControl.ForeColor = blu.OLETranslateColor( _
        IIf(SwapColours = False, My_ActiveColour, My_BaseColour) _
    )
End Sub
