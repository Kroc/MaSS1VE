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
End
Attribute VB_Name = "bluButton"
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
'CONTROL :: BluButton

'Status             In flux
'Dependencies       blu.bas, bluMouseEvents.cls (bluMagic.cls), WIN32.bas
'Last Updated       19-AUG-13
'Last Update        Removed nested bluLabel control and added API-driven painting, _
                    also added real rotation of text so that the same text API is _
                    used for horizontal and vertical text

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
    Let Me.Caption = "bluButton"
    Let Me.Orientation = bluORIENTATION.Horizontal
    Let Me.Style = bluSTYLE.Normal
    Let Me.State = bluSTATE.Inactive
End Sub

'CONTROL Paint _
 ======================================================================================
Private Sub UserControl_Paint()
    'First determine the colour scheme...
    Dim SwapColours As Boolean
    'If the style is Normal, the button will be default colours. _
     If the style is Invert, the colours will be swapped
    Let SwapColours = (My_Style <> Normal)
    'If the button state is Active, the colours will swap. _
     This cannot be overrided by hover
    If My_State = Active Then Let SwapColours = Not SwapColours
    'If the button is hovered over, the colours will swap (except when Active)
    If My_State = Inactive And IsHovered = True Then Let SwapColours = Not SwapColours
    
    'Clear background: _
     ----------------------------------------------------------------------------------
    'Set the colour for clearing the background
    Let UserControl.BackColor = WIN32.OLETranslateColor( _
        IIf(SwapColours = False, My_BaseColour, My_ActiveColour) _
    )
    Call WIN32.gdi32_SetDCBrushColor( _
        hndDeviceContext:=UserControl.hDC, _
        Color:=WIN32.OLETranslateColor(UserControl.BackColor) _
    )
    
    'Get the dimensions of the button
    Dim ClientRECT As RECT
    Call WIN32.user32_GetClientRect(UserControl.hWnd, ClientRECT)
    'Then use those to fill with the selected background colour
    Call WIN32.user32_FillRect( _
        UserControl.hDC, ClientRECT, _
        WIN32.gdi32_GetStockObject(DC_BRUSH) _
    )
    
    'Create and set the font: _
     ----------------------------------------------------------------------------------
    'Create the GDI font object that describes our font properties
    Dim hndFont As Long
    Let hndFont = WIN32.gdi32_CreateFont( _
        Height:=15, Width:=0, _
        Escapement:=0, Orientation:=0, _
        Weight:=FW_NORMAL, Italic:=API_FALSE, Underline:=API_FALSE, _
        StrikeOut:=API_FALSE, CharSet:=DEFAULT_CHARSET, _
        OutputPrecision:=OUT_DEFAULT_PRECIS, ClipPrecision:=CLIP_DEFAULT_PRECIS, _
        Quality:=DEFAULT_QUALITY, PitchAndFamily:=VARIABLE_PITCH Or FF_DONTCARE, _
        Face:="Arial" _
    )

    'Select the font (remembering the previous object selected to clean up later)
    Dim hndOld As Long
    Let hndOld = WIN32.gdi32_SelectObject(UserControl.hDC, hndFont)
    
    'The `DrawText` API doesn't work with the position set by `SetTextAlign`, _
     so we ensure it's set to a safe, non-interfering value
    Call WIN32.gdi32_SetTextAlign( _
        UserControl.hWnd, TA_TOP Or TA_LEFT Or TA_NOUPDATECP _
    )
    
    'Prepare the alignment value used for the `DrawText` API
    Dim Alignment As DT
    Let Alignment = Choose(My_Alignment + 1, DT.DT_LEFT, DT.DT_RIGHT, DT.DT_CENTER)
    
    'Rotate the text? _
     ----------------------------------------------------------------------------------
    'Made possible with directions from: <edais.mvps.org/Tutorials/GDI/DC/DCch8.html>
    If My_Orientation <> Horizontal Then
        'The button is already in a vertical shape, but we want to rotate a horizontal _
         piece of text, so we have to swap the dimensions of the button to begin with
        Call WIN32.user32_SetRect( _
            ClientRECT, _
            ClientRECT.Left, ClientRECT.Top, ClientRECT.Bottom, ClientRECT.Right _
        )
        'In addition to that, we also need to position our text with its center at _
         0,0, instead of the top-left corner, so that when we rotate, the text stays _
         centered and doesn't swing off out of place
        Call WIN32.user32_OffsetRect( _
            ClientRECT, -ClientRECT.Right \ 2, -ClientRECT.Bottom \ 2 _
        )
        
        'Now we need to move the origin point (0,0) to the centre of the button _
         so that the rotated text obviously appears in the center of the button _
         whilst the rotation occurs around the centrepoint of the text
        Dim Org As POINT
        Call WIN32.gdi32_SetViewportOrgEx( _
            UserControl.hDC, _
            UserControl.ScaleWidth \ 2, UserControl.ScaleHeight \ 2, _
            Org _
        )
        
        'In order to use Get/SetWorldTransform we have to make this call
        Dim OldGM As Long
        Let OldGM = WIN32.gdi32_SetGraphicsMode(UserControl.hDC, GM_ADVANCED)
        
        'Now calculate the rotation
        Const Pi As Single = 3.14159
        Dim RotAng As Single
        Let RotAng = IIf(My_Orientation = VerticalDown, -90, 90)
        Dim RotRad As Single
        Let RotRad = (RotAng / 180) * Pi
        
        'Read any current transform from the device context
        Dim OldXForm As XFORM, RotXForm As XFORM
        Call WIN32.gdi32_GetWorldTransform(UserControl.hDC, OldXForm)
        
        'Define our rotation matrix
        With RotXForm
            Let .eM11 = Cos(RotRad)
            Let .eM21 = Sin(RotRad)
            Let .eM12 = -.eM21
            Let .eM22 = .eM11
        End With
        
        'Apply the matrix -- rotate the world!
        Call WIN32.gdi32_SetWorldTransform(UserControl.hDC, RotXForm)
    End If
    
    'Draw the text! _
     ----------------------------------------------------------------------------------
    'Set the colour of the text
    Let UserControl.ForeColor = WIN32.OLETranslateColor( _
        IIf(SwapColours = False, My_ActiveColour, My_BaseColour) _
    )
    Call WIN32.gdi32_SetTextColor( _
        hndDeviceContext:=UserControl.hDC, _
        Color:=WIN32.OLETranslateColor(UserControl.ForeColor) _
    )
    
    'Add a little margin either side
    With ClientRECT
        Let .Left = .Left + 8
        Let .Right = .Right - 8
    End With

    'Now just paint the text
    Call WIN32.user32_DrawText( _
        hndDeviceContext:=UserControl.hDC, _
        Text:=My_Caption, Length:=Len(My_Caption), _
        BoundingBox:=ClientRECT, _
        Format:=Alignment _
                Or DT_VCENTER Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP _
    )
    
    'Clean up: _
     ----------------------------------------------------------------------------------
    'If we rotated the text, we need to do some additional clean up
    If My_Orientation <> Horizontal Then
        'Restore the previous world transform
        Call WIN32.gdi32_SetWorldTransform(UserControl.hDC, OldXForm)
        'Switch back to the previous graphics mode
        Call WIN32.gdi32_SetGraphicsMode(UserControl.hDC, OldGM)
        'Return the origin point (0,0) back to the upper-left corner
        Call WIN32.gdi32_SetViewportOrgEx(UserControl.hDC, Org.X, Org.Y, Org)
    End If
    
    'Select the previous object into the DC (i.e. unselect the font)
    Call WIN32.gdi32_SelectObject(UserControl.hDC, hndOld)
    Call WIN32.gdi32_DeleteObject(hndFont)
End Sub

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Let Me.ActiveColour = .ReadProperty(Name:="ActiveColour", DefaultValue:=blu.ActiveColour)
        Let Me.Alignment = .ReadProperty(Name:="Alignment", DefaultValue:=VBRUN.AlignmentConstants.vbCenter)
        Let Me.BaseColour = .ReadProperty(Name:="BaseColour", DefaultValue:=blu.BaseColour)
        Let Me.Caption = .ReadProperty(Name:="Caption", DefaultValue:="bluButton")
        Let Me.Orientation = .ReadProperty(Name:="Orientation", DefaultValue:=bluORIENTATION.Horizontal)
        Let Me.State = .ReadProperty(Name:="State", DefaultValue:=bluSTATE.Inactive)
        Let Me.Style = .ReadProperty(Name:="Style", DefaultValue:=bluSTYLE.Normal)
    End With
    
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
    Call Me.Refresh
    RaiseEvent MouseIn
End Sub

'EVENT MouseEvents MOUSEOUT : The mouse has gone out of the control _
 ======================================================================================
Private Sub MouseEvents_MouseOut()
    'The mouse has left the button, we need to undo the hover effect, as long as the _
     button is not locked into an active state
    Let IsHovered = False
    Call Me.Refresh
    RaiseEvent MouseOut
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY ActiveColour _
 ======================================================================================
Public Property Get ActiveColour() As OLE_COLOR
    Let ActiveColour = My_ActiveColour
End Property

Public Property Let ActiveColour(ByVal NewColour As OLE_COLOR)
    Let My_ActiveColour = NewColour
    Call Me.Refresh
    Call UserControl.PropertyChanged("ActiveColour")
End Property

'PROPERTY Alignment : Text alignment (left / center / right) _
 ======================================================================================
Public Property Get Alignment() As VBRUN.AlignmentConstants
    Let Alignment = My_Alignment
End Property

Public Property Let Alignment(ByVal NewAlignment As VBRUN.AlignmentConstants)
    Let My_Alignment = NewAlignment
    Call Me.Refresh
    Call UserControl.PropertyChanged("Alignment")
End Property

'PROPERTY BaseColour _
 ======================================================================================
Public Property Get BaseColour() As OLE_COLOR
    Let BaseColour = My_BaseColour
End Property

Public Property Let BaseColour(ByVal NewColour As OLE_COLOR)
    Let My_BaseColour = NewColour
    Call Me.Refresh
    Call UserControl.PropertyChanged("BaseColour")
End Property

'PROPERTY Caption _
 ======================================================================================
Public Property Get Caption() As String
    Let Caption = My_Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    Let My_Caption = NewCaption
    Call Me.Refresh
    Call UserControl.PropertyChanged("Caption")
End Property

'PROPERTY Orientation _
 ======================================================================================
Public Property Get Orientation() As bluORIENTATION
    Let Orientation = My_Orientation
End Property

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
    Call Me.Refresh
    Call UserControl.PropertyChanged("State")
End Property

'PROPERTY Style _
 ======================================================================================
Public Property Get Style() As bluSTYLE: Let Style = My_Style: End Property
Public Property Let Style(ByVal NewStyle As bluSTYLE)
    Let My_Style = NewStyle
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
