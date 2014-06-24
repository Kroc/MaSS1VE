VERSION 5.00
Begin VB.UserControl bluLabel 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   372
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1836
   ForeColor       =   &H00808080&
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   153
   ToolboxBitmap   =   "bluLabel.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "bluLabel"
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
'CONTROL :: bluLabel

'Status             Ready to use
'Dependencies       blu.bas
'Last Updated       31-AUG-13
'Last Update        Ignore errors in `Paint` procedure as `hDC` might not be available

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'Colours
Private My_BaseColour As OLE_COLOR
Private My_TextColour As OLE_COLOR
Private My_ActiveColour As OLE_COLOR
Private My_InertColour As OLE_COLOR

'Appearance
Private My_Orientation As bluORIENTATION
Private My_State As bluSTATE
Private My_Style As bluSTYLE

'Text
Private My_Caption As String
Private My_Alignment As VBRUN.AlignmentConstants

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Define our events we'll expose
Event Click()

'CONTROL InitProperties _
 ======================================================================================
Private Sub UserControl_InitProperties()
    'Colours
    Let Me.BaseColour = blu.BaseColour
    Let Me.TextColour = blu.TextColour
    Let Me.ActiveColour = blu.ActiveColour
    Let Me.InertColour = blu.InertColour
    'Appearance
    Let Me.Orientation = bluORIENTATION.Horizontal
    'Text
    Let Me.Alignment = vbLeftJustify
    Let Me.Caption = "bluLabel"
End Sub

'CONTROL Click _
 ======================================================================================
Private Sub UserControl_Click(): RaiseEvent Click: End Sub

'CONTROL Paint _
 ======================================================================================
Private Sub UserControl_Paint()
    'In some instances, the hDC may not be available to us
    On Error Resume Next
    
    'Select the background colour
    Call blu.gdi32_SetDCBrushColor( _
        UserControl.hDC, blu.OLETranslateColor(UserControl.BackColor) _
    )
    'Get the dimensions of the label _
     (can't use `GetClientRect` as we don't have a hWnd!)
    Dim ClientRECT As RECT
    Call blu.user32_SetRect( _
        ClientRECT, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight _
    )
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
Private Sub UserControl_ReadProperties(ByRef PropBag As PropertyBag)
    With PropBag
        Let My_ActiveColour = .ReadProperty(Name:="ActiveColour", DefaultValue:=blu.ActiveColour)
        Let My_Alignment = .ReadProperty(Name:="Alignment", DefaultValue:=VBRUN.AlignmentConstants.vbLeftJustify)
        Let My_BaseColour = .ReadProperty(Name:="BaseColour", DefaultValue:=blu.BaseColour)
        Let My_Caption = .ReadProperty(Name:="Caption", DefaultValue:="bluLabel")
        Let Me.Enabled = .ReadProperty(Name:="Enabled", DefaultValue:=True)
        Let My_InertColour = .ReadProperty(Name:="InertColour", DefaultValue:=blu.InertColour)
        Let My_Orientation = .ReadProperty(Name:="Orientation", DefaultValue:=bluORIENTATION.Horizontal)
        Let My_State = .ReadProperty(Name:="State", DefaultValue:=bluSTATE.Inactive)
        Let My_Style = .ReadProperty(Name:="Style", DefaultValue:=bluSTYLE.Normal)
        Let My_TextColour = .ReadProperty(Name:="TextColour", DefaultValue:=blu.TextColour)
    End With
    
    Call SetForeBackColours
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize()
    Call Me.Refresh
End Sub

'CONTROL WriteProperties _
 ======================================================================================
Private Sub UserControl_WriteProperties(ByRef PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(Name:="ActiveColour", Value:=My_ActiveColour, DefaultValue:=blu.ActiveColour)
        Call .WriteProperty(Name:="Alignment", Value:=My_Alignment, DefaultValue:=VBRUN.AlignmentConstants.vbLeftJustify)
        Call .WriteProperty(Name:="BaseColour", Value:=My_BaseColour, DefaultValue:=blu.BaseColour)
        Call .WriteProperty(Name:="Caption", Value:=My_Caption, DefaultValue:="bluLabel")
        Call .WriteProperty(Name:="Enabled", Value:=UserControl.Enabled, DefaultValue:=True)
        Call .WriteProperty(Name:="InertColour", Value:=My_InertColour, DefaultValue:=blu.InertColour)
        Call .WriteProperty(Name:="Orientation", Value:=My_Orientation, DefaultValue:=bluORIENTATION.Horizontal)
        Call .WriteProperty(Name:="State", Value:=My_State, DefaultValue:=bluSTATE.Inactive)
        Call .WriteProperty(Name:="Style", Value:=My_Style, DefaultValue:=bluSTYLE.Normal)
        Call .WriteProperty(Name:="TextColour", Value:=My_TextColour, DefaultValue:=blu.TextColour)
    End With
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY ActiveColour: _
 ======================================================================================
Public Property Get ActiveColour() As OLE_COLOR: Let ActiveColour = My_ActiveColour: End Property
Attribute ActiveColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let ActiveColour(ByVal NewColour As OLE_COLOR)
    If My_ActiveColour = NewColour Then Exit Property
    Let My_ActiveColour = NewColour
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("ActiveColour")
End Property

'PROPERTY Alignment _
 ======================================================================================
Public Property Get Alignment() As VBRUN.AlignmentConstants: Let Alignment = My_Alignment: End Property
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Text"
Public Property Let Alignment(ByVal NewAlignment As VBRUN.AlignmentConstants)
    If My_Alignment = NewAlignment Then Exit Property
    Let My_Alignment = NewAlignment
    Call Me.Refresh
    Call UserControl.PropertyChanged("Alignment")
End Property

'PROPERTY BaseColour _
 ======================================================================================
Public Property Get BaseColour() As OLE_COLOR: Let BaseColour = My_BaseColour: End Property
Attribute BaseColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BaseColour.VB_UserMemId = -501
Public Property Let BaseColour(ByVal NewColour As OLE_COLOR)
    If My_BaseColour = NewColour Then Exit Property
    Let My_BaseColour = NewColour
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("BaseColour")
End Property

'PROPERTY Caption _
 ======================================================================================
Public Property Get Caption() As String: Let Caption = My_Caption: End Property
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
Public Property Let Caption(ByVal NewCaption As String)
    If My_Caption = NewCaption Then Exit Property
    Let My_Caption = NewCaption
    Call Me.Refresh
    Call UserControl.PropertyChanged("Caption")
End Property

'PROPERTY Enabled : You can disable the control to allow click-through _
 ======================================================================================
Public Property Get Enabled() As Boolean: Let Enabled = UserControl.Enabled: End Property
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
Public Property Let Enabled(ByVal newValue As Boolean)
   Let UserControl.Enabled = newValue
   Call PropertyChanged("Enabled")
End Property

'PROPERTY InertColour _
 ======================================================================================
Public Property Get InertColour() As OLE_COLOR: Let InertColour = My_InertColour: End Property
Attribute InertColour.VB_ProcData.VB_Invoke_Property = ";Text"
Public Property Let InertColour(ByVal NewColour As OLE_COLOR)
    If My_InertColour = NewColour Then Exit Property
    Let My_InertColour = NewColour
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("InertColour")
End Property

'PROPERTY Orientation _
 ======================================================================================
Public Property Get Orientation() As bluORIENTATION
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Let Orientation = My_Orientation
End Property

Public Property Let Orientation(ByVal NewOrientation As bluORIENTATION)
    If My_Orientation = NewOrientation Then Exit Property
    'If switching between horizontal / vertical, rotate the control
    If ( _
        My_Orientation = bluORIENTATION.Horizontal And _
        (NewOrientation = bluORIENTATION.VerticalDown Or NewOrientation = bluORIENTATION.VerticalUp) And _
        UserControl.Width > UserControl.Height _
    ) Or ( _
        NewOrientation = bluORIENTATION.Horizontal And _
        (My_Orientation = bluORIENTATION.VerticalDown Or My_Orientation = bluORIENTATION.VerticalUp) And _
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

'PROPERTY State _
 ======================================================================================
Public Property Get State() As bluSTATE: Let State = My_State: End Property
Attribute State.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let State(ByVal NewState As bluSTATE)
    If My_State = NewState Then Exit Property
    Let My_State = NewState
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("State")
End Property

'PROPERTY Style : Normal or invert colour scheme _
 ======================================================================================
Public Property Get Style() As bluSTYLE: Let Style = My_Style: End Property
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let Style(ByVal NewStyle As bluSTYLE)
    If My_Style = NewStyle Then Exit Property
    Let My_Style = NewStyle
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("Style")
End Property

'PROPERTY TextColour _
 ======================================================================================
Public Property Get TextColour() As OLE_COLOR: Let TextColour = My_TextColour: End Property
Attribute TextColour.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute TextColour.VB_UserMemId = -513
Public Property Let TextColour(ByVal NewColour As OLE_COLOR)
    If My_TextColour = NewColour Then Exit Property
    Let My_TextColour = NewColour
    Call SetForeBackColours
    Call Me.Refresh
    Call UserControl.PropertyChanged("TextColour")
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Refresh : Force a repaint _
 ======================================================================================
Public Sub Refresh()
    Call UserControl_Paint
    Call UserControl.Refresh
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'SetForeBackColours : Based on the state of the label set the fore/back colours _
 ======================================================================================
Private Sub SetForeBackColours()
    Select Case My_Style
        Case bluSTYLE.Invert
            'Set the background colour
            Let UserControl.BackColor = blu.OLETranslateColor(My_ActiveColour)
            'Set the text colour
            Let UserControl.ForeColor = blu.OLETranslateColor( _
                IIf(My_State = Active, My_BaseColour, My_InertColour) _
            )

        Case Else
            'Set the background colour
            Let UserControl.BackColor = blu.OLETranslateColor(My_BaseColour)
            'Set the text colour
            Let UserControl.ForeColor = blu.OLETranslateColor( _
                IIf(My_State = Active, My_ActiveColour, My_TextColour) _
            )
    End Select
End Sub
