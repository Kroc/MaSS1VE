VERSION 5.00
Begin VB.UserControl bluLabel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   ForeColor       =   &H00808080&
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   122
End
Attribute VB_Name = "bluLabel"
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
'CONTROL :: bluLabel

'Status             In flux
'Dependencies       blu.bas, WIN32.bas
'Last Updated       10-AUG-13

'NOTE: This is currently in the process of being rewritten to be more API-drive _
 (i.e. handling the `WM_PAINT` ourselves) to reduce flicker. The behaviour is not _
 right yet and there will be visual problems with this control

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

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'To avoid repainting hundreds of times as each property gets set at initialisation _
 we use a flag here to tell the control to hold off on painting
Dim Freeze As Boolean

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Define our events we'll expose
Event Click()

'CONTROL InitProperties _
 ======================================================================================
Private Sub UserControl_InitProperties()
    'Avoid repainting until we're done with resizing
    Dim Frozen As Boolean
    If Freeze = True Then Let Frozen = True Else Let Freeze = True
    
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
    
    If Frozen = False Then Let Freeze = False
    Call UserControl_Paint
End Sub

'CONTROL Click _
 ======================================================================================
Private Sub UserControl_Click(): RaiseEvent Click: End Sub

'CONTROL Paint _
 ======================================================================================
Private Sub UserControl_Paint()
    Call Paint
End Sub

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(ByRef PropBag As PropertyBag)
    'Avoid repainting until we're done with resizing
    Dim Frozen As Boolean
    If Freeze = True Then Let Frozen = True Else Let Freeze = True
    
    With PropBag
        Let Me.ActiveColour = .ReadProperty(Name:="ActiveColour", DefaultValue:=blu.ActiveColour)
        Let Me.Alignment = .ReadProperty(Name:="Alignment", DefaultValue:=VBRUN.AlignmentConstants.vbLeftJustify)
        Let Me.BaseColour = .ReadProperty(Name:="BaseColour", DefaultValue:=blu.BaseColour)
        Let Me.Caption = .ReadProperty(Name:="Caption", DefaultValue:="bluLabel")
        Let Me.InertColour = .ReadProperty(Name:="InertColour", DefaultValue:=blu.InertColour)
        Let Me.Orientation = .ReadProperty(Name:="Orientation", DefaultValue:=bluORIENTATION.Horizontal)
        Let Me.State = .ReadProperty(Name:="State", DefaultValue:=bluSTATE.Inactive)
        Let Me.Style = .ReadProperty(Name:="Style", DefaultValue:=bluSTYLE.Normal)
        Let Me.TextColour = .ReadProperty(Name:="TextColour", DefaultValue:=blu.TextColour)
    End With
    
    If Frozen = False Then Let Freeze = False
    Call UserControl_Paint
End Sub

'CONTROL WriteProperties _
 ======================================================================================
Private Sub UserControl_WriteProperties(ByRef PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(Name:="ActiveColour", Value:=My_ActiveColour, DefaultValue:=blu.ActiveColour)
        Call .WriteProperty(Name:="Alignment", Value:=My_Alignment, DefaultValue:=VBRUN.AlignmentConstants.vbLeftJustify)
        Call .WriteProperty(Name:="BaseColour", Value:=My_BaseColour, DefaultValue:=blu.BaseColour)
        Call .WriteProperty(Name:="Caption", Value:=My_Caption, DefaultValue:="bluLabel")
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
Public Property Get ActiveColour() As OLE_COLOR
Attribute ActiveColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Let ActiveColour = My_ActiveColour
End Property

Public Property Let ActiveColour(ByVal NewColour As OLE_COLOR)
    If My_ActiveColour = NewColour Then Exit Property
    Let My_ActiveColour = NewColour
    Call UserControl_Paint
    Call UserControl.PropertyChanged("ActiveColour")
End Property

'PROPERTY Alignment _
 ======================================================================================
Public Property Get Alignment() As VBRUN.AlignmentConstants
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Text"
    Let Alignment = My_Alignment
End Property

Public Property Let Alignment(ByVal NewAlignment As VBRUN.AlignmentConstants)
    If My_Alignment = NewAlignment Then Exit Property
    Let My_Alignment = NewAlignment
    Call UserControl_Paint
    Call UserControl.PropertyChanged("Alignment")
End Property

'PROPERTY BaseColour _
 ======================================================================================
Public Property Get BaseColour() As OLE_COLOR: Let BaseColour = My_BaseColour: End Property
Attribute BaseColour.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let BaseColour(ByVal NewColour As OLE_COLOR)
    If My_BaseColour = NewColour Then Exit Property
    Let My_BaseColour = NewColour
    Call UserControl_Paint
    Call UserControl.PropertyChanged("BaseColour")
End Property

'PROPERTY Caption _
 ======================================================================================
Public Property Get Caption() As String: Let Caption = My_Caption: End Property
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_MemberFlags = "200"
Public Property Let Caption(ByVal NewCaption As String)
    If My_Caption = NewCaption Then Exit Property
    Let My_Caption = NewCaption
    Call UserControl_Paint
    Call UserControl.PropertyChanged("Caption")
End Property

'PROPERTY hWnd _
 ======================================================================================
Public Property Get hWnd() As Long: Let hWnd = UserControl.hWnd: End Property

'PROPERTY InertColour _
 ======================================================================================
Public Property Get InertColour() As OLE_COLOR: Let InertColour = My_InertColour: End Property
Attribute InertColour.VB_ProcData.VB_Invoke_Property = ";Text"
Public Property Let InertColour(ByVal NewColour As OLE_COLOR)
    If My_InertColour = NewColour Then Exit Property
    Let My_InertColour = NewColour
    Call UserControl_Paint
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
    'If switching between horizontal / vertical (or vice-versa) then rotate the control
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
    Call UserControl_Paint
    Call UserControl.PropertyChanged("Orientation")
End Property

'PROPERTY State _
 ======================================================================================
Public Property Get State() As bluSTATE: Let State = My_State: End Property
Attribute State.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let State(ByVal NewState As bluSTATE)
    If My_State = NewState Then Exit Property
    Let My_State = NewState
    Call UserControl_Paint
    Call UserControl.PropertyChanged("State")
End Property

'PROPERTY TextColour _
 ======================================================================================
Public Property Get Style() As bluSTYLE: Let Style = My_Style: End Property
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Appearance"
Public Property Let Style(ByVal NewStyle As bluSTYLE)
    If My_Style = NewStyle Then Exit Property
    Let My_Style = NewStyle
    Call UserControl_Paint
    Call UserControl.PropertyChanged("Style")
End Property

'PROPERTY TextColour _
 ======================================================================================
Public Property Get TextColour() As OLE_COLOR: Let TextColour = My_TextColour: End Property
Attribute TextColour.VB_ProcData.VB_Invoke_Property = ";Text"
Public Property Let TextColour(ByVal NewColour As OLE_COLOR)
    If My_TextColour = NewColour Then Exit Property
    Let My_TextColour = NewColour
    Call UserControl_Paint
    Call UserControl.PropertyChanged("TextColour")
End Property

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Paint : Does the actual painting so that it can be shared between Run / Design Time _
 ======================================================================================
Private Sub Paint()
    'Are we being told to hold off on painting?
    If Freeze = True Then Exit Sub
    
    Dim ClientRECT As RECT
    Call WIN32.user32_GetClientRect(UserControl.hWnd, ClientRECT)
    
    'Set the colours: _
     ----------------------------------------------------------------------------------
    Select Case My_Style
        Case bluSTYLE.Invert
            'Set the background colour. I don't know if it is actually any slower to _
             set the background colour every paint, but it seems stupid to do so
'            If UserControl.BackColor <> My_ActiveColour Then _
'                Let UserControl.BackColor = My_ActiveColour
            Call WIN32.gdi32_SetDCBrushColor(UserControl.hDC, My_ActiveColour)
            'Set the text colour
            Select Case My_State
                Case bluSTATE.Active
                    Call WIN32.gdi32_SetTextColor( _
                        hndDeviceContext:=UserControl.hDC, _
                        Color:=WIN32.OLETranslateColor(My_BaseColour) _
                    )
                Case Else
                    Call WIN32.gdi32_SetTextColor( _
                        hndDeviceContext:=UserControl.hDC, _
                        Color:=WIN32.OLETranslateColor(My_InertColour) _
                    )
            End Select

        Case Else
            'Set the background colour. I don't know if it is actually any slower to _
             set the background colour every paint, but it seems stupid to do so
'            If UserControl.BackColor <> My_BaseColour Then _
'                Let UserControl.BackColor = My_BaseColour
            Call WIN32.gdi32_SetDCBrushColor(UserControl.hDC, My_BaseColour)
            'Set the text colour
            Select Case My_State
                Case bluSTATE.Active
                    Call WIN32.gdi32_SetTextColor( _
                        hndDeviceContext:=UserControl.hDC, _
                        Color:=WIN32.OLETranslateColor(My_ActiveColour) _
                    )
                Case Else
                    Call WIN32.gdi32_SetTextColor( _
                        hndDeviceContext:=UserControl.hDC, _
                        Color:=WIN32.OLETranslateColor(My_TextColour) _
                    )
            End Select
    End Select
    
    Call WIN32.user32_FillRect( _
        UserControl.hDC, ClientRECT, _
        WIN32.gdi32_GetStockObject(DC_BRUSH) _
    )

    'Determine the rotation
    Dim Escapement As Long
    Select Case My_Orientation
        Case bluORIENTATION.Horizontal: Let Escapement = 0
        Case bluORIENTATION.VerticalDown: Let Escapement = -900
        Case bluORIENTATION.VerticalUp: Let Escapement = 900
    End Select

    'Create the font
    Dim hndFont As Long
    Let hndFont = WIN32.gdi32_CreateFont( _
        Height:=15, Width:=0, _
        Escapement:=Escapement, Orientation:=Escapement, _
        Weight:=FW_NORMAL, Italic:=API_FALSE, Underline:=API_FALSE, _
        StrikeOut:=API_FALSE, CharSet:=DEFAULT_CHARSET, _
        OutputPrecision:=OUT_DEFAULT_PRECIS, ClipPrecision:=CLIP_DEFAULT_PRECIS, _
        Quality:=DEFAULT_QUALITY, PitchAndFamily:=VARIABLE_PITCH Or FF_DONTCARE, _
        Face:="Arial" _
    )
    
    'Select the font (remembering the previous object selected to clean up later)
    Dim hndOld As Long
    Let hndOld = WIN32.gdi32_SelectObject(UserControl.hDC, hndFont)
    
    'Draw the text! _
     ----------------------------------------------------------------------------------
    Select Case My_Orientation
        Case bluORIENTATION.Horizontal
            Call WIN32.gdi32_SetTextAlign(UserControl.hWnd, TA_TOP Or TA_LEFT Or TA_NOUPDATECP)
            
            Dim Alignment As Long
            Select Case My_Alignment
                Case VBRUN.AlignmentConstants.vbCenter
                    Let Alignment = DT.DT_CENTER
                Case VBRUN.AlignmentConstants.vbLeftJustify
                    Let Alignment = DT.DT_LEFT
                Case VBRUN.AlignmentConstants.vbRightJustify
                    Let Alignment = DT.DT_RIGHT
            End Select
            
            With ClientRECT
                Let .Left = .Left + 8
                Let .Right = .Right - 8
            End With
            
            Call WIN32.user32_DrawText( _
                hndDeviceContext:=UserControl.hDC, _
                Text:=My_Caption, Length:=Len(My_Caption), _
                BoundingBox:=ClientRECT, _
                Format:=Alignment Or DT_VCENTER Or DT_NOPREFIX Or DT_SINGLELINE _
            )
            
        Case Else
            Select Case My_Alignment
                Case VBRUN.AlignmentConstants.vbCenter
                    Call WIN32.gdi32_SetTextAlign(UserControl.hDC, TA_TOPCENTER)
                Case VBRUN.AlignmentConstants.vbLeftJustify
                    Call WIN32.gdi32_SetTextAlign(UserControl.hDC, TA_LEFT)
                Case VBRUN.AlignmentConstants.vbRightJustify
                    Call WIN32.gdi32_SetTextAlign(UserControl.hDC, TA.TA_RIGHT)
            End Select
            
            Dim TextPos As Long
            Select Case My_Alignment
                Case VBRUN.AlignmentConstants.vbCenter
                    Let TextPos = IIf(My_Orientation = Horizontal, UserControl.ScaleWidth \ 2, UserControl.ScaleHeight \ 2)
                Case VBRUN.AlignmentConstants.vbLeftJustify
                    Let TextPos = IIf(My_Orientation = VerticalUp, UserControl.ScaleHeight - 8, 8)
                Case VBRUN.AlignmentConstants.vbRightJustify
                    If My_Orientation = Horizontal Then
                        Let TextPos = UserControl.ScaleWidth - 8
                    ElseIf My_Orientation = VerticalUp Then
                        Let TextPos = 8
                    ElseIf My_Orientation = VerticalDown Then
                        Let TextPos = UserControl.ScaleHeight - 8
                    End If
            End Select
            
            Dim X As Long, Y As Long
            Select Case My_Orientation
                Case bluORIENTATION.Horizontal
                    Let X = TextPos: Let Y = (UserControl.ScaleHeight - 15) \ 2
                Case bluORIENTATION.VerticalUp
                    Let X = (UserControl.ScaleWidth - 15) \ 2: Let Y = TextPos
                Case bluORIENTATION.VerticalDown
                    Let X = (UserControl.ScaleWidth + 15) \ 2: Let Y = TextPos
            End Select
            
            Call WIN32.gdi32_TextOut( _
                hndDeviceContext:=UserControl.hDC, _
                X:=X, Y:=Y, Text:=My_Caption, Length:=Len(My_Caption) _
            )
            
    End Select
    
    'Select the previous object into the DC (i.e. unselect the font)
    Call WIN32.gdi32_SelectObject(UserControl.hDC, hndOld)
    Call WIN32.gdi32_DeleteObject(hndFont)
End Sub

'/// SUBCLASS /////////////////////////////////////////////////////////////////////////

