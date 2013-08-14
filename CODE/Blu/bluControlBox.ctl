VERSION 5.00
Begin VB.UserControl bluControlBox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
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
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
End
Attribute VB_Name = "bluControlBox"
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
'CONTROL :: bluControlBox

'Provides a Minimise / Maximise / Close Window button. bluWindow will make a form _
 borderless, so we need to provide our own window control buttons. bluWindow could _
 draw these itself, but since there's no borders other controls on the form covering _
 where the buttons would be prevent this. This is mainly a problem with MDIForms _
 which must have an aligning picturebox if you want to place anything on the MDIForm

'Status             Ready, awaiting refactoring
'Dependencies       blu.bas, bluButton.ctl (bluLabel.ctl), WIN32.bas
'Last Updated       24-JUL-13

'NOTE: This is due to be rewritten to not depend upon nested controls. The plan is _
 also for this control to act not as one control box button, but as the entire set, _
 where we will detect from the parent form which buttons (min / max / close) should _
 be available

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Private My_ActiveColour As OLE_COLOR
Private My_BaseColour As OLE_COLOR

Public Enum bluControlBox_Kind
    Quit = 0
    Minimize = 1
    Maximize = 2
End Enum
Private My_Kind As bluControlBox_Kind

Private My_State As bluSTATE
Private My_Style As bluSTYLE

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'We'll use this to provide MouseIn/Out events
Private WithEvents MouseEvents As bluMouseEvents
Attribute MouseEvents.VB_VarHelpID = -1

'If the button is a hovered state
Private IsHovered As Boolean
'If the mouse is held down, so as to do the clicked effect
Private IsMouseDown As Boolean

'The form the control belongs too, so excluding any picturebox / frame containers
Private ParentForm As Object

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'The label in the button is already subclassed and will provide MouseIn/Out events _
 which we can then expose to the button controller
Event Click()
Event MouseIn()
Event MouseOut()

'CONTROL Click _
 ======================================================================================
Private Sub UserControl_Click()
    'What kind of action do we need to take?
    Select Case My_Kind
        Case bluControlBox_Kind.Quit
            Unload ParentForm
        Case bluControlBox_Kind.Minimize
            Let ParentForm.WindowState = VBRUN.FormWindowStateConstants.vbMinimized
        Case bluControlBox_Kind.Maximize
            Let ParentForm.WindowState = IIf( _
                Expression:=ParentForm.WindowState = VBRUN.FormWindowStateConstants.vbNormal, _
                TruePart:=VBRUN.FormWindowStateConstants.vbMaximized, _
                FalsePart:=VBRUN.FormWindowStateConstants.vbNormal _
            )
    End Select
End Sub

'CONTROL InitProperties : When a new control is placed on the form _
 ======================================================================================
Private Sub UserControl_InitProperties()
    Let Me.ActiveColour = blu.ActiveColour
    Let Me.BaseColour = blu.BaseColour
    Set ParentForm = GetUltimateParent()
End Sub

'CONTROL MouseDown _
 ======================================================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IsMouseDown = False Then
        Let IsMouseDown = True
        Call UserControl_Paint
    End If
End Sub

'CONTROL MouseUp _
 ======================================================================================
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IsMouseDown = True Then
        Let IsMouseDown = False
        Call UserControl_Paint
    End If
End Sub

'CONTROL Paint _
 ======================================================================================
Private Sub UserControl_Paint()
    'Clear the current button display
    Call UserControl.Cls

    'Set the colours: _
     ----------------------------------------------------------------------------------
    Dim BackColour As OLE_COLOR
    Dim ForeColour As OLE_COLOR
    
    Select Case My_Style
        Case bluSTYLE.Invert
            Let BackColour = My_ActiveColour
            Let ForeColour = My_BaseColour
            
        Case Else
            If IsMouseDown = True Then
                If My_Kind = bluControlBox_Kind.Quit Then
                    Let BackColour = blu.ClosePressColour
                Else
                    Let BackColour = My_ActiveColour
                End If
                Let ForeColour = My_BaseColour
                
            ElseIf IsHovered = True Then
                If My_Kind = bluControlBox_Kind.Quit Then
                    Let BackColour = blu.CloseHoverColour
                    Let ForeColour = My_BaseColour
                Else
                    Let BackColour = blu.BaseHoverColour
                    Let ForeColour = blu.TextHoverColour
                End If
                
            Else
                Let BackColour = My_BaseColour
                Let ForeColour = blu.TextColour
            End If
    End Select
    
    'Set the background colour. I don't know if it is actually any slower to _
     set the background colour every paint, but it seems stupid to do so
    If UserControl.BackColor <> BackColour Then _
        Let UserControl.BackColor = BackColour
    'Set the text colour
    Call WIN32.gdi32_SetTextColor( _
        hndDeviceContext:=UserControl.hDC, _
        Color:=WIN32.OLETranslateColor(ForeColour) _
    )
    
    'Set text alignment
    Call WIN32.gdi32_SetTextAlign(UserControl.hDC, TA_TOPCENTER)
    
    'Set the font: _
     ----------------------------------------------------------------------------------
    Dim hndFont As Long
    Let hndFont = WIN32.gdi32_CreateFont( _
        Height:=14, Width:=0, _
        Escapement:=0, Orientation:=0, _
        Weight:=FW_NORMAL, Italic:=API_FALSE, Underline:=API_FALSE, StrikeOut:=API_FALSE, _
        CharSet:=DEFAULT_CHARSET, _
        OutputPrecision:=OUT_DEFAULT_PRECIS, ClipPrecision:=CLIP_DEFAULT_PRECIS, _
        Quality:=DEFAULT_QUALITY, PitchAndFamily:=VARIABLE_PITCH Or FF_DONTCARE, _
        Face:="Marlett" _
    )
    Dim hndOld As Long
    Let hndOld = WIN32.gdi32_SelectObject(UserControl.hDC, hndFont)

    'Draw the text! _
     ----------------------------------------------------------------------------------
    Dim Letter As String
    Select Case My_Kind
        Case bluControlBox_Kind.Quit: Let Letter = "r"
        Case bluControlBox_Kind.Minimize: Let Letter = "0"
        Case bluControlBox_Kind.Maximize
            Let Letter = IIf( _
                Expression:=ParentForm.WindowState = VBRUN.FormWindowStateConstants.vbNormal, _
                TruePart:="1", FalsePart:="2" _
            )
    End Select
    
    Call WIN32.gdi32_TextOut( _
        hndDeviceContext:=UserControl.hDC, _
        X:=UserControl.ScaleWidth \ 2, _
        Y:=(UserControl.ScaleHeight - 14) \ 2, _
        Text:=Letter, Length:=Len(Letter) _
    )
    
    'Clean up
    Call WIN32.gdi32_SelectObject(hndDeviceContext:=UserControl.hDC, hndGdiObject:=hndOld)
    Call WIN32.gdi32_DeleteObject(hndGdiObject:=hndFont)
End Sub

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Find the parent form we need to control
    Set ParentForm = GetUltimateParent()
    
    With PropBag
        Let Me.ActiveColour = .ReadProperty(Name:="ActiveColour", DefaultValue:=blu.ActiveColour)
        Let Me.BaseColour = .ReadProperty(Name:="BaseColour", DefaultValue:=blu.BaseColour)
        Let Me.Style = .ReadProperty(Name:="Style", DefaultValue:=bluSTYLE.Normal)
        Let Me.Kind = .ReadProperty(Name:="Kind", DefaultValue:=bluControlBox_Kind.Quit)
    End With
    
    'Attach the mouse tracking
    If blu.UserMode = True Then
        Set MouseEvents = New bluMouseEvents
        Let MouseEvents.MousePointer = IDC_HAND
        Call MouseEvents.Attach(Me.hWnd)
    End If
End Sub

'CONTROL Resize : The developers is resizing the control on the form design _
 ======================================================================================
Private Sub UserControl_Resize()
    'Don't allow this control to be resized
    Let UserControl.Width = blu.Xpx(blu.Metric)
    Let UserControl.Height = blu.Ypx(blu.Metric)
End Sub

'CONTROL Show : The control has become visible, we should repaint _
 ======================================================================================
Private Sub UserControl_Show(): Call UserControl_Paint: End Sub

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
        Call .WriteProperty(Name:="BaseColour", Value:=My_BaseColour, DefaultValue:=blu.BaseColour)
        Call .WriteProperty(Name:="Style", Value:=My_Style, DefaultValue:=bluSTYLE.Normal)
        Call .WriteProperty(Name:="Kind", Value:=My_Kind, DefaultValue:=bluControlBox_Kind.Quit)
    End With
End Sub

'EVENT MouseEvents MOUSEIN : The mouse has entered the control _
 ======================================================================================
Private Sub MouseEvents_MouseIn()
    'The mouse is in the button, we'll cause a hover effect as long as the button is _
     not locked into an active state
    Let IsHovered = True: Let IsMouseDown = False
    Call UserControl_Paint
    RaiseEvent MouseIn
End Sub

'EVENT MouseEvents MOUSEOUT : The mouse has gone out of the control _
 ======================================================================================
Private Sub MouseEvents_MouseOut()
    'The mouse has left the button, we need to undo the hover effect, as long as the _
     button is not locked into an active state
    Let IsHovered = False: Let IsMouseDown = False
    Call UserControl_Paint
    RaiseEvent MouseOut
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY ActiveColour _
 ======================================================================================
Public Property Get ActiveColour() As OLE_COLOR: Let ActiveColour = My_ActiveColour: End Property
Public Property Let ActiveColour(ByVal NewColour As OLE_COLOR)
    Let My_ActiveColour = NewColour
    Call UserControl_Paint
    Call UserControl.PropertyChanged("ActiveColour")
End Property

'PROPERTY BaseColour _
 ======================================================================================
Public Property Get BaseColour() As OLE_COLOR: Let BaseColour = My_BaseColour: End Property
Public Property Let BaseColour(ByVal NewColour As OLE_COLOR)
    Let My_BaseColour = NewColour
    Call UserControl_Paint
    Call UserControl.PropertyChanged("BaseColour")
End Property

'PROPERTY hWnd : We need to expose this for the mouse tracking to attach _
 ======================================================================================
Public Property Get hWnd() As Long: Let hWnd = UserControl.hWnd: End Property

'PROPERTY Style _
 ======================================================================================
Public Property Get Style() As bluSTYLE: Let Style = My_Style: End Property
Public Property Let Style(ByVal NewStyle As bluSTYLE)
    Let My_Style = NewStyle
    Call UserControl_Paint
    Call UserControl.PropertyChanged("Style")
End Property

'PROPERTY Kind : Which kind of control button we are _
 ======================================================================================
Public Property Get Kind() As bluControlBox_Kind: Let Kind = My_Kind: End Property
Public Property Let Kind(ByVal NewKind As bluControlBox_Kind)
    Let My_Kind = NewKind
    Call UserControl_Paint
    Call UserControl.PropertyChanged("Kind")
End Property

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'GetUltimateParent : Recurses through the parent objects until we hit the top form _
 ======================================================================================
Private Function GetUltimateParent() As Object
    Set GetUltimateParent = UserControl.Parent
    Do
        On Error GoTo Fail
        Set GetUltimateParent = GetUltimateParent.Parent
    Loop
Fail:
End Function
