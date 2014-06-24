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
   ToolboxBitmap   =   "bluControlBox.ctx":0000
End
Attribute VB_Name = "bluControlBox"
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
'CONTROL :: bluControlBox

'Provides a Minimise / Maximise / Close Window button. bluBorderless will make a form _
 borderless, so we need to provide our own window control buttons. bluBorderless could _
 draw these itself, but since there's no borders other controls on the form covering _
 where the buttons would be prevent this. This is mainly a problem with MDIForms _
 which must have an aligning picturebox if you want to place anything on the MDIForm

'Status             Ready to use
'Dependencies       blu.bas, bluMouseEvents.cls (bluMagic.cls), bluBorderless.ctl
'Last Updated       19-SEP-13
'Last Update        `SendMessage` API was moved to WIN32

'/// API DEFS /////////////////////////////////////////////////////////////////////////

'All mouse events can be trapped by one window. VB apparently does this behind the _
 scenes, so we need to release the capture in order to resize the form from the control _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms646261%28v=vs.85%29.aspx>
Private Declare Function user32_ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Long

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Private My_ActiveColour As OLE_COLOR
Private My_BaseColour As OLE_COLOR

Public Enum bluControlBox_Kind
    Quit = 0
    Minimize = 1
    Maximize = 2
    Sizer = 3
End Enum
Private My_Kind As bluControlBox_Kind

Private My_Style As bluSTYLE

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'We'll use this to provide MouseIn/Out events
Private WithEvents MouseEvents As bluMouseEvents
Attribute MouseEvents.VB_VarHelpID = -1

'The form the control belongs too, so excluding any picturebox / frame containers
Private ParentForm As Object

'If we listen into the events of the parent form then we can automatically hide a _
 sizer control when the form is maximised. Unfortunately VB6 does not expose a common _
 class interface shared by regular Forms and MDIForms so we have to do both
Private WithEvents ParentFormEvents As Form
Attribute ParentFormEvents.VB_VarHelpID = -1
Private WithEvents ParentMDIFormEvents As MDIForm
Attribute ParentMDIFormEvents.VB_VarHelpID = -1

'If the parent form has a bluBorderless control, we can listen into its events so that _
 we can automatically hide bluControlBox controls if the window borders are present, _
 i.e. the Windows min / max / close buttons are visible
Private WithEvents bluBorderlessEvents As bluBorderless
Attribute bluBorderlessEvents.VB_VarHelpID = -1

'If the button is a hovered state
Private IsHovered As Boolean
'If the mouse is held down, so as to do the clicked effect
Private IsMouseDown As Boolean

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CONTROL InitProperties : When a new control is placed on the form _
 ======================================================================================
Private Sub UserControl_InitProperties()
    Let Me.ActiveColour = blu.ActiveColour
    Let Me.BaseColour = blu.BaseColour
    Call ReferenceParentForm
End Sub

'CONTROL MouseDown _
 ======================================================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Remember when the mouse is down so that any calls to repaint keep the clicked _
     effect in place
    If IsMouseDown = False Then
        Let IsMouseDown = True
        Call Refresh
    End If
    
    'If the kind of the control is a sizer, allow resizing of the form by it
    If Button = VBRUN.MouseButtonConstants.vbLeftButton And My_Kind = Sizer Then
        Const WM_NCLBUTTONDOWN As Long = &HA1
        Const HTBOTTOMRIGHT As Long = 17
        
        'With thanks to the following page for alerting me to the need to use _
         `ReleaseCapture` to get this to work! _
         <www.vbforums.com/showthread.php?250431-VB-Flexible-Shangle-%28window-resizing-grip%29>
        Call user32_ReleaseCapture
        'Simulate clicking on the lower-right window border
        Call blu.user32_SendMessage( _
            blu.GetParentForm_hWnd(UserControl.Parent, True), _
            WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0 _
        )
    End If
End Sub

'CONTROL MouseUp _
 ======================================================================================
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'When letting go of the mouse, refresh, removing the click effect
    Let IsMouseDown = False
    Call Refresh
    
    'If you hold the mouse button down inside the control but release the button _
     outside then it doesn't count (allows you to escape from an accidental close)
    Dim ClientRECT As RECT
    Call blu.user32_GetClientRect(UserControl.hWnd, ClientRECT)
    If blu.user32_PtInRect(ClientRECT, X, Y) = API_FALSE Then Exit Sub
    
    'Only left button applies to action
    If Button <> VBRUN.MouseButtonConstants.vbLeftButton Then Exit Sub
    
    'What kind of action do we need to take? _
     (the sizer is handled in `Mouse_Down`)
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

'CONTROL Paint _
 ======================================================================================
Private Sub UserControl_Paint()
    'Set the colours: _
     ----------------------------------------------------------------------------------
    'TODO: TextColour / InertColour are hard coded
    Dim BackColour As OLE_COLOR
    Dim ForeColour As OLE_COLOR
    
    Select Case My_Kind
        'A size box in the corner to resize the window
        Case bluControlBox_Kind.Sizer
            If My_Style = Invert Then
                Let BackColour = My_ActiveColour
                Let ForeColour = blu.InertColour
            Else
                Let BackColour = My_BaseColour
                Let ForeColour = blu.TextColour
            End If
            
        Case Else
            'Close / Min / Max buttons
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
    End Select
    
    'Clear the background: _
     ----------------------------------------------------------------------------------
    'Set the background colour. I don't know if it is actually any slower to _
     set the background colour every paint, but it seems stupid to do so
    If UserControl.BackColor <> BackColour Then _
        Let UserControl.BackColor = BackColour
    'Select the background colour
    Call blu.gdi32_SetDCBrushColor( _
        UserControl.hDC, UserControl.BackColor _
    )
    'Get the dimensions of the control
    Dim ClientRECT As RECT
    Call blu.user32_GetClientRect(UserControl.hWnd, ClientRECT)
    'Then use those to fill with the selected background colour
    Call blu.user32_FillRect( _
        UserControl.hDC, ClientRECT, _
        blu.gdi32_GetStockObject(DC_BRUSH) _
    )

    'Draw the text! _
     ----------------------------------------------------------------------------------
    Dim Letter As String
    Select Case My_Kind
        Case bluControlBox_Kind.Sizer: Let Letter = "p"
        Case bluControlBox_Kind.Quit: Let Letter = "r"
        Case bluControlBox_Kind.Minimize: Let Letter = "0"
        Case bluControlBox_Kind.Maximize
            Let Letter = IIf( _
                Expression:=ParentForm.WindowState = VBRUN.FormWindowStateConstants.vbNormal, _
                TruePart:="1", FalsePart:="2" _
            )
    End Select
    
    'Use the shared text drawing procedure to save effort
    Call blu.DrawText( _
        hndDeviceContext:=UserControl.hDC, BoundingBox:=ClientRECT, _
        Text:=Letter, Colour:=ForeColour, Alignment:=vbCenter, Orientation:=Horizontal, _
        FontName:="Marlett", FontSizePx:=14 _
    )
End Sub

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Call ReferenceParentForm
    
    With PropBag
        Let Me.ActiveColour = .ReadProperty(Name:="ActiveColour", DefaultValue:=blu.ActiveColour)
        Let Me.BaseColour = .ReadProperty(Name:="BaseColour", DefaultValue:=blu.BaseColour)
        Let Me.Style = .ReadProperty(Name:="Style", DefaultValue:=bluSTYLE.Normal)
        Let Me.Kind = .ReadProperty(Name:="Kind", DefaultValue:=bluControlBox_Kind.Quit)
    End With
    
    'Attach the mouse tracking
    If blu.UserMode = True Then
        Set MouseEvents = New bluMouseEvents
        'If a sizing control,
        If My_Kind = Sizer Then
            'Show the diagonal arrow,
            Let MouseEvents.MousePointer = IDC_SIZENWSE
        Else
            'Otherwise all other controls show the hand pointer
            Let MouseEvents.MousePointer = IDC_HAND
        End If
        Call MouseEvents.Attach(Me.hWnd)
    End If
End Sub

'CONTROL Resize : The developers is resizing the control on the form design _
 ======================================================================================
Private Sub UserControl_Resize()
    'If the parent form is maximised, hide ourselves if a sizer control. _
     We do this here so as to catch switching from one child to another in an MDI form
    Call ParentFormEvents_Resize
    
    'Don't allow this control to be resized _
     (The size box is smaller though)
    Let UserControl.Width = blu.Xpx(IIf(My_Kind = Sizer, 24, blu.Metric))
    Let UserControl.Height = blu.Ypx(IIf(My_Kind = Sizer, 24, blu.Metric))
End Sub

'CONTROL Terminate _
 ======================================================================================
Private Sub UserControl_Terminate()
    'Detatch the mouse tracking subclassing
    Set MouseEvents = Nothing
    'Derefernce the parent form, event listeners
    Set ParentForm = Nothing
    Set ParentFormEvents = Nothing
    Set ParentMDIFormEvents = Nothing
    Set bluBorderlessEvents = Nothing
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
    'The mouse is in the button, we'll cause a hover effect
    Let IsHovered = True: Let IsMouseDown = False
    Call Refresh
End Sub

'EVENT MouseEvents MOUSEOUT : The mouse has gone out of the control _
 ======================================================================================
Private Sub MouseEvents_MouseOut()
    'The mouse has left the button, we need to undo the hover effect
    Let IsHovered = False: Let IsMouseDown = False
    Call Refresh
End Sub

'EVENT Parent[MDI]Form_Events RESIZE _
 ======================================================================================
Private Sub ParentMDIFormEvents_Resize(): Call ParentFormEvents_Resize: End Sub
Private Sub ParentFormEvents_Resize()
    'If the window maximised, hide this control if it's a sizer
    If My_Kind = Sizer Then
        Let UserControl.Extender.Visible = Not ( _
            ParentForm.WindowState = VBRUN.FormWindowStateConstants.vbMaximized _
        )
    End If
End Sub

'EVENT Parent[MDI]Form_Events ACTIVATE _
 ======================================================================================
'When the parent form becomes visible, check for a bluBorderless control and hide _
 ourselves if bluBorderless is inactive (the Windows min / max /close buttons are _
 visible)
Private Sub ParentMDIFormEvents_Activate(): Call ParentFormEvents_Activate: End Sub
Private Sub ParentFormEvents_Activate()
    'If the form is borderless to begin with, we need to stay visible _
     (MDI forms don't have a `BorderStyle` property)
    If Not (TypeOf ParentForm Is MDIForm) Then
        If ParentForm.BorderStyle = VBRUN.FormBorderStyleConstants.vbBSNone _
            Then Exit Sub
    End If
    
    'Search the parent form for a bluBorderless control
    Dim VBControl As VB.Control
    For Each VBControl In ParentForm.Controls
        'Is this is a bluBorderless control?
        If (TypeOf VBControl Is bluBorderless) Then
            'Begin listening to its events
            Set bluBorderlessEvents = Nothing
            Set bluBorderlessEvents = VBControl
            'If we are min/max/close button, show or hide ourselves based on if _
             bluBorderless's borderless UI is active
            If My_Kind <> Sizer Then
                Let UserControl.Extender.Visible = bluBorderlessEvents.IsBorderless
            End If
            Exit For
        End If
    Next
    
    'If a sizer control, hide ourselves if the form is maximised
    Call ParentFormEvents_Resize
End Sub

'EVENT bluBorderlessEvents BORDERLESSSTATECHANGE _
 ======================================================================================
Private Sub bluBorderlessEvents_BorderlessStateChange(ByVal Enabled As Boolean)
    'If we are min/max/close button, show or hide ourselves based on if _
     bluBorderless's borderless UI is active
    If My_Kind <> Sizer Then
        Let UserControl.Extender.Visible = bluBorderlessEvents.IsBorderless
    End If
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY ActiveColour _
 ======================================================================================
Public Property Get ActiveColour() As OLE_COLOR: Let ActiveColour = My_ActiveColour: End Property
Public Property Let ActiveColour(ByVal NewColour As OLE_COLOR)
    Let My_ActiveColour = NewColour
    Call Refresh
    Call UserControl.PropertyChanged("ActiveColour")
End Property

'PROPERTY BaseColour _
 ======================================================================================
Public Property Get BaseColour() As OLE_COLOR: Let BaseColour = My_BaseColour: End Property
Public Property Let BaseColour(ByVal NewColour As OLE_COLOR)
    Let My_BaseColour = NewColour
    Call Refresh
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
    Call Refresh
    Call UserControl.PropertyChanged("Style")
End Property

'PROPERTY Kind : Which kind of control button we are _
 ======================================================================================
Public Property Get Kind() As bluControlBox_Kind: Let Kind = My_Kind: End Property
Public Property Let Kind(ByVal NewKind As bluControlBox_Kind)
    Let My_Kind = NewKind
    'The size of the control is fixed, when changing types call `Resize` to set the _
     correct size for the kind of control
    Call UserControl_Resize
    Call Refresh
    Call UserControl.PropertyChanged("Kind")
End Property

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'ReferenceParentForm : Set the references to the parent form & parent MDI form (if exists) _
 ======================================================================================
Private Sub ReferenceParentForm()
    'Find the parent form we need to control
    'NOTE: If your form is an MDI child, this will get the MDI parent. bluControlBox _
     only supports MDI interfaces in so much as a means to split functionality up _
     between forms rather than having your entire app on one form. It is assumed _
     that any MDI child is always maximised and that the child form is just one _
     interface of the whole app and not a 'document' window (don't use borderless UI _
     for that). Therefore bluControlBox will control the MDI parent, not the child _
     form when it comes to min / max / close / sizer
    Set ParentForm = blu.GetParentForm(UserControl.Parent, True)
    
    'NOTE: During compilation, the events won't bind, so we need to skip
    If blu.UserMode = False Then Exit Sub
    
    'Listen into the form events. For example, when the form is maximised, we can _
     hide the control if it's a size-box
    If (TypeOf ParentForm Is MDIForm) _
        Then Set ParentMDIFormEvents = ParentForm _
        Else Set ParentFormEvents = ParentForm
End Sub

'Refresh _
 ======================================================================================
Private Sub Refresh()
    'This isn't public as there's nothing the user can change that would require a _
     manual refresh. This is just here to force a repaint when properties are changed
    Call UserControl_Paint
    Call UserControl.Refresh
End Sub
