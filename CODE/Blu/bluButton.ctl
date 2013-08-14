VERSION 5.00
Begin VB.UserControl bluButton 
   AutoRedraw      =   -1  'True
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
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   97
   Begin MaSS1VE.bluLabel bluLabel 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Alignment       =   2
      State           =   1
      Style           =   1
   End
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
'Dependencies       blu.BAS, bluLabel.ctl, bluMagic.cls, bluMouseEvents.cls
'Last Updated       24-JUL-13

'NOTE: This is due to be rewritten to not depend upon a nested bluLabel control and _
 instead do all the painting ourselves. This will remove flicker and will stop the _
 whole project falling to pieces beacuse of changes in bluLabel cascading errors _
 through every blu control in the form!

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'The button flips the style (Normal / Invert) of the inner label when the user hovers _
 over the button. The user-facing state of the button that we expose acts as a lock, _
 keeping the button in a highlighted state
Private My_State As bluSTATE
Private My_Style As bluSTYLE

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

'CONTROL Initialize _
 ======================================================================================
Private Sub UserControl_Initialize()
    'Buttons do not use the text colour, they flip the base and active colours. _
     By setting the label to active, it will use the base/active colour for text, _
     instead of the usual text colour
    Let UserControl.bluLabel.State = Active
End Sub

'CONTROL InitProperties _
 ======================================================================================
Private Sub UserControl_InitProperties()
    Let Me.Caption = "Blu Button"
    Let Me.Orientation = bluORIENTATION.Horizontal
    Let Me.State = bluSTATE.Inactive
    Let Me.Style = bluSTYLE.Normal
End Sub

'CONTROL Paint _
 ======================================================================================
Private Sub UserControl_Paint()
    'We don't have to actually paint anything (that's handled in the label), but we _
     do need to manage which way the colours are around based on the button's _
     style, state and if hovered over
    'It's very important to note that despite changes in the active state of the _
     button, our internal label control is always active and has its style flipped
        
    'Is the button locked into an active, hovered state?
    If My_State = Active Then
        'Hover cannot affect this state so only one choice remains
        Let UserControl.bluLabel.Style = IIf( _
            Expression:=My_Style = Normal, _
            TruePart:=bluSTYLE.Invert, FalsePart:=bluSTYLE.Normal _
        )
    Else
        'The button is inactive, which means it can be hovered over, making it _
         temporarily active
        If IsHovered = True Then
            Let UserControl.bluLabel.Style = IIf( _
                Expression:=My_Style = Normal, _
                TruePart:=bluSTYLE.Invert, FalsePart:=bluSTYLE.Normal _
            )
        Else
            Let UserControl.bluLabel.Style = IIf( _
                Expression:=My_Style = Normal, _
                TruePart:=bluSTYLE.Normal, FalsePart:=bluSTYLE.Invert _
            )
        End If
    End If
End Sub

'CONTROL ReadProperties _
 ======================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Let Me.ActiveColour = .ReadProperty(Name:="ActiveColour", DefaultValue:=blu.ActiveColour)
        Let Me.Alignment = .ReadProperty(Name:="Alignment", DefaultValue:=VBRUN.AlignmentConstants.vbCenter)
        Let Me.BaseColour = .ReadProperty(Name:="BaseColour", DefaultValue:=blu.BaseColour)
        Let Me.Caption = .ReadProperty(Name:="Caption", DefaultValue:="Blu Button")
        Let Me.Orientation = .ReadProperty(Name:="Orientation", DefaultValue:=bluORIENTATION.Horizontal)
        Let Me.State = .ReadProperty(Name:="State", DefaultValue:=bluSTATE.Inactive)
        Let Me.Style = .ReadProperty(Name:="Style", DefaultValue:=bluSTYLE.Normal)
    End With
    
    If blu.UserMode = True Then
        'Attach the mouse tracking
        Set MouseEvents = New bluMouseEvents
        Let MouseEvents.MousePointer = IDC_HAND
        Call MouseEvents.Attach(UserControl.bluLabel.hWnd)
    End If
End Sub

'CONTROL Resize _
 ======================================================================================
Private Sub UserControl_Resize()
    Call UserControl.bluLabel.Move( _
        0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight _
    )
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
        Call .WriteProperty(Name:="ActiveColour", Value:=UserControl.bluLabel.ActiveColour, DefaultValue:=blu.ActiveColour)
        Call .WriteProperty(Name:="Alignment", Value:=UserControl.bluLabel.Alignment, DefaultValue:=VBRUN.AlignmentConstants.vbCenter)
        Call .WriteProperty(Name:="BaseColour", Value:=UserControl.bluLabel.BaseColour, DefaultValue:=blu.BaseColour)
        Call .WriteProperty(Name:="Caption", Value:=UserControl.bluLabel.Caption, DefaultValue:="Blu Button")
        Call .WriteProperty(Name:="Orientation", Value:=UserControl.bluLabel.Orientation, DefaultValue:=bluORIENTATION.Horizontal)
        Call .WriteProperty(Name:="State", Value:=My_State, DefaultValue:=bluSTATE.Inactive)
        Call .WriteProperty(Name:="Style", Value:=My_Style, DefaultValue:=bluSTYLE.Normal)
    End With
End Sub

'EVENT bluLabel CLICK : Expose the button click _
 ======================================================================================
Private Sub bluLabel_Click(): RaiseEvent Click: End Sub

'EVENT MouseEvents MOUSEIN : The mouse has entered the control _
 ======================================================================================
Private Sub MouseEvents_MouseIn()
    'The mouse is in the button, we'll cause a hover effect as long as the button is _
     not locked into an active state
    Let IsHovered = True
    Call UserControl_Paint
    RaiseEvent MouseIn
End Sub

'EVENT MouseEvents MOUSEOUT : The mouse has gone out of the control _
 ======================================================================================
Private Sub MouseEvents_MouseOut()
    'The mouse has left the button, we need to undo the hover effect, as long as the _
     button is not locked into an active state
    Let IsHovered = False
    Call UserControl_Paint
    RaiseEvent MouseOut
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY ActiveColour _
 ======================================================================================
Public Property Get ActiveColour() As OLE_COLOR
    Let ActiveColour = UserControl.bluLabel.ActiveColour
End Property

Public Property Let ActiveColour(ByVal NewColour As OLE_COLOR)
    Let UserControl.bluLabel.ActiveColour = NewColour
    Call UserControl.PropertyChanged("ActiveColour")
End Property

'PROPERTY Alignment : Text alignment (left / center / right) _
 ======================================================================================
Public Property Get Alignment() As VBRUN.AlignmentConstants
    Let Alignment = UserControl.bluLabel.Alignment
End Property

Public Property Let Alignment(ByVal NewAlignment As VBRUN.AlignmentConstants)
    Let UserControl.bluLabel.Alignment = NewAlignment
    Call UserControl.PropertyChanged("Alignment")
End Property

'PROPERTY BaseColour _
 ======================================================================================
Public Property Get BaseColour() As OLE_COLOR
    Let BaseColour = UserControl.bluLabel.BaseColour
End Property

Public Property Let BaseColour(ByVal NewColour As OLE_COLOR)
    Let UserControl.bluLabel.BaseColour = NewColour
    Call UserControl.PropertyChanged("BaseColour")
End Property

'PROPERTY Caption _
 ======================================================================================
Public Property Get Caption() As String
    Let Caption = UserControl.bluLabel.Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    Let UserControl.bluLabel.Caption = NewCaption
    Call UserControl.PropertyChanged("Caption")
End Property

'PROPERTY Orientation _
 ======================================================================================
Public Property Get Orientation() As bluORIENTATION
    Let Orientation = UserControl.bluLabel.Orientation
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
    
    Let UserControl.bluLabel.Orientation = NewOrientation
    Call UserControl.PropertyChanged("Orientation")
End Property

'PROPERTY State : Can be used to lock the button into a hovered state _
 ======================================================================================
'We won't rely on the inner label state as this flips about on hover
Public Property Get State() As bluSTATE: Let State = My_State: End Property
Public Property Let State(ByVal NewState As bluSTATE)
    Let My_State = NewState
    Call UserControl_Paint
    Call UserControl.PropertyChanged("State")
End Property

'PROPERTY Style _
 ======================================================================================
Public Property Get Style() As bluSTYLE: Let Style = My_Style: End Property
Public Property Let Style(ByVal NewStyle As bluSTYLE)
    Let My_Style = NewStyle
    Call UserControl_Paint
    Call UserControl.PropertyChanged("Style")
End Property
