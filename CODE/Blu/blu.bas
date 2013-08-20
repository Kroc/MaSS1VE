Attribute VB_Name = "blu"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Blu

'Stuff shared between the Blu ActiveX controls

'/// PUBLIC VARS //////////////////////////////////////////////////////////////////////

'When a user control is nested in another user control, the `Ambient.UserMode` _
 property returns the incorrect value of True when the control is being run in _
 "Design Mode" (on the form editor). This would cause the design mode controls _
 to be subclassed and crashes the IDE. To stop this, the variable below will _
 always be False when the controls are running in Design Mode. Set `UserMode` _
 to True in your `Sub Main()` to tell the controls it's okay to subclass. _
 (`Sub Main()` will only be run when you your app runs, not during design time)
Public UserMode As Boolean

'The default measurement (px) to base control layout around. _
 Use `blu.Xpx(blu.Metric)` / `blu.Ypx(blu.Metric)` to get it in Twips
Public Const Metric As Long = 32

'The default colour palette for our controls
Public Const BaseColour As Long = vbWhite
Public Const BaseHoverColour As Long = &HEEEEEE
Public Const TextColour As Long = &H999999
Public Const TextHoverColour As Long = &H666666
Public Const ActiveColour As Long = &HFFAF00
Public Const InertColour As Long = &HFFEABA

'The close control box button is red unlike the others
Public Const CloseHoverColour As Long = &H4343E0
Public Const ClosePressColour As Long = &H5050C7 '&H3D3D99

'Public Enums _
 --------------------------------------------------------------------------------------
'The Blu ActiveX controls use these to define friendly names for some properties

Public Enum bluORIENTATION
    Horizontal = 0
    VerticalUp = 1
    VerticalDown = 2
End Enum

Public Enum bluSTATE
    Inactive = 0
    Active = 1
End Enum

Public Enum bluSTYLE
    Normal = 0
    Invert = 1
End Enum

Public Enum bluTRUNCATE
    NoTruncation = 0
    EndElipsis = 1
    MiddleElipsis = 2
End Enum

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'ApplyColoursToForm : Change the colour scheme of the form controls _
 ======================================================================================
Public Sub ApplyColoursToForm( _
    ByRef TheForm As Object, _
    Optional ByVal BaseColour As OLE_COLOR = BaseColour, _
    Optional ByVal TextColour As OLE_COLOR = TextColour, _
    Optional ByVal ActiveColour As OLE_COLOR = ActiveColour, _
    Optional ByVal InertColour As OLE_COLOR = InertColour _
)
    'Deal with all blu controls automatically
    Dim FormControl As Control
    For Each FormControl In TheForm.Controls
        If (TypeOf FormControl Is bluLabel) _
        Or (TypeOf FormControl Is bluButton) _
        Or (TypeOf FormControl Is bluTab) _
        Then
            With FormControl
                On Error Resume Next
                Let .BaseColour = BaseColour
                Let .TextColour = TextColour
                Let .ActiveColour = ActiveColour
                Let .InertColour = InertColour
            End With
        End If
    Next FormControl
End Sub

'DrawText : Shared routine for drawing text, used by bluLabel/Button/Tab &c. _
 ======================================================================================
Public Sub DrawText( _
    ByVal hndDeviceContext As Long, _
    ByRef BoundingBox As RECT, _
    ByVal Text As String, _
    ByVal Colour As OLE_COLOR, _
    Optional ByVal Alignment As VBRUN.AlignmentConstants, _
    Optional ByVal Orientation As bluORIENTATION = Horizontal _
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
    Let hndOld = WIN32.gdi32_SelectObject(hndDeviceContext, hndFont)
    
    'The `DrawText` API doesn't work with the position set by `SetTextAlign`, _
     so we ensure it's set to a safe, non-interfering value
    Call WIN32.gdi32_SetTextAlign( _
        hndDeviceContext, TA_TOP Or TA_LEFT Or TA_NOUPDATECP _
    )
    
    'Rotate the text? _
     ----------------------------------------------------------------------------------
    'Made possible with directions from: <edais.mvps.org/Tutorials/GDI/DC/DCch8.html>
    If Orientation <> Horizontal Then
        'Determine the centre point of the bounding box before we begin to modify it
        Dim Centre As POINT
        Let Centre.X = BoundingBox.Left + (BoundingBox.Right - BoundingBox.Left) \ 2
        Let Centre.Y = BoundingBox.Top + (BoundingBox.Bottom - BoundingBox.Top) \ 2
        
        'The button is already in a vertical shape, but we want to rotate a horizontal _
         piece of text, so we have to swap the dimensions of the button to begin with
        Call WIN32.user32_SetRect( _
            BoundingBox, _
            BoundingBox.Left, BoundingBox.Top, BoundingBox.Bottom, BoundingBox.Right _
        )
        'In addition to that, we also need to position our text with its center at _
         0,0, instead of the top-left corner, so that when we rotate, the text stays _
         centered and doesn't swing off out of place
        Call WIN32.user32_OffsetRect( _
            BoundingBox, -BoundingBox.Right \ 2, -BoundingBox.Bottom \ 2 _
        )
        
        'Now we need to move the origin point (0,0) to the centre of the button _
         so that the rotated text obviously appears in the center of the button _
         whilst the rotation occurs around the centrepoint of the text
        Dim Org As POINT
        Call WIN32.gdi32_SetViewportOrgEx(hndDeviceContext, Centre.X, Centre.Y, Org)
        
        'In order to use Get/SetWorldTransform we have to make this call
        Dim OldGM As Long
        Let OldGM = WIN32.gdi32_SetGraphicsMode(hndDeviceContext, GM_ADVANCED)
        
        'Now calculate the rotation
        Const Pi As Single = 3.14159
        Dim RotAng As Single
        Let RotAng = IIf(Orientation = VerticalDown, -90, 90)
        Dim RotRad As Single
        Let RotRad = (RotAng / 180) * Pi
        
        'Read any current transform from the device context
        Dim OldXForm As XFORM, RotXForm As XFORM
        Call WIN32.gdi32_GetWorldTransform(hndDeviceContext, OldXForm)
        
        'Define our rotation matrix
        With RotXForm
            Let .eM11 = Cos(RotRad)
            Let .eM21 = Sin(RotRad)
            Let .eM12 = -.eM21
            Let .eM22 = .eM11
        End With
        
        'Apply the matrix -- rotate the world!
        Call WIN32.gdi32_SetWorldTransform(hndDeviceContext, RotXForm)
    End If
    
    'Draw the text! _
     ----------------------------------------------------------------------------------
    'Select the colour of the text
    Call WIN32.gdi32_SetTextColor( _
        hndDeviceContext, WIN32.OLETranslateColor(Colour) _
    )
    
    'Add a little padding either side
    With BoundingBox
        Let .Left = .Left + 8
        Let .Right = .Right - 8
    End With

    'Now just paint the text
    Call WIN32.user32_DrawText( _
        hndDeviceContext:=hndDeviceContext, _
        Text:=Text, Length:=Len(Text), _
        BoundingBox:=BoundingBox, _
        Format:=Choose(Alignment + 1, DT.DT_LEFT, DT.DT_RIGHT, DT.DT_CENTER) _
                Or DT_VCENTER Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP _
    )
    
    'Clean up: _
     ----------------------------------------------------------------------------------
    'If we rotated the text, we need to do some additional clean up
    If Orientation <> Horizontal Then
        'Restore the previous world transform
        Call WIN32.gdi32_SetWorldTransform(hndDeviceContext, OldXForm)
        'Switch back to the previous graphics mode
        Call WIN32.gdi32_SetGraphicsMode(hndDeviceContext, OldGM)
        'Return the origin point (0,0) back to the upper-left corner
        Call WIN32.gdi32_SetViewportOrgEx(hndDeviceContext, Org.X, Org.Y, Org)
    End If
    
    'Select the previous object into the DC (i.e. unselect the font)
    Call WIN32.gdi32_SelectObject(hndDeviceContext, hndOld)
    Call WIN32.gdi32_DeleteObject(hndFont)
End Sub

'Xpx : Shorthand for a number of horizontal pixels converted to twips _
 ======================================================================================
Public Function Xpx(Optional ByVal px As Long = 1) As Long
    'Yes, we could use `Form.ScaleX (...)` but this doesn't require a form and is _
     shorter to write
    Let Xpx = Screen.TwipsPerPixelX * px
End Function

'Ypx : Shorthand for a number of vertical pixels converted to twips _
 ======================================================================================
Public Function Ypx(Optional ByVal px As Long = 1) As Long
    Let Ypx = Screen.TwipsPerPixelY * px
End Function
