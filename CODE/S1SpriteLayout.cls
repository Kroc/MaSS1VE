VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S1SpriteLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: S1SpriteLayout

'The ROM stores objects as a pattern of sprites arranged in a 6x3 grid (though each _
 sprite itself is double-height so that the object appears as a 6x6 grid of 8x8 tiles)

'Note that this class handles the typical single-palette, single-tileset objects; _
 for complex objects such as the monitors or meta-objects, these are handled elsewhere

'TODO: Flame Thrower (&H16) will require special construction
'TODO: Caterkiller (&H1F) will require special construction
'TODO: Boss - Scrap Brain (&H22) will require special construction
'TODO: The capsule (&H25) is incomplete
'TODO: Propeller (&H31) will require special construction
'TODO: Rotating Turret (&H37) will require special construction
'TODO: Spear (&H3D) will require special construction
'TODO: Water line (&H40) will require a symbol
'TODO: Bubbles (&H41) will require special construction
'TODO: Boss - Electric Beam (&H46) will require special construction
'TODO: Trip zone (&H4B) will require a symbol
'TODO: Balance (&H4E) will require special construction
'TODO: Flower (&H50) will require special construction

Private Cache As bluImage

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Public Title As String                  'Friendly name of the object

Private My_Layout(0 To 17) As Byte      'This stores the 6x3 layout of sprite indexes
Private My_Palette As S1Palette         'Which palette to use for the sprite
Private My_Tileset As S1Tileset         'Which tileset to get the graphics from

'A hint for special kinds of objects that will need extra work
Public Enum S1SpriteLayout_Kind
    SL_Normal = 0                        'No special requirements
    SL_Monitor_Rings = &H1
    SL_Monitor_Speed = &H2
    SL_Monitor_Life = &H3
    SL_Monitor_Shield = &H4
    SL_Monitor_Stars = &H5               'Invincibility
    SL_Monitor_Check = &H51              'Checkpoint
    SL_Monitor_Cont = &H52               'Continue
    SL_Emerald = &H6
End Enum

Public Kind As S1SpriteLayout_Kind

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY hDC : Gets the device context of the cache image _
 ======================================================================================
Public Property Get hDC() As Long
    'If there's no cache yet, create it now
    Call CacheImage
    'Manual painting in a tight loop will be much faster than using the `Paint` _
     method in this class, though remember to cache this hDC for your loop as _
     object-negotiation is very slow in VB6
    Let hDC = Cache.hDC
End Property


'PROPERTY Layout : Define the sprite layout _
 ======================================================================================
Public Property Let Layout(ByVal PositionIndex As Long, ByVal SpriteIndex As Byte)
    Let My_Layout(PositionIndex) = SpriteIndex
    'Invalidate the cache
    Set Cache = Nothing
End Property

'PROPERTY Palette : Set the palette to use for the sprites _
 ======================================================================================
Public Property Get Palette() As S1Palette: Set Palette = My_Palette: End Property
Public Property Set Palette(ByRef NewPalette As S1Palette)
    Set My_Palette = NewPalette
    'Invalidate the cache
    Set Cache = Nothing
End Property

'PROPERTY Tileset : Set the tileset the sprite indexes refer to _
 ======================================================================================
Public Property Get Tileset() As S1Tileset: Set Tileset = My_Tileset: End Property
Public Property Set Tileset(ByRef NewTileset As S1Tileset)
    Set My_Tileset = NewTileset
    'Invalidate the cache
    Set Cache = Nothing
End Property

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'The layout defaults to $FF throughout (blank space)
    Dim Index As Long
    For Index = LBound(My_Layout) To UBound(My_Layout)
        Let My_Layout(Index) = &HFF
    Next Index
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Derefernce
    Set My_Palette = Nothing
    Set My_Tileset = Nothing
    'Clear the cache and data
    Set Cache = Nothing
    Erase My_Layout
End Sub

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Paint : Paint the sprite layout onto the specified surface, caching if needed _
 ======================================================================================
Public Function Paint( _
    ByVal hndDeviceContext As Long, _
    ByVal X As Long, ByVal Y As Long _
) As Long
    'If there's no cache yet, create it now. Any further paints will be much faster
    Call CacheImage
    
    Select Case Me.Kind
        Case SL_Monitor_Rings, SL_Monitor_Speed, SL_Monitor_Life, SL_Monitor_Shield, _
             SL_Monitor_Stars, SL_Monitor_Check, SL_Monitor_Cont
            Let X = X + 4: Let Y = Y - 8
        Case SL_Emerald
            Let X = X + 8
    End Select
    
    Let Paint = Cache.Paint(hndDeviceContext, X, Y)
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'CacheImage _
 ======================================================================================
Private Sub CacheImage()
    'If it's already cached, leave. _
     Changing the layout / tileset / palette will invalidate the cache
    If Not Cache Is Nothing Then Exit Sub
    
    'Create the image and set the size & palette
    Set Cache = New bluImage
    Call Cache.Create8Bit( _
        ImageWidth:=8 * 6, ImageHeight:=16 * 3, _
        Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
    )
    
    'Loop over the sprite layout definition
    Dim X As Long, Y As Long
    For Y = 0 To 2: For X = 0 To 5
        'Ignore sprite indexes of $FE/$FF as these represent a blank space
        Dim Index As Long: Let Index = My_Layout((Y * 6) + X)
        If Index < &HFE Then
            'Indexes 128 or over come from the shared HUD sprites
            If Index >= 128 Then
                Call GAME.HUD.PaintSprite(Cache.hDC, X * 8, Y * 16, Index - 128)
                
            'Indexes $5C & $5E are the 'current' power-up, the emerald uses this _
             specifically in its sprite layout
            ElseIf Index = &H5C Or Index = &H5E Then
                Call GAME.PowerUps.ApplyPalette(My_Palette)
                Let Index = Index - (&H5C - 20)
                'The power ups are positioned slightly off from normal
                Call GAME.PowerUps.PaintSprite(Cache.hDC, X * 8, Y * 16, Index)
                
            Else
                'Paint the sprite into the cache image
                Call My_Tileset.PaintSprite(Cache.hDC, X * 8, Y * 16, Index)
            End If
        End If
    Next X: Next Y
    
    Select Case Me.Kind
        Case S1SpriteLayout_Kind.SL_Monitor_Rings:
            Call GAME.PowerUps.ApplyPalette(My_Palette)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4, 0, 0)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4 + 8, 0, 2)
        Case S1SpriteLayout_Kind.SL_Monitor_Speed:
            Call GAME.PowerUps.ApplyPalette(My_Palette)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4, 0, 4)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4 + 8, 0, 6)
        Case S1SpriteLayout_Kind.SL_Monitor_Life:
            Call GAME.PowerUps.ApplyPalette(My_Palette)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4, 0, 8)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4 + 8, 0, 10)
        Case S1SpriteLayout_Kind.SL_Monitor_Shield:
            Call GAME.PowerUps.ApplyPalette(My_Palette)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4, 0, 12)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4 + 8, 0, 14)
        Case S1SpriteLayout_Kind.SL_Monitor_Stars:
            Call GAME.PowerUps.ApplyPalette(My_Palette)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4, 0, 16)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4 + 8, 0, 18)
        Case S1SpriteLayout_Kind.SL_Monitor_Check:
            Call GAME.PowerUps.ApplyPalette(My_Palette)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4, 0, 24)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4 + 8, 0, 26)
        Case S1SpriteLayout_Kind.SL_Monitor_Cont:
            Call GAME.PowerUps.ApplyPalette(My_Palette)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4, 0, 28)
            Call GAME.PowerUps.PaintSprite(Cache.hDC, 4 + 8, 0, 30)
    End Select
End Sub
