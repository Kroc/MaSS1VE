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

Private Cache As bluImage

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Private My_Layout(0 To 17) As Byte      'This stores the 6x3 layout of sprite indexes
Private My_Palette As S1Palette         'Which palette to use for the sprite
Private My_Tileset As S1Tileset         'Which tileset to get the graphics from

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY hDC : Gets the device context of the cache image _
 ======================================================================================
Public Property Get hDC() As Long
    'If there's no cache yet, create it now
    Call CacheImage
    'Manual painting in a tight loop will be much faster than using the `Paint` _
     method here, though remember to cache this hDC for your loop as _
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
Public Property Set Palette(NewPalette As S1Palette)
    Set My_Palette = NewPalette
    'Invalidate the cache
    Set Cache = Nothing
End Property

'PROPERTY Tileset : Set the tileset the sprite indexes refer to _
 ======================================================================================
Public Property Get Tileset() As S1Tileset: Set Tileset = My_Tileset: End Property
Public Property Set Tileset(NewTileset As S1Palette)
    Set My_Tileset = NewTileset
    'Invalidate the cache
    Set Cache = Nothing
End Property

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

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
    'Use the bluImage's built-in paint method
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
        ImageWidth:=6 * 8, ImageHeight:=16 * 3, _
        Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
    )
    
    'Loop over the sprite layout definition
    Dim X As Long, Y As Long
    For Y = 0 To 3: For X = 0 To 5
        'Ignore sprite indexes of $FF as these represent a blank space
        Dim Index As Long: Let Index = (Y * 6) + X
        If My_Layout(Index) < &HFF Then
            'Paint the sprite into the cache image
            Call Tileset.PaintSprite( _
                Cache.hDC, X * 8, Y * 16, _
                My_Layout((Y * 6) + X) _
            )
        End If
    Next X: Next Y
End Sub
