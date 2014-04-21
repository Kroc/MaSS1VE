VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S1Tileset"
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
'CLASS :: S1Tileset

'The actual graphics for something, as a long strip of 128 or 256 8x8px tiles

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'Our ID so that we don't get lost in a Collection
Public ID As String
'Friendly name to show in the editor
Public Title As String

Public Tiles As bluImage

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Call Me.Create(256)
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Set Tiles = Nothing
End Sub

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'ApplyPalette _
 ======================================================================================
Public Sub ApplyPalette(ByRef Palette As S1Palette)
    Dim i As Long
    For i = 0 To 15: Let Tiles.Colour(i) = Palette.Colour(i): Next
End Sub

'Create _
 ======================================================================================
Public Sub Create( _
    ByVal NumberOfTiles As Long, _
    Optional ByRef Palette As S1Palette = Nothing, _
    Optional ByVal UseTransparency As Boolean = False _
)
    Dim Palette_LongArray As Variant
    If Palette Is Nothing Then
        Let Palette_LongArray = Null
    Else
        Let Palette_LongArray = Palette.Colours
    End If
    
    Set Tiles = New bluImage
    Call Tiles.Create8Bit( _
        ImageWidth:=8 * NumberOfTiles, ImageHeight:=8, _
        Palette_LongArray:=Palette_LongArray, UseTransparency:=UseTransparency _
    )
End Sub

'PaintSprite _
 ======================================================================================
Public Sub PaintSprite(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal FirstIndex As Byte)
    'Master system sprites can be two tiles high, so often appear in the tileset as _
     two tiles next to each other designed to be stacked vertically. This procedure _
     saves some effort by painting the two tiles vertically in one call
    Call Tiles.Paint(hDC, X, Y, 8, 8, FirstIndex * 8, 0)
    Call Tiles.Paint(hDC, X, Y + 8, 8, 8, (FirstIndex + 1) * 8, 0)
End Sub

'PaintTile _
 ======================================================================================
Public Sub PaintTile(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Index As Byte)
    Call Tiles.Paint( _
        hndDeviceContext:=hDC, _
        DestLeft:=X, DestTop:=Y, _
        DestWidth:=8, DestHeight:=8, _
        SrcLeft:=Index * 8, SrcTop:=0 _
    )
End Sub
