VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S1BlockMapping"
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
'CLASS :: S1BlockMapping

'The Floor Layout is made up of 32x32px blocks that compose the background of the _
 level. Each block is made up of a 4x4 grid of tiles and the tiles come from the _
 Level Art. The Block Mappings define the set of 4x4 tile arrangements that construe _
 the Blocks that the Floor Layout uses

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Public Title As String
Public ID As String

Private Blocks() As Byte

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Define a change event this class can provide so that the editor can recache
Event Change(ByVal BlockIndex As Byte, ByVal TileIndex As Byte, ByVal Value As Byte)

'CLASS Terminate : Clean up our data _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Blocks
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////

'PROPERTY Block : Read and write the block mapping data _
 ======================================================================================
Public Property Get Block(ByVal Index As Byte) As Byte()
    Dim Output(0 To 15) As Byte
    
    Dim i As Long
    For i = 0 To 15: Output(i) = Blocks(i, Index): Next
    
    Let Block = Output
End Property

Public Property Let Block(ByVal Index As Byte, ByRef Tiles() As Byte)
    Dim i As Long
    For i = 0 To 15: Blocks(i, Index) = Tiles(i): Next
End Property

'PROPERTY Length : The number of block mappings _
 ======================================================================================
Public Property Get Length() As Byte
    Let Length = UBound(Blocks, 2)
End Property

Public Property Let Length(ByVal Indicies As Byte)
    ReDim Preserve Blocks(0 To 15, Indicies) As Byte
End Property

'PROPERTY Tile : Read and write individual tiles within block mappings _
 ======================================================================================
Public Property Get Tile(ByVal BlockIndex As Byte, ByVal TileIndex As Byte) As Byte
    Let Tile = Blocks(TileIndex, BlockIndex)
End Property

Public Property Let Tile(ByVal BlockIndex As Byte, ByVal TileIndex As Byte, ByVal Value As Byte)
    If Value <> Blocks(TileIndex, BlockIndex) Then
        Let Blocks(TileIndex, BlockIndex) = Value
        RaiseEvent Change(BlockIndex, TileIndex, Value)
    End If
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'RingsInBlock : How many rings are there in a block? _
 ======================================================================================
Public Function RingsInBlock(ByVal Index As Byte) As Long
    'TODO this information should be cached, and updated in the property changes
    'Rings are marked by tiles &HFC/FD/FE/FF in the block
    Dim i As Long
    For i = 0 To 15
        If Me.Tile(Index, i) = &HFF& Then Let RingsInBlock = RingsInBlock + 1
    Next i
End Function
