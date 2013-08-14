VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S1FloorLayout"
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
'CLASS :: S1FloorLayout

'What most would think of as a "level", this represents the layout of Blocks that _
 define the structure of the level you run and jump through. No graphics are stored _
 here, and nor do we reference the LevelArt directly either: Sonic 1 levels are _
 constructed of 4x4 Tile pre-defined "Blocks". This class holds a grid of indexes _
 to the Block Mappings and the Block Mappings define the 4x4 selection of tiles, and _
 the tiles are indexes to the actual graphics in the LevelArt, so a "level" has three _
 degrees of seperation from the graphics

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Public ID As String         'Our ID so we don't get lost in a collection
Public Title As String      'Friendly name to show in the level editor

Private Data() As Byte      'The actual floor layout data

Private My_Width As Long
Private My_Height As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Define a change event this class can provide so that the level class can recache
Event Change(ByVal X As Long, ByVal Y As Long, ByVal NewIndex As Byte, ByVal OldIndex As Byte)

'CLASS Terminate : Clean up our data _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Data
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////

'PROPERTY Block _
 ======================================================================================
Public Property Get Block(ByVal X As Long, ByVal Y As Long) As Byte
    Block = Data((Y * My_Width) + X)
End Property

Public Property Let Block(ByVal X As Long, ByVal Y As Long, ByVal Index As Byte)
    Dim Offset As Long
    Let Offset = (Y * My_Width) + X
    'If there's a change of data, raise the event to notify the level class
    If Index <> Data(Offset) Then
        Dim OldIndex As Byte
        Let OldIndex = Data(Offset)
        Let Data(Offset) = Index
        RaiseEvent Change(X, Y, Index, OldIndex)
    End If
End Property

'PROPERTY Width/Height _
 ======================================================================================
Public Property Get Width() As Long: Let Width = My_Width: End Property
Public Property Get Height() As Long: Let Height = My_Height: End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

Public Function GetByteStream() As Byte()
    Let GetByteStream = Data
End Function

Public Sub SetByteStream(NewData() As Byte)
    Let Data = NewData
End Sub

Public Sub Resize(ByVal NewWidth As Long, NewHeight As Long)
    'TODO: Preserve the old data
    ReDim Data(0 To (NewWidth * NewHeight) - 1) As Byte
    Let My_Width = NewWidth: Let My_Height = NewHeight
End Sub

