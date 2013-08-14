VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S1ObjectLayout"
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
'CLASS :: S1ObjectLayout

'Where the objects (effectively sprites) are in the level

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Public ID As String         'Our ID so we don't get lost in a collection
Public Title As String      'Friendly name to show in the level editor

'The actual object layout data: _
 (The `OBJ` type cannot be defined public in a class due to limitations in VB)
Private Objects(0 To 255) As ROM.OBJ

'The height of the water line (like in Labyrinth). This is dertermined by the presence _
 of ObjID &H40, where the Y value sets the water level _
 ("256" for no water level; below the maximum level height)
Public WaterLevel As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Define a change event this class can provide so that the level editor can recache
Event Change(ByVal X As Long, ByVal Y As Long, ByVal ObjID As OBJECT_TYPE)
'When the water level changes, so the level can redraw accordingly
Event WaterLevelChange()

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Begin with no water level (set it below the maximum level height)
    Let WaterLevel = 256
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Objects
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////

'PROPERTY Object : Read from the object list _
 ======================================================================================
Friend Property Get Object(ByVal Index As Byte) As ROM.OBJ
    Let Object = Objects(Index)
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Add : Add an object to the list _
 ======================================================================================
Public Function Add(ByVal ObjID As OBJECT_TYPE, ByVal X As Byte, ByVal Y As Byte) As Boolean
    'In any instance, ignore ObjIDs of 0 / 255, these are null objects
    If ObjID = 0 Or ObjID = 255 Then Exit Function
    
    'Find the next available object slot
    Dim Index As Long
    Let Index = NextIndex()
    'Have we run out of objects?!
    If Index = -1 Then Let Add = False: Exit Function
    
    'Assign the data and alert the controller
    Let Objects(Index).O = ObjID
    Let Objects(Index).X = X
    Let Objects(Index).Y = Y
    'TODO: The controller won't know to repaint multiple objects on the same cell
    RaiseEvent Change(X, Y, ObjID)
    
    'Is this the water level marker?
    If ObjID = META_WATER Then
        Let WaterLevel = Y
        'Notify the controller (i.e. the level can redraw)
        RaiseEvent WaterLevelChange
    End If
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'NextIndex : Find the next available slot for an object _
 ======================================================================================
Private Function NextIndex() As Long
    'This isn't exactly efficient, but it's only 256 items (max), so should be okay
    Dim i As Long
    For i = 0 To 255
        If Objects(i).O = 0 Then Let NextIndex = i: Exit Function
    Next
    Let NextIndex = -1
End Function
