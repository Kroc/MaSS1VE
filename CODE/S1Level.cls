VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S1Level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: S1Level

'A "level" in Sonic 1 is a pick-n-mix of many different elements brought together. _
 The level's "theme" is a combination of LevelArt, Palette, Object Layout and other _
 factors, all of which can freely be mixed in other combinations in other levels. _
 For example, one could have Night and Day levels that share the same LevelArt, but _
 use a different palette. Levels can share the same Floor Layout, but have different _
 object layouts, so when it comes to editing a "level", it's not guaranteed that each _
 level is entirely isolated from the others (as we would imagine most level editors _
 today). Therefore this level class is more of a "junction box", joining the various _
 elements together

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'Friendly name to show in the level editor
Public Title As String

Private My_LevelPalette As S1Palette
Private My_SpritePalette As S1Palette
Private My_LevelArt As S1Tileset
Private My_SpriteArt As S1Tileset
Private WithEvents My_BlockMapping As S1BlockMapping
Attribute My_BlockMapping.VB_VarHelpID = -1
Private WithEvents My_FloorLayout As S1FloorLayout
Attribute My_FloorLayout.VB_VarHelpID = -1
Private WithEvents My_ObjectLayout As S1ObjectLayout
Attribute My_ObjectLayout.VB_VarHelpID = -1

'Sonic's starting position in the level
Public StartX As Byte, StartY As Byte

'Whether the level has a water line, i.e. Labyrinth. Where the water line occurs is _
 defined in the Object Layout. An Object ID of &H40 will set the water line at its _
 Y-position
Private My_IsUnderWater As Boolean

'The number of rings available on the level
Private My_Rings As Long

'Stuff from the ROM Level Header that isn't integrated yet, so we have to store it to _
 export a ROM. This will go away once this stuff is integrated as ROM data shouldn't _
 be here -- the import process abstracts it away
'Documentation for these variables can be found here: _
 <info.sonicretro.org/SCHG:Sonic_the_Hedgehog_%288-bit%29#Level_Headers>
Public ROM_SP As Byte       'Solidity pointer (what tiles are solid / passable)
Public ROM_LX As Long       'Level X Offset
Public ROM_LW As Long       'Level Width
Public ROM_LY As Long       'Level Y Offset
Public ROM_LH As Long       'Level Height
Public ROM_BM As Long       'Pointer from $10000 to the Block Mapping ($10000)
Public ROM_LA As Long       'Pointer from $30000 to the Level Art ($32FE6)
Public ROM_SA As Long       'Pointer from $24000 to the Sprite Art ($2A12A)
Public ROM_IP As Byte       'Initial palette index
Public ROM_CS As Byte       'Cycle speed
Public ROM_CC As Byte       'Number of palette colour cycles
Public ROM_CP As Byte       'Cycle palette index
Public ROM_OL As Long       'Pointer from $15580 to the Object Layout ($15AB4)
Public ROM_SR As Byte       'Scrolling and Ring HUD flags
Public ROM_TL As Byte       'Time and Lightning flags
Public ROM_MU As Byte       'Music

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Notify the editor of changes to the block mappings
Event BlockMappingChange(ByVal BlockIndex As Byte, ByVal TileIndex As Byte, ByVal Value As Byte)
'Notify the editor of changes to the floor layout
Event FloorLayoutChange(ByVal X As Long, ByVal Y As Long, ByVal NewIndex As Byte, ByVal OldIndex As Byte)

'Allow the editor to respond to changes to the number of rings in the level
Event RingCountChange()
'Notify when the water level changes
Event WaterLevelChange()

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Clean up
    Set My_FloorLayout = Nothing
    Set My_BlockMapping = Nothing
    Set My_SpriteArt = Nothing
    Set My_ObjectLayout = Nothing
    Set My_LevelArt = Nothing
    Set My_LevelPalette = Nothing
    Set My_SpritePalette = Nothing
End Sub

'EVENT My_BlockMapping CHANGE : The block mappings have changed _
 ======================================================================================
Private Sub My_BlockMapping_Change(ByVal BlockIndex As Byte, ByVal TileIndex As Byte, ByVal Value As Byte)
    'Pass this upwards to the editor who can recache
    RaiseEvent BlockMappingChange(BlockIndex, TileIndex, Value)
End Sub

'EVENT My_FloorLayout CHANGE : The floor layout has been modified -- cache the image _
 ======================================================================================
Private Sub My_FloorLayout_Change(ByVal X As Long, ByVal Y As Long, ByVal NewIndex As Byte, ByVal OldIndex As Byte)
    'If the block mappings are not attached yet (not likely) then we can't check _
     ring counts
    If Not My_BlockMapping Is Nothing Then
        'Determine how many rings have been added / removed
        Dim RingDiff As Long
        Let RingDiff = ( _
            My_BlockMapping.RingsInBlock(NewIndex) - _
            My_BlockMapping.RingsInBlock(OldIndex) _
        )
        'Has the amount of rings on the level changed?
        If RingDiff <> 0 Then
            Let My_Rings = My_Rings + RingDiff
            RaiseEvent RingCountChange
        End If
    End If
    'Pass this on to the editor who can repaint
    RaiseEvent FloorLayoutChange(X, Y, NewIndex, OldIndex)
End Sub

'EVENT My_ObjectLayout WATERLEVELCHANGE : Some action changed the water level _
 ======================================================================================
Private Sub My_ObjectLayout_WaterLevelChange()
    'The ObjectLayout will trigger an event when object &H40 is set which controls the _
     where the water line is in the level (i.e. Labyrinth). Pass this up to the editor _
     so it can repaint
    RaiseEvent WaterLevelChange
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////

'PROPERTY BlockMapping _
 ======================================================================================
Public Property Get BlockMapping() As S1BlockMapping: Set BlockMapping = My_BlockMapping: End Property
Public Property Set BlockMapping(ByRef NewBlockMapping As S1BlockMapping)
    'TODO: Shouldn't the palette be fixed on this?
    Set My_BlockMapping = NewBlockMapping
    'Re-count the number of rings on the level
    Let My_Rings = CountRings()
End Property

'PROPERTY FloorLayout _
 ======================================================================================
Public Property Get FloorLayout() As S1FloorLayout: Set FloorLayout = My_FloorLayout: End Property
Public Property Set FloorLayout(ByRef NewFloorLayout As S1FloorLayout)
    'TODO: Update the level header width/height
    Set My_FloorLayout = NewFloorLayout
    'Re-count the number of rings on the level
    Let My_Rings = CountRings()
End Property

'PROPERTY IsUnderWater : If the level has a water line, i.e. Labyrinth _
 ======================================================================================
Public Property Get IsUnderWater() As Boolean: Let IsUnderWater = My_IsUnderWater: End Property
Public Property Let IsUnderWater(ByVal State As Boolean)
    Let My_IsUnderWater = State
End Property

'PROPERTY LevelArt _
 ======================================================================================
Public Property Get LevelArt() As S1Tileset: Set LevelArt = My_LevelArt: End Property
Public Property Set LevelArt(ByRef NewLevelArt As S1Tileset)
    Set My_LevelArt = NewLevelArt
    'Fix the level art's palette to match the palette attached to this level
    Call ApplyPalette
End Property

'PROPERTY LevelPalette : Apply a tile palette to the level _
 ======================================================================================
Public Property Get LevelPalette() As S1Palette: Set LevelPalette = My_LevelPalette: End Property
Public Property Set LevelPalette(ByRef NewPalette As S1Palette)
    Set My_LevelPalette = NewPalette
    'Apply this palette to the internal image palette in the level art
    Call ApplyPalette
End Property

'PROPERTY ObjectLayout _
 ======================================================================================
Public Property Get ObjectLayout() As S1ObjectLayout: Set ObjectLayout = My_ObjectLayout: End Property
Public Property Set ObjectLayout(ByRef NewObjectLayout As S1ObjectLayout)
    Set My_ObjectLayout = NewObjectLayout
    'Re-count the number of rings on the level
    Let My_Rings = CountRings()
End Property

'PROPERTY Rings : The number of rings available on the level _
 ======================================================================================
Public Property Get Rings() As Long: Let Rings = My_Rings: End Property

'PROPERTY SpriteArt : Apply a tile palette to the level _
 ======================================================================================
Public Property Get SpriteArt() As S1Tileset: Set SpriteArt = My_SpriteArt: End Property
Public Property Set SpriteArt(ByRef NewSpriteArt As S1Tileset)
    Set My_SpriteArt = NewSpriteArt
    'Fix the sprite art's palette to match the palette attached to this level
    Call ApplyPalette
End Property

'PROPERTY SpritePalette : Apply a tile palette to the level _
 ======================================================================================
Public Property Get SpritePalette() As S1Palette: Set SpritePalette = My_SpritePalette: End Property
Public Property Set SpritePalette(ByRef NewPalette As S1Palette)
    Set My_SpritePalette = NewPalette
    'Apply this palette to the internal image palette in the sprite art
    Call ApplyPalette
End Property

'PROPERTY Width / Height _
 ======================================================================================
Public Property Get Width() As Long: Let Width = My_FloorLayout.Width: End Property
Public Property Get Height() As Long: Let Height = My_FloorLayout.Height: End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Resize _
 ======================================================================================
Public Sub Resize(Width As Long, Height As Long)
    'Resize the floor layout
    'TODO: Preserve the old data
    Call My_FloorLayout.Resize(Width, Height)
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'ApplyPalette : When palettes / layouts / mappings are swapped around, fix the palette _
 ======================================================================================
Private Sub ApplyPalette()
    Dim i As Long
    'Because all the components of a level can be swapped about, and shared with other _
     levels, the palette inside the level art's image has to be changed to match with _
     the level's chosen palette whenever a change occurs
     
    'These objects might not be available since they get attached one-by-one
    If Not My_LevelArt Is Nothing And Not My_LevelPalette Is Nothing Then
        'Apply the attached palette to the attached level art
        Call My_LevelArt.ApplyPalette(My_LevelPalette)
    End If
    
    If Not My_SpriteArt Is Nothing And Not My_SpritePalette Is Nothing Then
        Call My_SpriteArt.ApplyPalette(My_SpritePalette)
    End If
End Sub

'CountRings : Go through the level and count how many rings are available _
 ======================================================================================
Private Function CountRings() As Long
    'Are all our marbles present?
    If My_BlockMapping Is Nothing _
        Or My_FloorLayout Is Nothing _
        Or My_ObjectLayout Is Nothing _
    Then Exit Function
    
    'Count the rings in the floor layout:
    Dim X As Long, Y As Long
    For Y = 0 To Me.Height - 1
        For X = 0 To Me.Width - 1
            'Add the number of rings in this block
            Let CountRings = CountRings + _
                My_BlockMapping.RingsInBlock( _
                    My_FloorLayout.Block(X, Y) _
                )
        Next X
    Next Y
    
    'Count the rings in the object layout:
    Dim i As Long
    For i = 0 To 255
        'Is there a ring monitor object here?
        If My_ObjectLayout.Object(i).O = Monitor_Ring Then
            'TODO: This should be level dimensions, not floor layout dimensions!
            'Is it within the level dimensions? _
             (don't count ring monitors that lie outside of the floor layout because _
              the object layout is larger than the floor layout)
            If My_ObjectLayout.Object(i).X < Me.Width _
            And My_ObjectLayout.Object(i).Y < Me.Height Then
                Let CountRings = CountRings + 10
            End If
        End If
    Next
End Function
