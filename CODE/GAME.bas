Attribute VB_Name = "GAME"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: GAME

'This module represents the whole game the user is creating, including the levels, _
 palettes, objects, graphics and so forth

Public Levels() As S1Level

Public LevelPalettes As Collection
Public SpritePalettes As Collection
Public LevelArt As Collection
Public SpriteArt As Collection
Public FloorLayouts As Collection
Public BlockMappings As Collection
Public ObjectLayouts As Collection

Public UnderwaterLevelPalette As S1Palette
Public UnderwaterSpritePalette As S1Palette

'Since the plan is to eventually allow total editing of the graphics, _
 these are stored here and will be written to the project file
Public Sonic As S1Tileset
Public Ring As S1Tileset
Public HUD As S1Tileset
Public PowerUps As S1Tileset

Public EndSignPalette As S1Palette
Public EndSignTileset As S1Tileset
    
Public BossPalette As S1Palette
Public BossTileset1 As S1Tileset
Public BossTileset2 As S1Tileset
Public BossTileset3 As S1Tileset
Public CapsuleTileset As S1Tileset

'The list of objects in the game (i.e. for ObjectLayouts)
Public Objects() As S1SpriteLayout

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Clear : Free the project data from the app, clearing the memory in use _
 ======================================================================================
Public Sub Clear()
    Erase Objects
    
    Erase Levels
    Set LevelPalettes = Nothing:    Set LevelPalettes = New Collection
    Set SpritePalettes = Nothing:   Set SpritePalettes = New Collection
    Set LevelArt = Nothing:         Set LevelArt = New Collection
    Set SpriteArt = Nothing:        Set SpriteArt = New Collection
    Set FloorLayouts = Nothing:     Set FloorLayouts = New Collection
    Set BlockMappings = Nothing:    Set BlockMappings = New Collection
    Set ObjectLayouts = Nothing:    Set ObjectLayouts = New Collection
    
    Set UnderwaterLevelPalette = Nothing
    Set UnderwaterSpritePalette = Nothing
    
    Set Sonic = Nothing
    Set Ring = Nothing
    Set HUD = Nothing
    Set PowerUps = Nothing
    
    Set EndSignPalette = Nothing
    Set EndSignTileset = Nothing
    
    Set BossPalette = Nothing
    Set BossTileset1 = Nothing
    Set BossTileset2 = Nothing
    Set BossTileset3 = Nothing
End Sub
