Attribute VB_Name = "ROM"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: ROM

'This module can read a Sonic 1 ROM image and decode all the various information, _
 converting it into native objects for the level editor to use. It's not a class _
 because it modifies a whole bunch of external objects and whilst you could do this _
 a lot more modular, it's far more readable to use linear style code here

'/// PUBLIC VARS //////////////////////////////////////////////////////////////////////

'This will hold the path to an original Sonic 1 ROM to work from (see `Run.Main`) _
 When you run MaSS1VE for the first time it will ask you to drag-and-drop a ROM into _
 the form, where it will copy it into the user's app data so we can find it any time _
 we need it without too much fear it will get moved about
Public Path As String

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'The entire ROM binary image. It should be faster to read it in as a single binary _
 array and manipulate it then seek all over the file on disk
Private BIN As BinaryFile

'ROM Level Header Declarations: _
 --------------------------------------------------------------------------------------

'Where in the ROM the level pointers begin. the level headers begin afterwards, as _
 given by the destination of the first pointer
Private Const ROM_LEVEL_POINTERS = &H15580

'Floor Layout: _
 --------------------------------------------------------------------------------------
'The maximum number of bytes the level ("floor") layout data can occupy:
'This begins at $16DEA, and in an original ROM ends at $1FBA1 with free space until _
 $1FFFF, providing 1'119 bytes for expansion
Private Const ROM_FLOOR_SPACE = 37398

'Where in the ROM the level layout is relative to (this must be due to bank-switching)
Private Const ROM_FLOORDATA = &H14000
'The actual location where layout data begins
Private Const ROM_FLOORDATA_ABS = &H16DEA
'Pointers to layout data are relative from the beginning of the bank (#2DEA+)
Private Const ROM_FLOORDATA_REL = ROM_FLOORDATA_ABS - ROM_FLOORDATA

'Block Mappings: _
 --------------------------------------------------------------------------------------

'Blocks are the 4x4 tile patterns that the level is made of (tiles aren't set 1-by-1)

'Since at this moment we don't know where in the ROM the length of each block map is, _
 we can only determine the length of each block map by where the next one begins. _
 To that end, we need to keep a list of the "Mapping Location" of each level
Private BlockMappings As New Collection

'Where in the ROM the block mappings are found
Private Const ROM_BLOCKMAPPINGS = &H10000

'A relative pointer to the end of the block mappings in an original ROM
Private Const ROM_BLOCKMAPPINGS_END = &H4CA0&

'ROM Level Art / Sprite Art Declarations: _
 --------------------------------------------------------------------------------------

'The base address the level art pointers are relative to, that is, the level headers _
 contain a pointer that specifies that the level art can be found x number of bytes _
 from THIS address. In reality the level art data begins at $32FE6. _
 I imagine this has something to do with bank switching
Private Const ROM_LEVELART = &H30000
'As with level art, sprite art data is a pointer
Private Const ROM_SPRITEART = &H24000

'A relative pointer (from $30000) to the end of the level art in an original ROM
Private Const ROM_LEVELART_END = &HDA28&
'and where the sprite art ends (from $24000)
Private Const ROM_SPRITEART_END = &HB92E&      '$2F92E-$24000

'ROM Palette Declarations: _
 --------------------------------------------------------------------------------------

'Where in the ROM the palette pointers begin. these point to the 32-colour palettes _
 used for each tileset (Green Hill, Bridge, Jungle &c.)
Private Const ROM_PALETTE_POINTERS = &H627C

'Object Layout Declarations: _
 --------------------------------------------------------------------------------------

'The base address the object layout pointers are relative to
Private Const ROM_OBJECTLAYOUT = &H15580

'The different allowable types of objects
Public Enum OBJECT_TYPE
    NONE = &H0&
    'Power-up monitors / pickups
    Monitor_Ring = &H1&
    Monitor_Speed = &H2&
    Monitor_Life = &H3&
    Monitor_Shield = &H4&
    Monitor_Stars = &H5&        'Invincibility
    Monitor_Check = &H51&       'Checkpoint
    Monitor_Cont = &H52&        'Continue
    Emerald = &H6&
    BUBBLES = &H41&             'Air bubbles in Labyrinth
    'End of level stuff
    END_SIGN = &H7&
    BOSS_GREEN = &H12&          'Green Hill boss
    BOSS_BRIDGE = &H48&         'Bridge boss
    BOSS_JUNGLE = &H2C&         'Jungle boss
    BOSS_LABYRINTH = &H49&      'Labyrinth boss
    BOSS_SCRAP = &H22&          'Scrap Brain boss
    BOSS_SKY = &H4A&            'Sky Base boss
    CAPSULE = &H25&
    ANIMAL_RABBIT = &H23&       'Freed rabbit
    ANIMAL_BIRD = &H24&         'Freed bird
    ANIM_FINAL = &H53&          'Final animation
    ANIM_EMERALDS = &H54&       'Emerald animation when completing game
    'Badnicks
    BAD_CRAB = &H8&             'Crab Meat
    BAD_BUZZ = &HE&             'Buzz Bomber
    BAD_MOTO = &H10&            'Motobug
    BAD_NEWT = &H11&            'Newtron (the chameleon)
    BAD_HOG = &H1B&             'Ball Hog
    BAD_CAT = &H1F&             'Caterkiller
    BAD_CHOP = &H26&            'Chopper
    BAD_SPIKES = &H2D&          'Spiked crab "Yadrin"
    BAD_BOMB = &H32&            'Walking bomb
    BAD_ORB = &H35&             'Orbinaut / Unidus
    BAD_JAWS = &H3C&
    BAD_BURROB = &H44&          'Burrobot
    'Traps (non-animal enemies)
    TRAP_FLAME = &H16&          'Flame thrower in Scrap Brain
    TRAP_ELEC = &H1A&           'Electric sphere in Scrap Brain
    TRAP_PROP = &H31&           'Propeller
    TRAP_BIG_GUN = &H33&        'Large "Bullet-Bill" style gun turrets in Sky Base 2
    TRAP_GUN_ROT = &H37&        'Rotating gun turret
    TRAP_WALL = &H39&           'Spiked moving wall
    TRAP_GUN = &H3A&            'Fixed gun turret (Sky Base 1)
    TRAP_SPIKEBALL = &H3D&      'Rotating spike ball
    TRAP_SPEAR = &H3E&          'Up-Down spear
    TRAP_FIREBALL = &H3F&       'Fire-shooting head
    TRAP_ELEC_SKY = &H46&       'Electrical hazard in sky base boss?
    'Moving platforms
    PLAT_SWING = &H9&
    PLAT_WOOD = &HB&
    PLAT_FALL = &HC&            'Falls when touched
    PLAT_MOVE = &HF&            'Left-right moving platform
    PLAT_BUMP_MOVE = &H21&      'Moving bumper (Special Stage)
    PLAT_LOG_VERT = &H27&       'Vertical falling log
    PLAT_LOG_HORZ = &H28&       'Horizontal falling log
    PLAT_LOG_FLOAT = &H29&      'Floating log you can run on
    PLAT_FLY = &H38&            'Flying platform (Sky Base 2)
    PLAT_FLY_UPDWN = &H3B&      'Up-Down flying platform
    PLAT_UP = &H45&             'Platform that moves up when touched
    PLAT_FLIP = &H4C&           'Flipper (Special Stage)
    PLAT_BALANCE = &H4E&        'Weight balance (Bridge)
    'Doors / Switches
    DOOR_LEFT = &H17&           'Door that opens from the left only (Scrap Brain)
    DOOR_RIGHT = &H18&          'Door that opens from the right only
    DOOR_BOTH = &H19&           'Door that opens from both sides
    DOOR_SWITCH = &H1E&         'Switch-activated door
    SWITCH_BUTTON = &H1D&       'Push-switch
    'Effects and meta-objects
    META_CLOUDS = &H30&         'Passing clouds (Sky Base 2)
    META_WATER = &H40&          'Sets water level
    META_TRIP = &H4B&           'Makes Sonic fall (e.g. Green Hill 2)
    'Other
    UNKNOWN_1 = &H1C&           'Ball from the Ball Hog?
    FLOWER = &H50&              'Flower, Green Hill
    BLINK = &H55&               'Makes Sonic blink?
End Enum

'An object location. This is used by S1ObjectLayout and has to be defined here as you _
 cannot define a type as public in a class due to limitations in VB
Public Type OBJ
    X As Byte
    Y As Byte
    O As Byte                   'The object ID. We won't define this as `OBJECT_TYPE` _
                                 as that would take up 4 bytes, not 1
End Type

'Misc. Declerations: _
 --------------------------------------------------------------------------------------
Private Const ROM_UWPALETTE = &H24B             'The underwater palette

Private Const ROM_SONICART = &H20000            'Location of Sonic sprites
Private Const ROM_POWERUPSART = &H15180         'Location of power-up sprites
Private Const ROM_RINGART = &H2FD70             'Location of ring animation
Private Const ROM_HUDART = &H2F92E              'Location of HUD art

Private Const ROM_ENDSIGN_PALETTE = &H626C      'Location of palette for the end-sign
Private Const ROM_ENDSIGN_ART = &H28294         'Location of end-sign sprite

Private Const ROM_BOSS_PALETTE = &H731C         'Location of palette for boss sprites
Private Const ROM_BOSS_ART = &H2EEB1            'Location of boss sprites


'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Import : Read a Sonic 1 ROM and populate the level editor's data _
 ======================================================================================
Public Function Import() As Boolean
    Dim i As Long, ii As Long
    
    'Remove any existing project in memory
    Call GAME.Clear
    
    'TODO: Error if `ROM.Path` not set
    Debug.Print "Importing ROM: " & ROM.Path
    Dim StartTime As Single
    Let StartTime = Timer
    
    'Read the ROM into memory
    Set BIN = New BinaryFile
    Call BIN.Load(ROM.Path)
    
    'Misc: _
     ==================================================================================
    Debug.Print "* Sonic"
    Set GAME.Sonic = Import_Art_Uncompressed( _
        ROMOffset:=ROM_SONICART, NumberOfTiles:=512, _
        Skip4thByte:=True, UseTransparency:=True _
    )
    'The ring graphic is not a part of the level art. It gets painted on to the level _
     art in the last four tiles
    Debug.Print "* Power Ups"
    Set GAME.PowerUps = Import_Art_Uncompressed(ROM_POWERUPSART, 32)
    Debug.Print "* Ring Art"
    Set GAME.Ring = Import_Art_Uncompressed(ROM_RINGART, 20)
    Debug.Print "* HUD Art"
    Set GAME.HUD = Import_Art(ROM_HUDART, , True)
    Debug.Print "* Underwater Palettes"
    Set GAME.UnderwaterLevelPalette = Import_Palette(ROM_UWPALETTE)
    Set GAME.UnderwaterSpritePalette = Import_Palette(ROM_UWPALETTE + 16)
    Debug.Print "* End Sign"
    Set GAME.EndSignPalette = Import_Palette(ROM_ENDSIGN_PALETTE)
    Set GAME.EndSignTileset = Import_Art(ROM_ENDSIGN_ART)
    Call GAME.EndSignTileset.ApplyPalette(GAME.EndSignPalette)
    Debug.Print "* Boss"
    Set GAME.BossPalette = Import_Palette(ROM_BOSS_PALETTE)
    Set GAME.BossTileset = Import_Art(ROM_BOSS_ART)
    Call GAME.BossTileset.ApplyPalette(GAME.BossPalette)
    
    'Load Levels: _
     ==================================================================================
    'Where we are in the level list
    Dim LevelIndex As Byte
    'The address at which the level pointers end and the level data begins. _
     Determined by the first level pointer's destination (right after the pointers)
    Dim EndOfPointers As Long
    'For now, set it to the end of the ROM, we can't search farther than that!
    Let EndOfPointers = BIN.LOF
    
    'The length of each of the block mappings is -- to my knowledge -- not in the ROM. _
     Therefore we need to remember where each one begins and work out the lengths where _
     the beginnings and ends meet. We begin by including the end of the block mappings _
     section in the ROM so that we can know the length of the last one in the list
    Dim MappingLocations() As Long
    ReDim MappingLocations(0) As Long
    Let MappingLocations(0) = ROM_BLOCKMAPPINGS_END
    'Since we can't load the Block Mappings until the end, we also need to keep track _
     of which level uses which Block Mapping so that after loading them we can attach _
     them to the relevant level
    Dim LevelBlockMappings() As String
    
    'The current pointer (palettes / levels &c.) as we read it
    Dim Pointer As Long, Length As Long
    
    'LEVEL POINTERS: _
     ----------------------------------------------------------------------------------
    'A label / GoTo is used to account for VB's lack of `Continue` statement...
Continue_LevelPointers:
    'Stop reading level pointers when the level headers begin
    Do While ROM_LEVEL_POINTERS + (LevelIndex * 2) < EndOfPointers
        'Read in the LevelPointer from the ROM
        Let Pointer = BIN.IntLE(ROM_LEVEL_POINTERS + (LevelIndex * 2))
        'Convert the pointer to an absolute address and keep it for later, a lot of _
         stuff hinges on the data in the level header
        Dim LevelHeader As Long
        Let LevelHeader = ROM_LEVEL_POINTERS + Pointer
        'If this is the first level pointer, the destination denotes the end of the _
         level pointer table (the headers immediately follow the pointers)
        If LevelIndex = 0 Then Let EndOfPointers = LevelHeader
        
        'Expand the number of level entries
        If Lib.ArrayDimmed(GAME.Levels) = False Then _
           ReDim GAME.Levels(0) As S1Level Else _
           ReDim Preserve GAME.Levels(UBound(GAME.Levels) + 1) As S1Level
        
        'Skip null pointers (no level here) -- the first level must always be present!
        If Pointer = 0 Then Let LevelIndex = LevelIndex + 1: GoTo Continue_LevelPointers
        
        'Create our VB level object for the level editor
        Set GAME.Levels(LevelIndex) = New S1Level
        'Be aware of this; we'll be referring directly to the level object as we go
        With GAME.Levels(LevelIndex)
        
        'LEVEL HEADER: _
         ------------------------------------------------------------------------------
        'CRC the level header to try identify original data we can title
        Select Case BIN.CRC(LevelHeader, 37)
            Case &H9743C94D: Let .Title = "Green Hill Act 1"
            Case &H636DFA2: Let .Title = "Green Hill Act 2"
            Case &HC950F20F: Let .Title = "Green Hill Act 3"
            Case &H4D1FCE04: Let .Title = "Bridge Act 1"
            Case &HFEF13FFE: Let .Title = "Bridge Act 2"
            Case &H8B7743E8: Let .Title = "Bridge Act 3"
            Case &HAA5E6FA2: Let .Title = "Jungle Act 1"
            Case &HED46081F: Let .Title = "Jungle Act 2"
            Case &HF110344C: Let .Title = "Jungle Act 3"
            Case &HF2FD070F: Let .Title = "Labyrinth Act 1"
            Case &HA94FFEAF: Let .Title = "Labyrinth Act 2"
            Case &HB330D3DB: Let .Title = "Labyrinth Act 3"
            Case &H911CFAC0: Let .Title = "Scrap Brain Act 1"
            Case &H27F7F3A4: Let .Title = "Scrap Brain Act 2"
            Case &HAD8993AB: Let .Title = "Scrap Brain Act 3"
            Case &HC98E0AC1: Let .Title = "Sky Base Act 1"
            Case &HA1C1B189: Let .Title = "Sky Base Act 2"
            Case &HA83051EE: Let .Title = "Sky Base Act 3"
            Case &HABF3CA6: Let .Title = "End Sequence (Green Hill Act 1)"
            Case &H6F5D130D: Let .Title = "Scrap Brain Act 2 (Emerald Maze), from corridor"
            Case &H382208B2: Let .Title = "Scrap Brain Act 2 (Ballhog Area)"
            Case &HFA8C9A09: Let .Title = "Scrap Brain Act 2, from transporter"
            Case &H611B9564: Let .Title = "Scrap Brain Act 2 (Emerald Maze), from transporter"
            Case &HF9DE4E94: Let .Title = "Scrap Brain Act 2, from Emerald Maze"
            Case &HF8158BFE: Let .Title = "Scrap Brain Act 2, from Ballhog Area"
            Case &H2CCA058E: Let .Title = "Sky Base Act 2 (Interior)"
            Case &H225C4216: Let .Title = "Special Stage 1"
            Case &HAE04857E: Let .Title = "Special Stage 2"
            Case &H546AD778: Let .Title = "Special Stage 3"
            Case &H65FA76F5: Let .Title = "Special Stage 4"
            Case &HADE4A707: Let .Title = "Special Stage 5"
            Case &H4E7FAF49: Let .Title = "Special Stage 6"
            Case &H85F9E7D2: Let .Title = "Special Stage 7"
            Case &H1A9528B9: Let .Title = "Special Stage 8"
            Case Else: Let .Title = "Level #" & LevelIndex + 1
        End Select
        Debug.Print .Title & " Header: $" & Hex(LevelHeader) & " (#" & Hex(Pointer) & ")"
        
        'Record the "Mapping Location" (pointer to the level's Block Mappings)
        Call Lib.PushLong( _
            What:=BIN.IntLE(LevelHeader + 19), Where:=MappingLocations, _
            AllowDuplicates:=False _
        )
        'Record which Block Mapping this level uses so it can be set after the _
         Block Mappings have been loaded
        ReDim Preserve LevelBlockMappings(LevelIndex) As String
        Let LevelBlockMappings(LevelIndex) = Hex(BIN.IntLE(LevelHeader + 19))
        
        'Sonic's starting location
        Let .StartX = BIN.B(LevelHeader + 13)
        Let .StartY = BIN.B(LevelHeader + 14)
        
        'Whether the level has a water line (i.e. Labyrinth)
        Let .IsUnderWater = (BIN.B(LevelHeader + 33) = &H80&)
        
        'Copy across the unknown and unimplemented stuff from the ROM header, it's _
         needed when it comes to exporting
        Let .ROM_SP = BIN.B(LevelHeader)          'Solidity pointer -- UNDOCUMENTED
        Let .ROM_X1 = BIN.B(LevelHeader + 5)      'Unknown byte 1
        Let .ROM_X2 = BIN.B(LevelHeader + 6)      'Unknown byte 2
        Let .ROM_LW = BIN.B(LevelHeader + 7)      '"Level Width" (Unknown)
        Let .ROM_LH = BIN.B(LevelHeader + 8)      '"Level Height" (Unknown)
        Let .ROM_X3 = BIN.B(LevelHeader + 9)      'Unknown byte 3
        Let .ROM_X4 = BIN.B(LevelHeader + 10)     'Unknown byte 4
        Let .ROM_X5 = BIN.B(LevelHeader + 11)     'Unknown byte 5
        Let .ROM_X6 = BIN.B(LevelHeader + 12)     'Unknown byte 6
        Let .ROM_ML = BIN.IntLE(LevelHeader + 19) 'Pointer from $10000 to the Block Mappings ($10000)
        Let .ROM_LA = BIN.IntLE(LevelHeader + 21) 'Pointer from $30000 to the Level Art ($32FE6)
        Let .ROM_SA = BIN.IntLE(LevelHeader + 24) 'Pointer from $24000 to the Sprite Art ($2A12A)
        Let .ROM_IP = BIN.B(LevelHeader + 26)     'Initial palette index
        Let .ROM_CS = BIN.B(LevelHeader + 27)     'Palette cycle speed
        Let .ROM_CC = BIN.B(LevelHeader + 28)     'Number of palette colour cycles
        Let .ROM_CP = BIN.B(LevelHeader + 29)     'Cycle Palette Index
        Let .ROM_OL = BIN.IntLE(LevelHeader + 30) 'Pointer from $15580 to the Object Layout ($15AB4)
        Let .ROM_SR = BIN.B(LevelHeader + 32)     'Scrolling and Ring HUD flags
        Let .ROM_TL = BIN.B(LevelHeader + 34)     'Time and Lightning flags
        Let .ROM_MU = BIN.B(LevelHeader + 36)     'Music
        
        'FLOOR LAYOUT: _
         ------------------------------------------------------------------------------
        'The location of the Floor Layout in the ROM, this is a pointer from $14000 _
         to the Floor Layout ($16DEA). We'll also use this as an ID
        Let Pointer = BIN.IntLE(LevelHeader + 15)
        
        'Has this floor layout been seen before? (Some levels share the same layout, _
         such as the special stages and parts of Scrap Brain)
        If Not Lib.Exists(Key:=Hex(Pointer), Col:=GAME.FloorLayouts) Then
            'Create a new Floor Layout object for the editor
            Dim FloorLayout As S1FloorLayout
            Set FloorLayout = New S1FloorLayout
            Let FloorLayout.ID = Hex(Pointer)
            
            'Set the size of the level
            Call FloorLayout.Resize( _
                NewWidth:=BIN.IntLE(LevelHeader + 1), _
                NewHeight:=BIN.IntLE(LevelHeader + 3) _
            )
            
            'CRC the floor data and title original floor layouts accordingly
            Select Case BIN.CRC(ROM_FLOORDATA + Pointer, BIN.IntLE(LevelHeader + 17))
                Case &HE9391806: Let FloorLayout.Title = "Green Hill Act 1 / End Sequence"
                Case &H9545C866: Let FloorLayout.Title = "Green Hill Act 2"
                Case &HBCEA9E8E: Let FloorLayout.Title = "Green Hill Act 3"
                Case &H5DF36C9B: Let FloorLayout.Title = "Bridge Act 1"
                Case &H9B27E806: Let FloorLayout.Title = "Bridge Act 2"
                Case &HF67903F: Let FloorLayout.Title = "Bridge Act 3"
                Case &HDEA8051E: Let FloorLayout.Title = "Jungle Act 1"
                Case &HB773643C: Let FloorLayout.Title = "Jungle Act 2 / Special Stage 4 & 8"
                Case &H9E27F737: Let FloorLayout.Title = "Jungle Act 3"
                Case &H23A4D189: Let FloorLayout.Title = "Labyrinth Act 1"
                Case &H2F15BB84: Let FloorLayout.Title = "Labyrinth Act 2"
                Case &H6592A7F5: Let FloorLayout.Title = "Labyrinth Act 3"
                Case &HAC504F28: Let FloorLayout.Title = "Scrap Brain Act 1"
                Case &H6D9F0F2C: Let FloorLayout.Title = "Scrap Brain Act 2"
                Case &HF6907409: Let FloorLayout.Title = "Scrap Brain Act 3"
                Case &H2FF1B76D: Let FloorLayout.Title = "Sky Base Act 1"
                Case &HFE8FA4A6: Let FloorLayout.Title = "Sky Base Act 2"
                Case &H8B656D1A: Let FloorLayout.Title = "Sky Base Act 2 & 3 Interiors"
                Case &H3E5D5F18: Let FloorLayout.Title = "Scrap Brain Act 2 (Emerald Maze)"
                Case &HA6627B27: Let FloorLayout.Title = "Scrap Brain Act 2 (Ballhog Area)"
                Case &HDCF86937: Let FloorLayout.Title = "Special Stage 1 / 2 / 3 / 5 / 6 / 7"
                Case Else: Let FloorLayout.Title = "Floor Layout (" & FloorLayout.ID & ")"
            End Select
            
            'Fetch the Floor Layout data from the ROM (the size of the compressed data _
             is specified in the level header) and decompress it
            Dim Data() As Byte
            Let Data = DecompressRLEData( _
                Offset:=ROM_FLOORDATA + Pointer, _
                Length:=BIN.IntLE(LevelHeader + 17) _
            )
            
            'All level sizes are such that they add up to 4K of decompressed data. _
             Even though the size of "Scrap Brain 2 Ballhog Area" in the level headers _
             is 4K, there is only 2K of data when the floor layout is decompressed!
            If UBound(Data) < 4095 Then
                Debug.Print "! Fixing incorrect level size"
                'Expand it to 4K
                ReDim Preserve Data(4095) As Byte
            End If
            
            Call FloorLayout.SetByteStream(Data)
            Erase Data
            
            'Add the Floor Layout to the global pool of layouts available
            Call GAME.FloorLayouts.Add(Item:=FloorLayout, Key:=FloorLayout.ID)
        End If
        Debug.Print "- Floor Layout: $" & Hex(ROM_FLOORDATA + Pointer) & " (#" & Hex(Pointer) & ") '" & FloorLayout.Title & "'"
        
        'Apply the Floor Layout to the level
        Set .FloorLayout = GAME.FloorLayouts(Hex(Pointer))
        
        'OBJECT LAYOUT: _
         ------------------------------------------------------------------------------
        'Refer to the Object Layout pointed to by the Level Header
        Let Pointer = BIN.IntLE(LevelHeader + 30)
        
        'Has this Object Layout been seen before?
        If Not Lib.Exists(Key:=Hex(Pointer), Col:=GAME.ObjectLayouts) Then
            Debug.Print "- Object Layout: $" & Hex(ROM_OBJECTLAYOUT + Pointer) & " (#" & Hex(Pointer) & ")"
            
            'Create a new Object Layout to populate from the ROM
            Dim ObjectLayout As S1ObjectLayout
            Set ObjectLayout = New S1ObjectLayout
            ObjectLayout.ID = Hex(Pointer)
            
            'The first byte of the Object Layout is the number of objects
            For i = 0 To BIN.B(ROM_OBJECTLAYOUT + Pointer) - 1
                'NOTE: There are null objects (&H00 / &HFF), often occuring a few _
                 times within the same location, but these are not required and _
                 appear to be some left over from the way the original developer's _
                 editor worked. The `Add` method will ignore these automatically
                Call ObjectLayout.Add( _
                    ObjID:=BIN.B(ROM_OBJECTLAYOUT + Pointer + 1 + (3 * i)), _
                    X:=BIN.B(ROM_OBJECTLAYOUT + Pointer + 1 + (3 * i) + 1), _
                    Y:=BIN.B(ROM_OBJECTLAYOUT + Pointer + 1 + (3 * i) + 2) _
                )
            Next
            'Add the Object Layout to the global stock
            Call GAME.ObjectLayouts.Add(Item:=ObjectLayout, Key:=ObjectLayout.ID)
        End If
        'Apply the Object Layout to this level
        Set .ObjectLayout = GAME.ObjectLayouts(Hex(Pointer))
        
        'PALETTE: _
         ------------------------------------------------------------------------------
        'Whilst we could read the palettes independently of the levels, the level _
         header contains the number of colour cycles, which makes our job much easier.
        Dim Palette As S1Palette
        Dim InitialPalette As Byte
        Let InitialPalette = BIN.B(LevelHeader + 26)
        
        'Has this palette already been processed?
        If Not Lib.Exists(Key:=CStr(InitialPalette), Col:=GAME.LevelPalettes) Then
            'The "Initial Palette" value in the Level Header tells us which of the _
             palette pointers to use to look up the actual palette :S
            'Note that the palette pointers are absolute rather than relative
            Let Pointer = BIN.IntLE(ROM_PALETTE_POINTERS + (InitialPalette * 2))
            
            'TILE PALETTE:
            'The first half of the palette (16 colours) is used for the background
            Set Palette = Import_Palette(Offset:=Pointer)
            Let Palette.ID = CStr(InitialPalette)
            'Add to the global stock
            Call GAME.LevelPalettes.Add(Item:=Palette, Key:=Palette.ID)
            
            'SPRITE PALETTE:
            'Do the same thing, but for the sprite palette
            Set Palette = Import_Palette(Offset:=Pointer + 16)
            Let Palette.ID = CStr(InitialPalette)
            'Add to the global stock
            Call GAME.SpritePalettes.Add(Item:=Palette, Key:=Palette.ID)
        End If
        'Apply the palettes to the level object we are building
        Set .LevelPalette = GAME.LevelPalettes(CStr(InitialPalette))
        Set .SpritePalette = GAME.SpritePalettes(CStr(InitialPalette))
        
        'LEVEL ART: _
         ------------------------------------------------------------------------------
        Let Pointer = BIN.IntLE(LevelHeader + 21)
        'Has this Level Art been processed yet?
        If Not Lib.Exists(Key:=Hex(Pointer), Col:=GAME.LevelArt) Then
            Debug.Print "- Level Art: $" & Hex(ROM_LEVELART + Pointer) & " (#" & Hex(Pointer) & ")"
            
            Dim LevelArt As S1Tileset
            Set LevelArt = Import_Art(ROM_LEVELART + Pointer, .LevelPalette)
            Let LevelArt.ID = Hex(Pointer)
            
            'Add this Level Art to the global stock
            Call GAME.LevelArt.Add(Item:=LevelArt, Key:=LevelArt.ID)
        End If
        'Apply the Level Art to the level
        Set .LevelArt = GAME.LevelArt(Hex(Pointer))
        
        'SPRITE ART: _
         ------------------------------------------------------------------------------
        Let Pointer = BIN.IntLE(LevelHeader + 24)
        'Has this Level Art been processed yet?
        If Not Lib.Exists(Key:=Hex(Pointer), Col:=GAME.SpriteArt) Then
            Debug.Print "- Sprite Art: $" & Hex(ROM_SPRITEART + Pointer) & " (#" & Hex(Pointer) & ")"
            
            Dim SpriteArt As S1Tileset
            Set SpriteArt = Import_Art(ROM_SPRITEART + Pointer, .SpritePalette, True)
            Let SpriteArt.ID = Hex(Pointer)
            
            'Add this Level Art to the global stock
            Call GAME.SpriteArt.Add(Item:=SpriteArt, Key:=SpriteArt.ID)
        End If
        'Apply the Level Art to the level
        Set .SpriteArt = GAME.SpriteArt(Hex(Pointer))
        
        'Next level -- onwards and upwards!
        End With
        Let LevelIndex = LevelIndex + 1
    Loop
    
    'BLOCK MAPPINGS: _
     ==================================================================================
    'Floor Layouts are made up of indicies to the block mappings - 4x4 tile blocks, _
     selected out of the tile set. We will need to import these after the level _
     headers have been read as we do not know the length of each block mapping until _
     we know where one ends and another begins
    'Begin by sorting the list of starting addresses we took from the level headers
    Call Lib.CombSort(MappingLocations)
    'Loop through it, working out the lengths in between _
     (the last element is the end of the block mappings ROM space)
    For i = LBound(MappingLocations) To UBound(MappingLocations) - 1
        'Create our VB object for representing the block mappings
        Dim BlockMapping As S1BlockMapping
        Set BlockMapping = New S1BlockMapping
        Let BlockMapping.ID = Hex(MappingLocations(i))
        
        'The length of the block mappings is based on where one level's block _
         mappings end and the next begin
        Let Length = MappingLocations(i + 1) - MappingLocations(i)
        'Copy the block mappings over into the VB object
        Let BlockMapping.Length = Length \ 16
        For ii = 0 To Length
            BlockMapping.Tile(BlockIndex:=ii \ 16, TileIndex:=ii Mod 16) = _
                BIN.B(ROM_BLOCKMAPPINGS + MappingLocations(i) + ii)
        Next ii
        
        'Add to the global stock
        Call GAME.BlockMappings.Add(Item:=BlockMapping, Key:=BlockMapping.ID)
        Set BlockMapping = Nothing
    Next
    
    'Apply Block Mappings / Art to Level: _
     ==================================================================================
    For LevelIndex = LBound(GAME.Levels) To UBound(GAME.Levels)
    If Not GAME.Levels(LevelIndex) Is Nothing Then
    With GAME.Levels(LevelIndex)
        Set .BlockMapping = GAME.BlockMappings(LevelBlockMappings(LevelIndex))
        
        Call GAME.Ring.ApplyPalette(.LevelPalette)
        Call GAME.Ring.PaintTile(.LevelArt.Tiles.hDC, 252 * 8, 0, 16)
        Call GAME.Ring.PaintTile(.LevelArt.Tiles.hDC, 253 * 8, 0, 17)
        Call GAME.Ring.PaintTile(.LevelArt.Tiles.hDC, 254 * 8, 0, 18)
        Call GAME.Ring.PaintTile(.LevelArt.Tiles.hDC, 255 * 8, 0, 19)
    End With: End If
    Next LevelIndex
    
    'Post Processing: _
     ==================================================================================
    'There are some oddities and wasted space in the original ROM that we can clean up:
    
    'One of the biggest problems (as far as our editor is concerned) is that Special _
     stage 4 & 8 and Jungle Act 2 are on the same floor layout. This means that from _
     each level the other looks like garbage and the ring count is way off
    'Detect if the original Special Stage 4/8 and Jungle Act 2 levels exist:
    If GAME.Levels(7).FloorLayout.Title = "Jungle Act 2 / Special Stage 4 & 8" Then
        Debug.Print "Fixing Jungle Act 2 / Special Stage 4 & 8"
        
        Dim FloorLayout1 As S1FloorLayout
        Dim FloorLayout2 As New S1FloorLayout
        Set FloorLayout1 = GAME.FloorLayouts(GAME.Levels(7).FloorLayout.ID)
    
        'We need to split the combined layout into two separate floor layouts
        Let FloorLayout1.Title = "Jungle Act 2"
        Let FloorLayout2.Title = "Special Stage 4 / 8"
        Let FloorLayout2.ID = "Jungle Act 2 / Special Stage Fix"
        'Make a duplicate copy of the floor layout data
        Call FloorLayout2.Resize(FloorLayout1.Width, FloorLayout1.Height)
        Call FloorLayout2.SetByteStream(FloorLayout1.GetByteStream())
        
        'On Jungle Act 2, erase the Special Stage 4/8 data
        Dim X As Long, Y As Long
        For Y = 0 To 139: For X = 0 To FloorLayout1.Width - 1
            Let FloorLayout1.Block(X, Y) = 0
        Next X: Next Y
        
        'On Special Stage 4 / 8, erase the Jungle Act 2 data
        For Y = 140 To FloorLayout2.Height - 1: For X = 0 To FloorLayout2.Width - 1
            Let FloorLayout2.Block(X, Y) = 0
        Next X: Next Y
        
        'Add the new layout to the global pool
        Call GAME.FloorLayouts.Add(FloorLayout2, FloorLayout2.ID)
        'Apply the new layout to Special Stage 4 & 8
        Set GAME.Levels(31).FloorLayout = GAME.FloorLayouts(FloorLayout2.ID)
        Set GAME.Levels(35).FloorLayout = GAME.FloorLayouts(FloorLayout2.ID)
        'And clean up
        Set FloorLayout1 = Nothing
        Set FloorLayout2 = Nothing
    End If
    
    'ROM imported!
    Let Import = True
    Debug.Print "ROM was imported: " & Round(Timer - StartTime, 3)
    
    'Free the ROM from memory
    Set BIN = Nothing
End Function

'Export : Write the level to a playable ROM _
 ======================================================================================
Public Sub Export(ByVal FilePath As String, Optional ByVal StartingLevel As Byte = 255)
    Dim i As Long

    'Let's do this thing!
    Debug.Print "Exporting ROM: " & FilePath
    Debug.Print "(Using original ROM: " & ROM.Path & ")"
    
    'Read the ROM into memory
    Set BIN = New BinaryFile
    Call BIN.Load(ROM.Path)

    'Compress levels: _
     ==================================================================================
    'We need to check the size of all levels combined after compression, it cannot be _
     more than 36.5KB -- the maximum space in the ROM for levels ($16DEA-$1FFFF)

    'Track which floor layouts have already been processed
    Dim FloorLocations As New Collection
    'And the sizes of the compressed data
    Dim FloorSizes As New Collection
    'This tracks where the last floor layout ends and the next should begin
    Dim CurrentFloorLocation As Long
    
    'Loop through the levels
    Dim LevelIndex As Byte
    For LevelIndex = LBound(GAME.Levels) To UBound(GAME.Levels)
        'If this level is not valid, skip it (there's two blank levels in the original _
         ROM that act as spaces between the main set of levels and the Scrap Brain _
         inter-linked areas / special stages
        If Not GAME.Levels(LevelIndex) Is Nothing Then
            'Different levels in the ROM may share the same layout, for example: _
             the interlinked parts of Scrap Brain Act 2, so only identify unique _
             floor layouts that we have to compress
            If Not Lib.Exists(Key:=GAME.Levels(LevelIndex).FloorLayout.ID, Col:=FloorLocations) Then
                'We need to record the *new* ROM address of this floor layout and _
                 associate it with the old location so that we can quickly remap _
                 the ROM addresses in the level headers
                Call FloorLocations.Add( _
                    Item:=ROM_FLOORDATA_REL + CurrentFloorLocation, _
                    Key:=GAME.Levels(LevelIndex).FloorLayout.ID _
                )
                'Compress the floor layout data
                Dim Data() As Byte
                Let Data = CompressRLEData( _
                    GAME.Levels(LevelIndex).FloorLayout.GetByteStream _
                )
                'Record the compressed size to use in the level header
                Call FloorSizes.Add( _
                    Item:=UBound(Data), _
                    Key:=GAME.Levels(LevelIndex).FloorLayout.ID _
                )
                'Write the compressed data to the ROM
                Call BIN.SetBArr( _
                    Index:=ROM_FLOORDATA_ABS + CurrentFloorLocation, _
                    Arr:=Data _
                )
                'The next floor layout will begin where the last eneded
                Let CurrentFloorLocation = CurrentFloorLocation + UBound(Data) + 1
            End If
        End If
    Next LevelIndex
    Debug.Print "Total Compressed Floor Layout Size: " & CurrentFloorLocation & " Bytes"
    If CurrentFloorLocation > ROM_FLOOR_SPACE Then Stop
    
    'Write Level Headers: _
     ==================================================================================
    For LevelIndex = LBound(GAME.Levels) To UBound(GAME.Levels)
        If Not GAME.Levels(LevelIndex) Is Nothing Then
            'Locate the level header from the pointers table
            Dim Pointer As Long
            Let Pointer = ROM_LEVEL_POINTERS + _
                          BIN.IntLE(ROM_LEVEL_POINTERS + LevelIndex * 2)
            
            'Write the level header:
            With GAME.Levels(LevelIndex)
                'Graphics:
                Let BIN.IntLE(Pointer + 21) = .ROM_LA   'Level Art pointer
                Let BIN.IntLE(Pointer + 24) = .ROM_SA   'Sprite Art pointer
                Let BIN.B(Pointer + 26) = .ROM_IP       'Initial palette index
                Let BIN.B(Pointer + 29) = .ROM_CP       'Cycle palette index
                Let BIN.B(Pointer + 28) = .ROM_CC       'Number of palette cycles
                Let BIN.B(Pointer + 27) = .ROM_CS       'Cycle palette speed
                
                'Level dimensions:
                Let BIN.IntLE(Pointer + 1) = .Width     'Actual level width
                Let BIN.IntLE(Pointer + 3) = .Height    'Actual level height
                Let BIN.B(Pointer + 7) = .ROM_LW        'Level Width -- UNDOCUMENTED
                Let BIN.B(Pointer + 8) = .ROM_LH        'Level Height -- UNDOCUMENTED
                
                'Level attributes
                Let BIN.B(Pointer + 36) = .ROM_MU       'Music index
                Let BIN.B(Pointer + 13) = .StartX       'Sonic starting position
                Let BIN.B(Pointer + 14) = .StartY
                Let BIN.B(Pointer + 32) = .ROM_SR       'Scrolling and ring HUD flags
                Let BIN.B(Pointer + 34) = .ROM_TL       'Time and lightning flags
                'Underwater flag
                Let BIN.B(Pointer + 33) = IIf(.IsUnderWater, &H80&, &H0&)
                
                'Level structure
                Let BIN.B(Pointer) = .ROM_SP            'Solidity pointer -- UNDOCUMENTED
                Let BIN.IntLE(Pointer + 19) = .ROM_ML   'Block Mappings pointer
                Let BIN.IntLE(Pointer + 15) = _
                    FloorLocations(.FloorLayout.ID)     'The Floor Layout pointer
                Let BIN.IntLE(Pointer + 17) = _
                    FloorSizes(.FloorLayout.ID)         'Size of the Floor Layout data
                Let BIN.IntLE(Pointer + 30) = .ROM_OL   'Object Layout pointer
                
                'Unknown / unimplemented bytes:
                Let BIN.B(Pointer + 35) = 0             'Always "0"
                Let BIN.B(Pointer + 5) = .ROM_X1
                Let BIN.B(Pointer + 6) = .ROM_X2
                Let BIN.B(Pointer + 9) = .ROM_X3
                Let BIN.B(Pointer + 10) = .ROM_X4
                Let BIN.B(Pointer + 11) = .ROM_X5
                Let BIN.B(Pointer + 12) = .ROM_X6
                Let BIN.B(Pointer + 23) = &H9&          'Always "9"
            End With
        End If
    Next LevelIndex
    
    'Customisations: _
     ==================================================================================
    'If a level skip is given (playtesting for example), write it in _
     (replaces the first level with the desired level)
    'TODO: This does not set up the underwater effect on Labyrinth Act 3; there must
    '      be somewhere else that controls the starting level in the ROM?
    If StartingLevel < 255 Then
        BIN.IntLE(ROM_LEVEL_POINTERS) = &H4A + (StartingLevel * 37)
    End If
    
    'Save the ROM to the new location
    Call BIN.Save(FilePath)
    Debug.Print "ROM was exported" & vbCrLf
    
    'Free the ROM from memory
    Set BIN = Nothing
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Import_Art_Uncompressed _
 ======================================================================================
Private Function Import_Art_Uncompressed( _
    ByRef ROMOffset As Long, _
    ByVal NumberOfTiles As Long, _
    Optional ByRef Palette As S1Palette = Nothing, _
    Optional ByVal UseTransparency As Boolean = False, _
    Optional ByVal Skip4thByte As Boolean = False _
) As S1Tileset
    'Create a new tileset in the return field
    Set Import_Art_Uncompressed = New S1Tileset
    Call Import_Art_Uncompressed.Create( _
        NumberOfTiles:=NumberOfTiles, Palette:=Palette, _
        UseTransparency:=UseTransparency _
    )
    'Fetch the image data as a byte array; easier/faster to manipulate
    Dim ImageData() As Byte
    Let ImageData = Import_Art_Uncompressed.Tiles.GetByteStream
    
    'Where we are in the uncompressed data as we move through it
    Dim ReadHead As Long
    
    'Start processing each tile...
    Dim TileIndex As Long, Row As Byte
    For TileIndex = 0 To NumberOfTiles - 1: For Row = 0 To 7
        'The Sonic player sprites do not include the 4th byte (as a form of very _
         simple compression) as no more than 8 colours are used and the 4th byte _
         is always zero
        Dim Fourth As Byte
        If Skip4thByte = False _
            Then Let Fourth = BIN.B(ROMOffset + ReadHead + 3) _
            Else Let Fourth = 0
        
        'Get one row of pixels from the data
        Dim Pixels() As Byte
        Let Pixels = DecodeSMSTileRow( _
            Byte1:=BIN.B(ROMOffset + ReadHead), _
            Byte2:=BIN.B(ROMOffset + ReadHead + 1), _
            Byte3:=BIN.B(ROMOffset + ReadHead + 2), _
            Byte4:=Fourth _
        )
        If Skip4thByte = False _
            Then Let ReadHead = ReadHead + 4 _
            Else Let ReadHead = ReadHead + 3
        
        'Paint the pixels on to the tileset
        Dim Pixel As Long
        For Pixel = 0 To 7
            Let ImageData( _
              (Row * NumberOfTiles * 8) + _
              (TileIndex * 8) + Pixel _
            ) = Pixels(Pixel)
        Next Pixel
    Next Row: Next TileIndex
    
    Call Import_Art_Uncompressed.Tiles.SetByteStream(ImageData)
    Erase ImageData
End Function

'Import_Art: Decompresses ROM art data into an image _
 ======================================================================================
Private Function Import_Art( _
    ByVal ROMOffset As Long, _
    Optional ByRef Palette As S1Palette = Nothing, _
    Optional ByVal UseTransparency As Boolean = False _
) As S1Tileset
    'WARNING: Here comes the science bit! _
     If you want to understand how this works then read VERY carefully. First, some _
     background info: The Master System uses character-based graphics, that is, 8x8 _
     pixel tiles arranged into a 32x24 grid. The screen is not drawn to directly, _
     instead up to 448 tiles can be defined and images are crafted by arranging _
     different tiles on the grid.
    'The graphics for a level consist of 256 tiles. Each tile consists of 8 rows _
     (of 8 pixels), each row is 4 bytes (more on that later). The level art is _
     compressed in a very complex, but effective, manner. First, the rows of all the _
     tiles are divided into unique rows and duplicate rows.
    
    'The data is stored like this:
    '* Header (8-bytes) _
     * List of unique rows _
     * List of duplicate rows _
     * Art Data (the actual graphics)
        
    'The list of unique rows is 256 bytes long (one byte for each tile we are defining) _
     where the bits tell us which rows in a tile are the unique ones to draw first.
    'The list of duplicate rows fills in the rows missed by the unique list. For each _
     row in the tile that was missed by the unique rows list the duplicate rows list _
     will contain an index value to where in the art data the row is
    
    'Where in the data stream we'll find:
    Dim UniqueRowsLocation As Long, _
        DuplicateRowsLocation As Long, _
        ArtDataLocation As Long
    'The number of tiles in the art data; 128 for sprites, 256 for tiles
    Dim TileCount As Long
    
    'Each art block begins with an 8-byte header, the first two bytes are always _
     "48 59". Bytes 3 & 4 are an offset from the beginning of the header to the _
     duplicate rows list, skipping over the unique rows list. From this, we can tell _
     the length the unique rows list which is one byte per tile, less 8 bytes for _
     the header
    If BIN.IntBE(ROMOffset) <> &H4859 Then Stop
    Let UniqueRowsLocation = 8
    Let DuplicateRowsLocation = BIN.IntLE(ROMOffset + 2)
    Let ArtDataLocation = BIN.IntLE(ROMOffset + 4)
    Let TileCount = DuplicateRowsLocation - UniqueRowsLocation
    
    'Create a new level art object in the return
    Set Import_Art = New S1Tileset
    Call Import_Art.Create(TileCount, Palette, UseTransparency)
    
    'Fetch the image as a byte stream; this is a thousand times faster to manipulate _
     than setting pixels one by one
    Dim ImageData() As Byte
    Let ImageData = Import_Art.Tiles.GetByteStream()
    
    'We conditionally move through the art data as we process the unique rows list, _
     this will track the current "read head" on the art data section
    Dim ArtDataPointer As Long
    Let ArtDataPointer = ArtDataLocation
    
    'We'll walk through the duplicate row data manually _
     (some values are one byte, some are two)
    Dim DuplicateRowsPointer As Long
    Let DuplicateRowsPointer = DuplicateRowsLocation
    
    'Somewhere to store the decoded pixels
    Dim Pixels() As Byte
    
    'Each byte in the list corresponds to one tile in the tile set we're drawing
    Dim TileIndex As Long
    For TileIndex = 0 To TileCount - 1
        'One bit in the list byte refers to one row in the graphic tile
        Dim Bit As Long: For Bit = 0 To 7
            'Unique Row: _
             --------------------------------------------------------------------------
            'Each byte in the list is a bit-pattern determining which of the 8 _
             rows in a tile are unique and which are duplicated elsewhere
            'If this bit is 0, this is a unique row: (the brackets are totally _
             required here otherwise there's a boolean / bitwise confusion!)
            If (Not BIN.B(ROMOffset + UniqueRowsLocation + TileIndex) And 2 ^ Bit) > 0 Then
                'Decode 4 bytes from the art data into 8 pixels of colour
                Let Pixels = DecodeSMSTileRow( _
                    Byte1:=BIN.B(ROMOffset + ArtDataPointer), _
                    Byte2:=BIN.B(ROMOffset + ArtDataPointer + 1), _
                    Byte3:=BIN.B(ROMOffset + ArtDataPointer + 2), _
                    Byte4:=BIN.B(ROMOffset + ArtDataPointer + 3) _
                )
                Let ArtDataPointer = ArtDataPointer + 4
            
            'Duplicate Row: _
             --------------------------------------------------------------------------
            Else
                'Read in the index to which art data row to paint. _
                 indexes can be more than 256 so two bytes are sometimes used
                Dim Index As Long
                Let Index = CLng(BIN.B(ROMOffset + DuplicateRowsPointer))
                Let DuplicateRowsPointer = DuplicateRowsPointer + 1
                'If the index is > &HEF then it's two bytes long
                If Index > &HEF Then
                    'The two-byte index is in the format of &HFxxx where _
                     "xxx" is the actual value and the F is the two-byte marker
                    Let Index = BIN.IntBE(ROMOffset + DuplicateRowsPointer - 1) And &HFFF
                    Let DuplicateRowsPointer = DuplicateRowsPointer + 1
                End If
                'Decode 4 bytes from the art data into 8 pixels of colour
                Let Pixels = DecodeSMSTileRow( _
                    Byte1:=BIN.B(ROMOffset + ArtDataLocation + (Index * 4)), _
                    Byte2:=BIN.B(ROMOffset + ArtDataLocation + (Index * 4) + 1), _
                    Byte3:=BIN.B(ROMOffset + ArtDataLocation + (Index * 4) + 2), _
                    Byte4:=BIN.B(ROMOffset + ArtDataLocation + (Index * 4) + 3) _
                )
            End If
            'Draw the pixels to the tile set
            Dim Pixel As Long
            For Pixel = 0 To 7
                Let ImageData( _
                    (Bit * TileCount * 8) + _
                    (TileIndex * 8) + Pixel _
                ) = Pixels(Pixel)
            Next Pixel
        Next Bit
    Next TileIndex
    
    Call Import_Art.Tiles.SetByteStream(ImageData)
    Erase ImageData
End Function

'Import_Palette _
 ======================================================================================
Private Function Import_Palette(ByVal Offset) As S1Palette
    'Create a new palette in the return field
    Set Import_Palette = New S1Palette
    
    Dim i As Long
    For i = 0 To 15
        Let Import_Palette.Colour(i) = DecodeSMSColour(BIN.B(Offset + i))
    Next i
End Function

'DecompressRLEData : Decompress Run-Length-Encoded ROM data (e.g. floor layouts) _
 ======================================================================================
Private Function DecompressRLEData(ByVal Offset, ByVal Length As Long) As Byte()
    Dim i As Long, ii As Long, Repeat As Long
    Dim Output() As Byte
    
    'Whilst the repeat-length is a byte, we need a null value and an overflow _
     (due to a possible bug/oversight of the original developers)
    Dim Previous As Integer
    Let Previous = -1
    
    'step through the compressed data stream...
    For i = Offset To Offset + Length
        'If this byte and the previous are the same, the next byte determines the _
         number of duplicate bytes to write (less 1, it includes the current byte)
        If BIN.B(i) = Previous Then
            'Refer to the next byte
            Let Repeat = BIN.B(i + 1)
            'A byte-rollover occurs with lengths of 256, so a length of 0 is 256
            If Repeat = 0 Then Let Repeat = 256
            'Fill out the space
            For ii = 1 To Repeat
                Call Lib.PushByte(What:=BIN.B(i), Where:=Output)
            Next ii
            'Don't count the length byte as data
            Let i = i + 1: Let Previous = -1
        Else
            'Normal data? Add as-is
            Call Lib.PushByte(What:=BIN.B(i), Where:=Output)
            Let Previous = BIN.B(i)
        End If
    Next i
    Let DecompressRLEData = Output
End Function

'CompressRLEData : Compress a level layout to be written to the ROM _
 ======================================================================================
Private Function CompressRLEData(ByRef Data() As Byte) As Byte()
    Dim i As Long, Output() As Byte
    
    'This will be used to keep track of which byte is currently repeating
    Dim Previous As Integer, Count As Integer
    Let Previous = -1
    
    'loop through the source data byte by byte
    For i = LBound(Data) To UBound(Data)
        'If this byte is the same as the previous byte then start counting
        If Data(i) = Previous Then Let Count = Count + 1
        
        'If the end is reached before the repeating data ends, cut it short here
        If i = UBound(Data) And Count > 0 Then
            'Write out the length of the final repeating data
            If Count = 256 Then Let Count = 0
            Call Lib.PushByte(Where:=Output, What:=CByte(Previous))
            Call Lib.PushByte(Where:=Output, What:=CByte(Count))
        
        'Can't repeat more than 256 times
        ElseIf Count = 256 Then
            'Due to what appears to be an overflow bug in the original ROM, _
             if a section is 256 bytes long, write it as 0
            Call Lib.PushByte(Where:=Output, What:=CByte(Previous))
            Call Lib.PushByte(Where:=Output, What:=0)
            Let Count = 0: Let Previous = -1
        
        'If counting and bytes differ, the repeating data has ended
        ElseIf (Data(i) <> Previous And Count > 0) Or Count = 256 Then
            'The repeating data has ended, write the compressed data and the stray byte
            Call Lib.PushByte(Where:=Output, What:=CByte(Previous))
            Call Lib.PushByte(Where:=Output, What:=CByte(Count))
            Call Lib.PushByte(Where:=Output, What:=Data(i))
            Let Count = 0: Let Previous = Data(i)
        
        ElseIf Data(i) <> Previous Then
            'Write any differing bytes as-is
            Call Lib.PushByte(Where:=Output, What:=Data(i))
            Let Previous = Data(i)
        End If
    Next i
    
    Let CompressRLEData = Output
End Function

'DecodeSMSTileRow : Take an SMS 4-byte value and decode into 8 palette indexes _
 ======================================================================================
Private Function DecodeSMSTileRow( _
    ByVal Byte1 As Byte, ByVal Byte2 As Byte, _
    ByVal Byte3 As Byte, ByVal Byte4 As Byte _
) As Byte()
    'Produced with the help of: _
     "Guide to the Sega Master System (0.02) (Super Majik Spiral Crew)" _
     <emu-docs.org/Master%20System/>
     
    'The SMS uses a palette of 16 colours, therefore the data that defines the tile _
     graphics do not specify the actual R/G/B values, but just an index (0-15) to the _
     palette colour to use. Therefore an 8-pixel row is stored as 4 bytes, where _
     4 bits (values 0-15) are assigned to each pixel.
    'However, instead of being stored as nybbles (4-bits in a row), the data is stored _
     across 4 bit-planes. To understand what bit-planes are, imagine 4 bytes stacked _
     on top of each other like layers in a cake. A bit plane would be a slice of the _
     cake -- 1 bit from each of the 4 bytes.
    'To decode the 4 bytes into 8 colour pixels we will need to slice the bit-planes _
     to get the index values (0-15) for each pixel and then look up the final colour _
     from the palette. This function doesn't deal with the final colours directly, _
     it just passes back the indexes to the palette
    
    'This is where we'll keep the palette indexes as we decode them
    Dim Indexes(0 To 7) As Byte
    
    Dim B As Long
    'The bit-planes are ordered the same direction as a byte, right-to-left, _
     so the first pixel on screen is the 8th bit
    For B = 7 To 0 Step -1
        'Take 1 bit from each byte
        If Byte1 And (2 ^ B) Then Indexes(7 - B) = Indexes(7 - B) Or 1
        If Byte2 And (2 ^ B) Then Indexes(7 - B) = Indexes(7 - B) Or 2
        If Byte3 And (2 ^ B) Then Indexes(7 - B) = Indexes(7 - B) Or 4
        If Byte4 And (2 ^ B) Then Indexes(7 - B) = Indexes(7 - B) Or 8
    Next
    
    DecodeSMSTileRow = Indexes
End Function

'EncodeSMSTileRow _
 ======================================================================================
Private Function EncodeSMSTileRow( _
    ByVal Index1 As Byte, ByVal Index2 As Byte, ByVal Index3 As Byte, _
    ByVal Index4 As Byte, ByVal Index5 As Byte, ByVal Index6 As Byte, _
    ByVal Index7 As Byte, ByVal Index8 As Byte _
) As Byte()
    Dim Bytes(0 To 3) As Byte
    
    Dim B As Long
    For B = 0 To 3
        If Index8 And (2 ^ B) Then Bytes(B) = Bytes(B) Or (2 ^ 7)
        If Index7 And (2 ^ B) Then Bytes(B) = Bytes(B) Or (2 ^ 6)
        If Index6 And (2 ^ B) Then Bytes(B) = Bytes(B) Or (2 ^ 5)
        If Index5 And (2 ^ B) Then Bytes(B) = Bytes(B) Or (2 ^ 4)
        If Index4 And (2 ^ B) Then Bytes(B) = Bytes(B) Or (2 ^ 3)
        If Index3 And (2 ^ B) Then Bytes(B) = Bytes(B) Or (2 ^ 2)
        If Index2 And (2 ^ B) Then Bytes(B) = Bytes(B) Or (2 ^ 1)
        If Index1 And (2 ^ B) Then Bytes(B) = Bytes(B) Or (2 ^ 0)
    Next B
    
    Let EncodeSMSTileRow = Bytes
End Function

'DecodeSMSColour : Convert Master System 6-bit colours to VB 24-bit colours _
 ======================================================================================
Private Function DecodeSMSColour(ByVal Colour As Byte) As Long
    '"The SMS use 6-bit colors, so, each color use one byte (the two first bits of the _
      byte are 0). This color can be defined like that : R(0-3), G(0-3), B(0-3). _
      So, there are only 4 possible values for each primary color. _
      Here are the RGB correspondance : 0=0, 1=80, 2=175, 3=255. _
      So, you can have a maximum of 64 colors. _
      The format for the byte is (in bits) : 00RR GGBB _
      So, the value of the byte shouldn't be greather than 3F."
     '<sonicology.fateback.com/hacks/s1smsrom.htm>
    Dim Red As Byte, Blue As Byte, Green As Byte
    
    If Colour > &H3F Then Stop
    
    Dim Luminance As Variant
    Let Luminance = Array(0, 80, 175, 255)
    
    'Isolate bits 7+8
    Let Red = Colour And 3
    'Isolate bits 5+6 and shift right twice
    Let Green = (Colour And 12) \ 4
    'Isolate bits 3+4 and shift right four times
    Let Blue = (Colour And 48) \ 16
    
    Let DecodeSMSColour = RGB( _
        Red:=Luminance(Red), Green:=Luminance(Green), Blue:=Luminance(Blue) _
    )
End Function

'EncodeSMSColour _
 ======================================================================================
Private Function EncodeSMSColour(Colour As Long) As Byte
    '
End Function
