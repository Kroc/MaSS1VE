VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S1Sprites"
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
'CLASS :: S1Sprites

'A class for holding a load of cached sprite images. The graphics from the ROM are _
 stored as tile sheets (a strip of 8x8 tiles) and the actual graphics have to be _
 pieced together out of these. This class pieces together the sprites into images _
 that can be quickly bltted where-ever needed. This is quite complex because this _
 class has to support palette changes for each level and for the underwater effect

'NOTE: At this time, the parts of the ROM that define which level has which sprites _
 is not yet documented. We therefore have to make some unfortunate assumptions that _
 lock us into the existing level order and themes

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

Private My_SpriteArt As S1Tileset
Private My_Palette As S1Palette

Private My_Sonic As bluImage

Private My_Monitor_Ring As bluImage
Private My_Monitor_Speed As bluImage
Private My_Monitor_Life As bluImage
Private My_Monitor_Shield As bluImage
Private My_Monitor_Stars As bluImage
Private My_Monitor_Check As bluImage
Private My_Monitor_Cont As bluImage
Private My_Emerald As bluImage

Private My_EndSign As bluImage

Private My_Badnick_Motobug As bluImage
Private My_Badnick_Crabmeat As bluImage
Private My_Badnick_BuzzBomber As bluImage
Private My_Badnick_Newtron As bluImage
Private My_Badnick_Chopper As bluImage


'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Erase the cached images
    Call Invalidate
    'Detach the level art & palette,
    Set My_SpriteArt = Nothing
    Set My_Palette = Nothing
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////
'We expose the images manually so that we can cache-on-demand (saves having to paint _
 every sprite in one go, especially those that go un-used on a level)

Public Property Set SpriteArt(NewSpriteArt As S1Tileset)
    Set My_SpriteArt = NewSpriteArt
    Call Invalidate
End Property

Public Property Get Palette() As S1Palette: Set Palette = My_Palette: End Property
Public Property Set Palette(NewPalette As S1Palette)
    Set My_Palette = NewPalette

    'Apply the new palette to any images that have already been cached
    Call My_Palette.ApplyToImage(My_Sonic)
    
    Call My_Palette.ApplyToImage(My_Monitor_Ring)
    Call My_Palette.ApplyToImage(My_Monitor_Speed)
    Call My_Palette.ApplyToImage(My_Monitor_Life)
    Call My_Palette.ApplyToImage(My_Monitor_Shield)
    Call My_Palette.ApplyToImage(My_Monitor_Stars)
    Call My_Palette.ApplyToImage(My_Monitor_Check)
    Call My_Palette.ApplyToImage(My_Monitor_Cont)
    Call My_Palette.ApplyToImage(My_Emerald)
    
    Call My_Palette.ApplyToImage(My_EndSign)
    
    Call My_Palette.ApplyToImage(My_Badnick_Motobug)
    Call My_Palette.ApplyToImage(My_Badnick_Crabmeat)
    Call My_Palette.ApplyToImage(My_Badnick_BuzzBomber)
    Call My_Palette.ApplyToImage(My_Badnick_Newtron)
    Call My_Palette.ApplyToImage(My_Badnick_Chopper)
End Property

Public Property Get Sonic() As bluImage
    If My_Sonic Is Nothing Then
        Set My_Sonic = New bluImage
        Call My_Sonic.Create8Bit( _
            ImageWidth:=8 * 3, ImageHeight:=8 * 4, _
            Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
        )
        With GAME.Sonic
            Call .ApplyPalette(My_Palette)
            'First row
            Call .PaintSprite(My_Sonic.hDC, 0, 0, 72)
            Call .PaintSprite(My_Sonic.hDC, 8, 0, 74)
            Call .PaintSprite(My_Sonic.hDC, 16, 0, 76)
            'Second row
            Call .PaintSprite(My_Sonic.hDC, 0, 16, 78)
            Call .PaintSprite(My_Sonic.hDC, 8, 16, 80)
            Call .PaintSprite(My_Sonic.hDC, 16, 16, 82)
        End With
    End If
    Set Sonic = My_Sonic
End Property

Public Property Get Monitor_Ring() As bluImage
    If My_Monitor_Ring Is Nothing Then Set My_Monitor_Ring = CreateMonitor(0)
    Set Monitor_Ring = My_Monitor_Ring
End Property

Public Property Get Monitor_Speed() As bluImage
    If My_Monitor_Speed Is Nothing Then Set My_Monitor_Speed = CreateMonitor(4)
    Set Monitor_Speed = My_Monitor_Speed
End Property

Public Property Get Monitor_Life() As bluImage
    If My_Monitor_Life Is Nothing Then Set My_Monitor_Life = CreateMonitor(8)
    Set Monitor_Life = My_Monitor_Life
End Property

Public Property Get Monitor_Shield() As bluImage
    If My_Monitor_Shield Is Nothing Then Set My_Monitor_Shield = CreateMonitor(12)
    Set Monitor_Shield = My_Monitor_Shield
End Property

Public Property Get Monitor_Stars() As bluImage
    If My_Monitor_Stars Is Nothing Then Set My_Monitor_Stars = CreateMonitor(16)
    Set Monitor_Stars = My_Monitor_Stars
End Property

Public Property Get Monitor_Check() As bluImage
    If My_Monitor_Check Is Nothing Then Set My_Monitor_Check = CreateMonitor(24)
    Set Monitor_Check = My_Monitor_Check
End Property

Public Property Get Monitor_Cont() As bluImage
    If My_Monitor_Cont Is Nothing Then Set My_Monitor_Cont = CreateMonitor(28)
    Set Monitor_Cont = My_Monitor_Cont
End Property

Public Property Get Emerald() As bluImage
    If My_Emerald Is Nothing Then
        Set My_Emerald = New bluImage
        'The emerald is located at tiles 20-23 on the power-ups tileset, but is not _
         painted in a monitor (it appears the original plan was to have it as a _
         monitor, but was later changed)
        Call My_Emerald.Create8Bit( _
            ImageWidth:=8 * 2, ImageHeight:=8 * 2, _
            Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
        )
        Call GAME.PowerUps.ApplyPalette(My_Palette)
        Call GAME.PowerUps.PaintSprite(My_Emerald.hDC, 0, 0, 20)
        Call GAME.PowerUps.PaintSprite(My_Emerald.hDC, 8, 0, 22)
    End If
    Set Emerald = My_Emerald
End Property

Public Property Get EndSign() As bluImage
    If My_EndSign Is Nothing Then
        Set My_EndSign = New bluImage
        Call My_EndSign.Create8Bit( _
            ImageWidth:=8 * 4, ImageHeight:=8 * 5, _
            Palette_LongArray:=GAME.EndSignPalette.Colours, UseTransparency:=True _
        )
        With GAME.EndSignTileset
            'First row - top of sign
            .PaintTile My_EndSign.hDC, 0, 0, 79
            .PaintTile My_EndSign.hDC, 8, 0, 81
            .PaintTile My_EndSign.hDC, 16, 0, 83
            .PaintTile My_EndSign.hDC, 24, 0, 85
            'Second and third rows - middle of sign
            .PaintSprite My_EndSign.hDC, 0, 8, 110
            .PaintSprite My_EndSign.hDC, 8, 8, 112
            .PaintSprite My_EndSign.hDC, 16, 8, 114
            .PaintSprite My_EndSign.hDC, 24, 8, 116
            'The stem
            .PaintSprite My_EndSign.hDC, 8, 24, 66
            .PaintSprite My_EndSign.hDC, 16, 24, 68
            .PaintTile My_EndSign.hDC, 8, 40, 67
            .PaintTile My_EndSign.hDC, 16, 40, 69
        End With
    End If
    Set EndSign = My_EndSign
End Property

Public Property Get Badnick_Motobug() As bluImage
    If My_Badnick_Motobug Is Nothing Then
        Set My_Badnick_Motobug = New bluImage
        Call My_Badnick_Motobug.Create8Bit( _
            ImageWidth:=8 * 2, ImageHeight:=8 * 2, _
            Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
        )
        With My_SpriteArt
            Call .PaintSprite(My_Badnick_Motobug.hDC, 0, 0, 96)
            Call .PaintSprite(My_Badnick_Motobug.hDC, 8, 0, 98)
        End With
    End If
    Set Badnick_Motobug = My_Badnick_Motobug
End Property

Public Property Get Badnick_Crabmeat() As bluImage
    If My_Badnick_Crabmeat Is Nothing Then
        Set My_Badnick_Crabmeat = New bluImage
        Call My_Badnick_Crabmeat.Create8Bit( _
            ImageWidth:=8 * 3, ImageHeight:=8 * 3, _
            Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
        )
        With My_SpriteArt
            .PaintTile My_Badnick_Crabmeat.hDC, 0, 0, 1
            .PaintTile My_Badnick_Crabmeat.hDC, 8, 0, 3
            .PaintTile My_Badnick_Crabmeat.hDC, 16, 0, 5
            .PaintSprite My_Badnick_Crabmeat.hDC, 0, 8, 32
            .PaintSprite My_Badnick_Crabmeat.hDC, 8, 8, 34
            .PaintSprite My_Badnick_Crabmeat.hDC, 16, 8, 36
        End With
    End If
    Set Badnick_Crabmeat = My_Badnick_Crabmeat
End Property

Public Property Get Badnick_BuzzBomber() As bluImage
    If My_Badnick_BuzzBomber Is Nothing Then
        Set My_Badnick_BuzzBomber = New bluImage
        Call My_Badnick_BuzzBomber.Create8Bit( _
            ImageWidth:=8 * 3, ImageHeight:=8 * 4, _
            Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
        )
        With My_SpriteArt
            Call .PaintSprite(My_Badnick_BuzzBomber.hDC, 8, 0, 10)
            Call .PaintSprite(My_Badnick_BuzzBomber.hDC, 0, 16, 12)
            Call .PaintSprite(My_Badnick_BuzzBomber.hDC, 8, 16, 14)
            Call .PaintSprite(My_Badnick_BuzzBomber.hDC, 16, 16, 16)
        End With
    End If
    Set Badnick_BuzzBomber = My_Badnick_BuzzBomber
End Property

Public Property Get Badnick_Newtron() As bluImage
    If My_Badnick_Newtron Is Nothing Then
        Set My_Badnick_Newtron = New bluImage
        Call My_Badnick_Newtron.Create8Bit( _
            ImageWidth:=8 * 2, ImageHeight:=8 * 4, _
            Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
        )
        With My_SpriteArt
            Call .PaintSprite(My_Badnick_Newtron.hDC, 0, 0, 28)
            Call .PaintSprite(My_Badnick_Newtron.hDC, 8, 0, 30)
            Call .PaintSprite(My_Badnick_Newtron.hDC, 8, 16, 62)
        End With
    End If
    Set Badnick_Newtron = My_Badnick_Newtron
End Property

Public Property Get Badnick_Chopper() As bluImage
    If My_Badnick_Chopper Is Nothing Then
        Set My_Badnick_Chopper = New bluImage
        Call My_Badnick_Chopper.Create8Bit( _
            ImageWidth:=8 * 2, ImageHeight:=8 * 2, _
            Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
        )
        With My_SpriteArt
            Call .PaintSprite(My_Badnick_Chopper.hDC, 0, 0, 96)
            Call .PaintSprite(My_Badnick_Chopper.hDC, 8, 0, 98)
        End With
    End If
    Set Badnick_Chopper = My_Badnick_Chopper
End Property

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'CreateMonitor : The monitor sprites share a lot in common _
 ======================================================================================
Private Function CreateMonitor(Optional ByVal Index As Long = -1) As bluImage
    Set CreateMonitor = New bluImage
    With CreateMonitor
        Call .Create8Bit( _
            ImageWidth:=8 * 3, ImageHeight:=8 * 3, _
            Palette_LongArray:=My_Palette.Colours, UseTransparency:=True _
        )
        'Draw the monitor first
        Call My_SpriteArt.PaintSprite(.hDC, 0, 0, 84)
        Call My_SpriteArt.PaintSprite(.hDC, 8, 0, 86)
        Call My_SpriteArt.PaintSprite(.hDC, 16, 0, 88)
        'the base of the monitor is on the HUD tileset
        Call GAME.HUD.ApplyPalette(My_Palette)
        Call GAME.HUD.PaintTile(.hDC, 0, 16, 42)
        Call GAME.HUD.PaintTile(.hDC, 8, 16, 44)
        Call GAME.HUD.PaintTile(.hDC, 16, 16, 46)
        'If an index into the Power Ups tilesheet is given, paint the symbol on the monitor
        If Index >= 0 Then
            Call GAME.PowerUps.ApplyPalette(My_Palette)
            Call GAME.PowerUps.PaintSprite(.hDC, 4, 0, Index)
            Call GAME.PowerUps.PaintSprite(.hDC, 4 + 8, 0, Index + 2)
        End If
    End With
End Function

'Invalidate : If the palette or art changes, we need to wipe the existing graphics _
 ======================================================================================
Private Sub Invalidate()
    Set My_Sonic = Nothing
    
    Set My_Monitor_Ring = Nothing
    Set My_Monitor_Speed = Nothing
    Set My_Monitor_Life = Nothing
    Set My_Monitor_Shield = Nothing
    Set My_Monitor_Stars = Nothing
    Set My_Monitor_Check = Nothing
    Set My_Monitor_Cont = Nothing
    Set My_Emerald = Nothing
    
    Set My_EndSign = Nothing
    
    Set My_Badnick_Motobug = Nothing
    Set My_Badnick_Crabmeat = Nothing
    Set My_Badnick_BuzzBomber = Nothing
    Set My_Badnick_Newtron = Nothing
    Set My_Badnick_Chopper = Nothing
End Sub
