Attribute VB_Name = "Lib"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Lib

'A bunch of common functions VB should have had built-in

'Used in converting colours to Hue / Saturation / Lightness
Public Type HSL
    Hue As Long
    Saturation As Long
    Luminance As Long
End Type

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'ArrayDimmed : Is an array dimmed? _
 ======================================================================================
'Taken from: https://groups.google.com/forum/?_escaped_fragment_=msg/microsoft.public.vb.general.discussion/3CBPw3nMX2s/zCcaO-hiCI0J#!msg/microsoft.public.vb.general.discussion/3CBPw3nMX2s/zCcaO-hiCI0J
Public Function ArrayDimmed(varArray As Variant) As Boolean
    Dim pSA As Long
    'Make sure an array was passed in:
    If IsArray(varArray) Then
        'Get the pointer out of the Variant:
        Call WIN32.kernel32_RtlMoveMemory( _
            ptrDestination:=pSA, ptrSource:=ByVal VarPtr(varArray) + 8, Length:=4 _
        )
        If pSA Then
            'Try to get the descriptor:
            Call WIN32.kernel32_RtlMoveMemory( _
                ptrDestination:=pSA, ptrSource:=ByVal pSA, Length:=4 _
            )
            'Array is initialized only if we got the SAFEARRAY descriptor:
            Let ArrayDimmed = (pSA <> 0)
        End If
    End If
End Function

'PushByte : Expand an array, adding a byte on the end _
 ======================================================================================
Public Function PushByte(ByVal What As Byte, ByRef Where() As Byte) As Long
    If ArrayDimmed(Where) Then
        'We will return the new length
        Let PushByte = UBound(Where) + 1
        ReDim Preserve Where(LBound(Where) To PushByte) As Byte
    Else
        'Array is not dimmed "()", begin at 0
        Let PushByte = 0
        ReDim Where(0) As Byte
    End If
    'Add the data
    Let Where(PushByte) = What
End Function

'Push Long : Expand an array, adding a long value on the end _
 ======================================================================================
Public Function PushLong( _
    ByVal What As Long, ByRef Where() As Long, _
    Optional ByVal AllowDuplicates As Boolean = True _
) As Boolean
    If ArrayDimmed(Where) Then
        'Is this a duplicate value?
        If AllowDuplicates = False Then
            'Don't add the value; return false to notify this
            If Lib.InArray(What, Where) = True Then Exit Function
        End If
        'Expand the array
        ReDim Preserve Where(LBound(Where) To UBound(Where) + 1) As Long
    Else
        'Array is not dimmed "()", begin at 0
        ReDim LongArray(0) As Long
    End If
    'Add the data
    Let Where(UBound(Where)) = What
    Let PushLong = True
End Function

'BytesToHex : Dump a byte array as a hexadecimal string _
 ======================================================================================
Public Function BytesToHex(var() As Byte) As String
    Dim i As Long
    For i = LBound(var) To UBound(var)
        BytesToHex = BytesToHex & Right("0" & Hex(var(i)), 2)
    Next i
End Function

'Exists : Check if an item exists in a Collection object _
 ======================================================================================
'<stackoverflow.com/questions/40651/check-if-a-record-exists-in-a-vb6-collection/9535221#9535221>
Public Function Exists(ByVal Key As String, ByRef Col As Collection) As Boolean
    Dim var As Variant
    
TryObject:
    On Error GoTo ExistsTryObject
    Set var = Col(Key)
    Let Exists = True
    Exit Function

TryNonObject:
    On Error GoTo ExistsTryNonObject
    Let var = Col(Key)
    Let Exists = True
    Exit Function

ExistsTryObject:
    'This will reset your Err Handler
    Resume TryNonObject

ExistsTryNonObject:
    Let Exists = False
End Function

'In Array : Check if a long value exists in an array _
 ======================================================================================
Public Function InArray(ByVal What As Long, ByRef Where() As Long) As Boolean
    Let InArray = False
    If Lib.ArrayDimmed(Where) = False Then Exit Function
    
    'This is slow, but it doesn't rely on creating an array error which would crash _
     the executable if we disable array bound checking for speed purposes
    Dim Index As Long
    For Index = LBound(Where) To UBound(Where)
        If Where(Index) = What Then Let InArray = True: Exit Function
    Next
End Function

'CombSort : Sorty an array _
 ======================================================================================
'<www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-%28sort-array-sorting-arrays%29&p=2909248#post2909248>
Public Sub CombSort(ByRef pvarArray As Variant)
    Const ShrinkFactor = 1.3
    Dim lngGap As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
   
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray)
    lngGap = iMax - iMin + 1
    Do
        If lngGap > 1 Then
            lngGap = Int(lngGap / ShrinkFactor)
            If lngGap = 10 Or lngGap = 9 Then lngGap = 11
        End If
        blnSwapped = False
        For i = iMin To iMax - lngGap
            If pvarArray(i) > pvarArray(i + lngGap) Then
                varSwap = pvarArray(i)
                pvarArray(i) = pvarArray(i + lngGap)
                pvarArray(i + lngGap) = varSwap
                blnSwapped = True
            End If
        Next
    Loop Until lngGap = 1 And Not blnSwapped
End Sub

'FileExists : See if a file exists or not _
 ======================================================================================
Public Function FileExists(fName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Let Temp = VBA.GetAttr(fName)
    Let FileExists = Not CBool(Err)
End Function

'RoundUp : Always round a number upwards _
 ======================================================================================
Public Function RoundUp(ByVal InputNumber As Double) As Double
    If Int(InputNumber) = InputNumber _
        Then Let RoundUp = InputNumber _
        Else Let RoundUp = Int(InputNumber) + 1
End Function

'Min : Limit a number to a minimum value _
 ======================================================================================
Public Function Min(ByVal InputNumber As Long, Optional ByVal Minimum As Long = 0) As Long
    Let Min = IIf(InputNumber < Minimum, Minimum, InputNumber)
End Function

'Max : Limit a number to a maximum value _
 ======================================================================================
Public Function Max(ByVal InputNumber As Long, Optional ByVal Maximum As Long = 2147483647) As Long
    Let Max = IIf(InputNumber > Maximum, Maximum, InputNumber)
End Function

'Range : Limit a number to a minimum and maximum value _
 ======================================================================================
Public Function Range(ByVal InputNumber As Long, Optional ByVal Maximum = 2147483647, Optional ByVal Minimum = -2147483648#) As Long
    Let Range = Max(Min(InputNumber, Minimum), Maximum)
End Function

'RGBToHSL : Convert Red, Green, Blue to Hue, Saturation, Lightness _
 ======================================================================================
'<www.xbeat.net/vbspeed/c_RGBToHSL.htm>
Public Function RGBToHSL(ByVal RGBValue As Long) As HSL
    'by Paul - wpsjr1@syix.com, 20011120
    Dim r As Long, G As Long, B As Long
    Dim lMax As Long, lMin As Long
    Dim q As Single
    Dim lDifference As Long
    Static Lum(255) As Long
    Static QTab(255) As Single
    Static init As Long
    
    If init = 0 Then
        For init = 2 To 255 ' 0 and 1 are both 0
            Lum(init) = init * 100 / 255
        Next
        For init = 1 To 255
            QTab(init) = 60 / init
        Next init
    End If
    
    r = RGBValue And &HFF
    G = (RGBValue And &HFF00&) \ &H100&
    B = (RGBValue And &HFF0000) \ &H10000
    
    If r > G Then
        lMax = r: lMin = G
    Else
        lMax = G: lMin = r
    End If
    If B > lMax Then
        lMax = B
    ElseIf B < lMin Then
        lMin = B
    End If
    
    RGBToHSL.Luminance = Lum(lMax)
    
    lDifference = lMax - lMin
    If lDifference Then
        'Do a 65K 2D lookup table here for more speed if needed
        RGBToHSL.Saturation = (lDifference) * 100 / lMax
        q = QTab(lDifference)
        Select Case lMax
            Case r
                If B > G Then
                    RGBToHSL.Hue = q * (G - B) + 360
                Else
                    RGBToHSL.Hue = q * (G - B)
                End If
            Case G
                RGBToHSL.Hue = q * (B - r) + 120
            Case B
                RGBToHSL.Hue = q * (r - G) + 240
        End Select
    End If
End Function

'HSLToRGB : Convert Hue, Saturation, Ligthness to (roughly) Red, Green, Blue _
 ======================================================================================
'<www.xbeat.net/vbspeed/c_HSLToRGB.htm>
Public Function HSLToRGB( _
    ByVal Hue As Long, _
    ByVal Saturation As Long, _
    ByVal Luminance As Long _
) As Long
    'by Donald (Sterex 1996), donald@xbeat.net, 20011124
    Dim r As Long, G As Long, B As Long
    Dim lMax As Long, lMid As Long, lMin As Long
    Dim q As Single

    lMax = (Luminance * 255) / 100
  
    If Saturation > 0 Then

        lMin = (100 - Saturation) * lMax / 100
        q = (lMax - lMin) / 60
        
        Select Case Hue
            Case 0 To 60
                lMid = (Hue - 0) * q + lMin
                r = lMax: G = lMid: B = lMin
            Case 60 To 120
                lMid = -(Hue - 120) * q + lMin
                r = lMid: G = lMax: B = lMin
            Case 120 To 180
                lMid = (Hue - 120) * q + lMin
                r = lMin: G = lMax: B = lMid
            Case 180 To 240
                lMid = -(Hue - 240) * q + lMin
                r = lMin: G = lMid: B = lMax
            Case 240 To 300
                lMid = (Hue - 240) * q + lMin
                r = lMid: G = lMin: B = lMax
            Case 300 To 360
                lMid = -(Hue - 360) * q + lMin
                r = lMax: G = lMin: B = lMid
        End Select
        HSLToRGB = B * &H10000 + G * &H100& + r
    Else
        HSLToRGB = lMax * &H10101
    End If
End Function

