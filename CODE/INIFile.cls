VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: INIFile

'Allow reading and writing of data with a ".ini" file. _
 This class does not use the deprecated `Get/WritePrivateProfileString` APIs as they _
 are buggy and may potentially (though not likely) be removed from Windows

'Status             INCOMPLETE
'Dependencies       Lib.bas, Scripting [scrrun.dll]
'Last Updated       04-OCT-13
'Last Update        Created class

'--------------------------------------------------------------------------------------

'Why not use the registry? That's no good for a portable application, and also it's _
 better for backup / restore (most people don't backup and restore selective parts _
 of the registry)

'Why not XML? Too slow, too fragile. If users edit XML then the strictness is a _
 killer and just gets in the way of storing very simple data. XML demands too much _
 perfectness to expect of end users who may be unfamiliar with it. INI's line-by-line _
 nature means errors are isolated and one error doesn't break the entire file

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'The current file being managed
Private My_FilePath As String
'The INI file will be serialised into a Dictionary object
Private Data As Scripting.Dictionary

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Set Data = New Scripting.Dictionary
    Let Data.CompareMode = TextCompare
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Set Data = Nothing
End Sub

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'PROPERTY FilePath : Which file to work with _
 ======================================================================================
Public Property Get FilePath() As String: Let FilePath = My_FilePath: End Property
Public Property Let FilePath(ByVal NewPath As String)
    'Clear up any existing file (you should have called save before doing this)
    Call Class_Terminate: Call Class_Initialize
    
    'This class does not care if the file doesn't exist since we are reading settings _
     that will have a default value and the INI file (and its properties) will exist _
     usually only if these settings are different from their defaults
    Let My_FilePath = NewPath
    If Lib.FileExists(My_FilePath) = False Then Exit Property
    
    'Read in and convert the INI file into a Dictionary heirarchy
    On Error GoTo Fail
    Dim FileNumber As Integer: Let FileNumber = FreeFile()
    Open FilePath For Input Lock Write As #FileNumber
    
    'This will reference the current INI section the properties are being added to
    Dim Section As Scripting.Dictionary
    Set Section = Data
    
    Do While Not EOF(FileNumber)
        Dim Line As String
        Line Input #FileNumber, Line
        Let Line = Trim$(Line)
        
        'Is this a section?
        If InStr(1, Line, "[") > 0 Then
            'Does this section already exist?
            If Data.Exists(Line) = True Then
                'If the section is defined twice, just re-use it
                Set Section = Data.Item(Line)
            Else
                'Add a new Dictionary to the root and send new properties there
                Set Section = New Scripting.Dictionary
                Call Data.Add(Line, Section)
            End If
        End If
        
        'Is this a comment or blank line?
        If Line = vbNullString Or InStr(1, Line, ";") > 0 Or InStr(1, Line, "#") > 0 Then
            'Add it wholesale
            Call Section.Add("=COMMENT:" & Rnd(1) * Timer, Line)
        
        'Finally, this should be a property
        ElseIf InStr(1, Line, "=") > 0 Then
            'Divide at the equals sign and assign key and value to the Dictionary
            Call Section.Add( _
                Mid$(Line, 1, InStr(1, Line, "=") - 1), _
                Mid$(Line, InStr(1, Line, "=") + 1) _
            )
        End If
    Loop
    
    Close #FileNumber
    Exit Property

Fail:
    'If the file was left open at the point of error, try to free the handle
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'GetString _
 ======================================================================================
Public Function GetString( _
    ByVal Property As String, Optional ByVal Section As String = vbNullString, _
    Optional ByVal Default As String = vbNullString _
) As String
    Let GetString = GetProperty(Property, Section)
    If GetString = vbNullString Then Let GetString = Default
End Function

'GetLong _
 ======================================================================================
Public Function GetLong( _
    ByVal Property As String, Optional ByVal Section As String = vbNullString, _
    Optional ByVal Default As Long = 0 _
) As Long
    Dim INIValue As String
    Let INIValue = GetProperty(Property, Section)
    If INIValue = vbNullString Then
        Let GetLong = Default
    Else
        On Error GoTo Fail
        Let GetLong = CDbl(INIValue)
        Exit Function
    End If
Fail:
    Let GetLong = Default
End Function

'GetSingle _
 ======================================================================================
Public Function GetSingle( _
    ByVal Property As String, Optional ByVal Section As String = vbNullString, _
    Optional ByVal Default As Single = 0 _
) As Single
    Dim INIValue As String
    Let INIValue = GetProperty(Property, Section)
    If INIValue = vbNullString Then
        Let GetSingle = Default
    Else
        On Error GoTo Fail
        Let GetSingle = CSng(INIValue)
        Exit Function
    End If
Fail:
    Let GetSingle = Default
End Function

'GetDouble : Retrieve a value converted to a double-precision (64-bit) number _
 ======================================================================================
Public Function GetDouble( _
    ByVal Property As String, Optional ByVal Section As String = vbNullString, _
    Optional ByVal Default As Double = 0 _
) As Double
    Dim INIValue As String
    Let INIValue = GetProperty(Property, Section)
    If INIValue = vbNullString Then
        Let GetDouble = Default
    Else
        On Error GoTo Fail
        Let GetDouble = CDbl(INIValue)
        Exit Function
    End If
Fail:
    Let GetDouble = Default
End Function

'SetValue _
 ======================================================================================
Public Function SetValue( _
    ByVal Value As Variant, ByVal Property As String, _
    Optional ByVal Section As String = vbNullString _
) As Double
    Call SetProperty(CStr(Value), Property, Section)
End Function

'Save _
 ======================================================================================
Public Function Save() As Boolean
    'If no file has been loaded, there's nothing to save
    If My_FilePath = vbNullString Then Exit Function
    
    On Error GoTo Fail
    Dim FileNumber As Integer: Let FileNumber = FreeFile
    Open My_FilePath For Output Lock Read As #FileNumber
    
    Dim i As Long, ii As Long
    For i = 0 To Data.Count - 1
        'Is this item a section
        If (TypeOf Data.Items(i) Is Scripting.Dictionary) = True Then
            Print #FileNumber, Data.Keys(i)
            
            Dim Section As Scripting.Dictionary
            Set Section = Data.Items(i)
            'Now write the sub-items
            For ii = 0 To Section.Count - 1
                'Is this a comment / blank line?
                If Left$(Section.Items(ii), 1) = "=" Then
                    Print #FileNumber, Section.Items(ii)
                
                'This is a property
                Else
                    Print #FileNumber, Section.Keys(ii) & "=" & Section.Items(ii)
                End If
            Next ii
            Set Section = Nothing
        
        'Is this item a comment / blank line?
        ElseIf Left$(Data.Item(i), 1) = "=" Then
            Print #FileNumber, Data.Items(i)
        
        'This is a property
        Else
            Print #FileNumber, Data.Keys(i) & "=" & Data.Items(i)
        End If
    Next i
    
Fail:
    'If the file was left open at the point of error, try to free the handle
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'GetProperty : Retrieve a string from the INI file _
 ======================================================================================
Private Function GetProperty( _
    ByVal Property As String, _
    Optional ByVal Section As String = vbNullString _
) As String
    'Are we referencing the root section?
    If Section = vbNullString Then
        If Data.Exists(Property) = True Then Let GetProperty = CStr(Data(Property))
    Else
        'Check if the section exists
        If Data.Exists("[" & Section & "]") = True Then
            'Now check if the property exists
            If Data("[" & Section & "]").Exists(Property) = True Then
                'Finally get the value
                Let GetProperty = Data.Item("[" & Section & "]").Item(Property)
            End If
        End If
    End If
End Function

'SetProperty : Update or add a string value _
 ======================================================================================
Private Sub SetProperty( _
    ByVal Value As String, ByVal Property As String, _
    Optional ByVal Section As String = vbNullString _
)
    'This will reference which section the property will go into
    Dim Destination As Object
    
    'Has a section been specified?
    If Section <> vbNullString Then
        'Does the section even exist?
        If Data.Exists("[" & Section & "]") Then
            'The section already exists, set this as the destination
            Set Destination = Data("[" & Section & "]")
            'Now check if the property exists / &c.
            GoTo CheckProperty
        Else
            'Create the section
            Call Data.Add("[" & Section & "]", New Scripting.Dictionary)
            Set Destination = Data.Item("[" & Section & "]")
            'Fall through here and check the property
        End If
    Else
        'Root section
        Set Destination = Data
        'Fall through here and check the property
    End If

CheckProperty:
    'At this point the correct section has been selected / created, _
     now test if the property already exists
    If Destination.Exists(Property) = False Then
        'No? Create the property
        Call Destination.Add(Property, Value)
    Else
        'Update the existing value
        Let Destination.Item(Property) = Value
    End If
    
CleanUp:
    Set Destination = Nothing
End Sub
