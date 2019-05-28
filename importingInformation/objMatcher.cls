VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objMatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Microsoft Scripting Runtime"
'Microsoft VBScript Regular Expression 5.5

Private m_patternDictionary As New Scripting.Dictionary


Public Sub addPattern(ByVal patternName As String, ByVal pattern As String)
    Dim regex As New RegExp
    If Left(pattern, 1) = "(" And Right(pattern, 1) = ")" Then GoTo BuildRegex
    pattern = "(" & pattern & ")"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
BuildRegex:
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "^" & pattern & "$"
    End With
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    m_patternDictionary.Add key:=patternName, Item:=regex
End Sub

Public Sub addFieldDescriptor(ByVal fieldPattern As String)
    m_pattern = m_pattern & "(" & fieldPattern & ")"
End Sub

Public Function match( _
    processor As IDataProcessor, _
    ByVal recordInformation As String) As Variant
    
    Dim key As Variant
    Dim testResult As Boolean
    Dim matches As Object
    Dim values As New Collection
    Dim ctr As Long
    
    For Each key In m_patternDictionary.Keys
        testResult = m_patternDictionary(key).Test(recordInformation)
        If testResult Then
            Set matches = m_patternDictionary(key).Execute(recordInformation)
            For ctr = 0 To matches(0).SubMatches.count - 1
                values.Add matches(0).SubMatches(ctr)
            Next ctr
            match = processor.process(key, values)
            GoTo EndFunction
        End If
    Next key
    values.Add (recordInformation)
    match = processor.process("", values)
EndFunction:
 End Function

