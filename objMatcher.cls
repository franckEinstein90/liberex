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

Private m_patternDictionary As New Scripting.dictionary
Private m_pattern As String

Public Sub addPattern(ByVal patternName As String, ByVal pattern As String)
    Dim regex As New RegExp
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "^" & pattern & "$"
    End With
    m_patternDictionary.Add key:=patternName, Item:=regex
End Sub

Public Sub addFieldDescriptor(ByVal fieldPattern As String)
    m_pattern = m_pattern & "(" & fieldPattern & ")"
End Sub

Public Sub match( _
    processor As IDataProcessor, _
    ByVal recordInformation As String)
    
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
            Call processor.process(key, values)
            GoTo ExitMatcher
        End If
    Next key
    Stop 'no match?
ExitMatcher:
End Sub

