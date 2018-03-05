Option Compare Database
Option Explicit

'Encapsulate raw information associated
'with a file when read from a data source
Private recordFields As Scripting.dictionary
Public fields As Scripting.dictionary

Public Property Get toString() As String
    Dim field As Variant
    For Each field In fields.Keys
        toString = toString & field & ":" & fields.Item(field) & vbCrLf
    Next field
End Property

Public Sub value(ByVal fieldName As String, ByVal fieldValue As String)
    If fields.exists(fieldName) Then
        fields(fieldName) = fieldValue
    End If
End Sub

Public Sub clear()
    Dim key As Variant
    For Each key In recordFields.Keys
        fields(key) = ""
    Next key
End Sub

Public Sub populateFromTabDelimitedString(ByVal rawInfo As String)
    'Takes a string that contains the raw information
    'as argument, parses the argument string
    'and populates the "fields" dictionary accordingly
    Dim InfoFields() As String
    Dim index As Integer
    Call clear
    InfoFields = Split(rawInfo, vbTab)
    For index = 1 To UBound(InfoFields) + 1
        fields.Item(recordFields.Item(index)) = InfoFields(index - 1)
    Next
    
End Sub


Public Sub initialize(ByVal scanTypeID As Long)
    Dim rs As Recordset
    Dim fieldTag As String
    Dim scanColIDX As Integer
    Dim correspondingStdField As String
    
    Set fields = New Scripting.dictionary
  
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT ColIDX, ScanCol, CorrespondingInputFileInfoStructField FROM" & _
        " tblScanFormats WHERE TSPScanType = " & scanTypeID & " ORDER BY ColIDX")
    
    With rs
        If .RecordCount > 0 Then
            Set recordFields = New Scripting.dictionary
            Set fields = New Scripting.dictionary
            
            .MoveFirst
            While Not .EOF
                scanColIDX = rs!colIdx
                correspondingStdField = rs!CorrespondingInputFileInfoStructField
                recordFields.Add key:=scanColIDX, Item:=correspondingStdField
                fields.Add key:=correspondingStdField, Item:=""
                .MoveNext
            Wend
        End If
        .Close
    End With
    Set rs = Nothing
End Sub
