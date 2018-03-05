Option Compare Database
Option Explicit


'Class constants
Private Const tblScanTypes As String = "tblTSPScanTypes"

Private m_sourceFile As objComputerFile

'Information associated with a data grab
Public m_dataGrabID As Long
Public m_numFileUpdated As Long

Public m_businessOwnerID As Long
Public m_scan_file_type As String
Public is_group_data As Boolean

'************************************************
'cols format contains the indexes of each columns
'and their label
Public m_colsFormat As Scripting.dictionary

'************************************************
'For grouped data, the header_cols_format
'dictionary contains the necessary information
'to identify a row as a header
Public header_cols_format As Scripting.dictionary


'Each possible data source type has
'an object that can be used to read it
Private recordList As Variant


'************************************************
'Methods and variables to read from the data source
Private m_fileOpened As Boolean

Public numberOfRecords As Long
Public numberOfFileRecords As Long
Public numberOfFolderRecords As Long



Public Property Get noMoreRecords() As Boolean
    noMoreRecords = recordList.noMoreRecords
End Property
Public Property Get path() As String
    path = m_sourceFile.path
End Property
Public Property Get fileExists() As Boolean
    fileExists = m_sourceFile.exists
End Property



'makeFormatMask returns an empty dictionary, with keys corresponding to the dataSource object columns
Public Sub makeFormatMask(ByRef dictionary As Scripting.dictionary)
    
    Dim col_idx As Variant
    Dim col_name As String
    
    For Each col_idx In cols_format.Keys
        col_name = cols_format.Item(col_idx)
        dictionary.Add key:=col_name, Item:=""
    Next
End Sub



'***********************************************************
'***********Identifying data sections***********************
'***********************************************************
'identify a row in the data input as being the file header
Public Function isFileHeader(ByVal inputstr As String, ByVal delimiter As String) As Boolean
    Dim input_elements() As String
    Dim col_val As Variant
    Dim col_label As Variant
    Dim col_ctr As Integer
        
    isFileHeader = True
    col_ctr = 1
    input_elements = Split(inputstr, delimiter)
    If Not UBound(input_elements) = m_colsFormat.count - 1 Then
        isFileHeader = False
        Exit Function
    End If
    For Each col_val In input_elements
        col_label = m_colsFormat.Item(col_ctr)
        If Not col_val = col_label Then
            isFileHeader = False
            Exit Function
        End If
        col_ctr = col_ctr + 1
    Next col_val
End Function

Public Function is_group_header(ByVal inputstr As String, ByVal delimiter As String) As Boolean
'identifies the data as a group header

    Dim input_elements() As String
    Dim group_header_col_idx As Variant
    Dim identifying_regex As Object
    
    Set identifying_regex = CreateObject("VBScript.RegExp")
    input_elements = Split(inputstr, delimiter)
    For Each group_header_col_idx In header_cols_format.Keys
        identifying_regex.Pattern = header_cols_format(group_header_col_idx)
        If Not identifying_regex.Test(input_elements(group_header_col_idx - 1)) Then
            is_group_header = False
            Exit Function
        End If
    Next group_header_col_idx
    is_group_header = True
End Function

Public Sub openSource()
    Select Case m_sourceFile.m_extension
        Case "txt"
            Set recordList = New objDataSetText
        Case "xls", "xlsx"
            Set recordList = New objDataSetExcel
        Case Else
            MsgBox ("unable to glob content in this format")
            Call Err.Raise(10015, "Unable to read records")
            Exit Sub
    End Select
End Sub

Public Sub initialize(ByVal filePath As String)
    Set m_sourceFile = New objComputerFile
    With m_sourceFile
        .initialize (filePath)
        If Not .exists Then
            Err.Raise "Unable to find source data file"
        End If
    End With
    m_fileOpened = False
End Sub
