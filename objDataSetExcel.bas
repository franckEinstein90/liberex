Option Compare Database
Option Explicit

'Requires reference to Microsoft Excel Object Library

'Internals of worksheet
Private xlApp As Excel.Application
Private wksheets() As String
Private currentWKS As Integer

Private m_recordPointer As Long


'total number of records (1 per row) in data source
Private m_numRecords As Long
Private m_numColumns As Long


'************************************************
'Reading from an excel file

Private m_dataTable As Excel.ListObject
    
Public Property Get numRecords() As Long
    numRecords = m_numRecords
End Property
Public Property Get noMoreRecords() As Boolean
    If m_recordPointer > m_numRecords Then
        noMoreRecords = True
    Else
        noMoreRecords = False
    End If
End Property

Public Property Get currentRecordNumber() As Long
    currentRecordNumber = m_numRecords
End Property

Public Function getNextRow(Optional delimiter As String = "|") As String
    ' getNextRow reads the next row of data and returns it
    ' as a text string with each field delimited with
    ' the character designated in delimiter.
    Dim colCounter As Long
    
    
    For colCounter = 1 To m_numColumns
        If colCounter > 1 Then getNextRow = getNextRow & delimiter
        getNextRow = getNextRow & m_dataTable.DataBodyRange.Cells(m_recordPointer, colCounter).value
    Next colCounter
    
 
    m_recordPointer = m_recordPointer + 1
End Function


Public Sub initialize(ByRef dataSource As objDataSource)
    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet
    Dim sheetNames As String
    Dim ctr As Integer
    
    Set xlApp = New Excel.Application
    xlApp.Visible = True
    Set wkb = xlApp.Workbooks.Open(dataSource.path, True, False)
    With wkb
        'string delimited by '-' that collects the
        'names of the sheets in the workbook
        For ctr = 1 To .sheets.Count
            If ctr > 1 Then
                sheetNames = sheetNames & "-" & .sheets(ctr).name
            Else
                sheetNames = .sheets(ctr).name
            End If
        Next ctr
    End With
    wksheets = Split(sheetNames, "-")
    
    Set wks = wkb.sheets(wksheets(0))
    Set m_dataTable = wks.ListObjects("Table1")
    m_recordPointer = 1
    
    
    m_numRecords = m_dataTable.DataBodyRange.Rows.Count
    m_numColumns = m_dataTable.DataBodyRange.Columns.Count
End Sub

Sub closeSource()
     xlApp.Workbooks.Close
End Sub
   

