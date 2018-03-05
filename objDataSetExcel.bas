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


'************************************************
'Reading from an excel file

Private m_dataTable As Excel.ListObject
    

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

Public Sub getNextRow( _
    ByRef recordContainer As structInputFileInfo, _
    ByRef colsTemplate As Scripting.dictionary)
    
    'getNextRow reads the next row of data
    'and stores it in the recordContainer
    'as specified by the colsTemplate structure
    Dim colIdx As Variant
    Dim colName As String
    
    
        
    With recordContainer
        .Clear
        For Each colIdx In colsTemplate
            colName = colsTemplate(colIdx)
            Select Case colName
                Case "Name"
                    Call .value("FileName", m_dataTable.DataBodyRange.Cells(m_recordPointer, colIdx).value)
                Case "Path"
                    Call .value("Path", m_dataTable.DataBodyRange.Cells(m_recordPointer, colIdx).value)
                Case "Size"
                    Call .value("Size", m_dataTable.DataBodyRange.Cells(m_recordPointer, colIdx).value)
                End Select
        Next colIdx
    End With
    
    m_recordPointer = m_recordPointer + 1
End Sub


Public Sub initialize(ByRef dataSource As objDataSource)
    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet
    Dim sheetNames As String
    Dim ctr As Integer
    
    Set xlApp = New Excel.Application
    xlApp.Visible = True
    Set wkb = xlApp.Workbooks.Open(sourceFile.path, True, False)
    With wkb
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
End Sub

Sub closeSource()
     xlApp.Workbooks.Close
End Sub
   

