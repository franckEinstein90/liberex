VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objDataSetExcelRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDataSet

Private m_wks As Excel.Worksheet
Private m_dataRange As Excel.Range

Private m_lastRow As Long
Private m_firstRow As Long


Private m_currentRow As Long
Private m_rangeRowOffset As Long

Private m_numRecords As Long

Private m_defaultDelimiter

Public Property Get IDataSet_EOF() As Boolean
    If m_currentRow - m_firstRow >= m_numRecords Then
        IDataSet_EOF = True
    Else
        IDataSet_EOF = False
    End If
End Property
 
Public Property Get IDataSet_record() As iRecord
    Set IDataSet_record = New objRangeRecord
End Property
Public Property Get IDataSet_defaultDelimiter() As String
    IDataSet_defaultDelimiter = m_defaultDelimiter
End Property

Public Sub IDataSet_setDefaultDelimiter(ByVal strDelimiting As String)
    m_defaultDelimiter = strDelimiting
End Sub


Public Property Get iDataSet_fieldValue(ByVal fieldName As String) As Variant
    iDataSet_fieldValue = m_wks.Range(fieldName & m_currentRow + m_rangeRowOffset)
End Property


Public Sub IDataSet_moveNext() 'next
    If Not IDataSet_EOF Then m_currentRow = m_currentRow + 1
End Sub
 
Public Sub IDataSet_read(ByVal datasource As String)
    Dim dataRange() As String
    Dim dataSourceArray() As String
    If Len(IDataSet_defaultDelimiter) = 0 Then Call IDataSet_setDefaultDelimiter("\")
    
    dataSourceArray = Split(datasource, IDataSet_defaultDelimiter)
    Set m_wks = Worksheets(dataSourceArray(0))
    Set m_dataRange = m_wks.Range(dataSourceArray(1))
    dataRange = Split(dataSourceArray(1), ":")
    With m_wks
        m_firstRow = .Range(dataRange(0)).Row
        m_lastRow = .Range(dataRange(1)).Row
    End With
    m_numRecords = m_lastRow - m_firstRow + 1
    m_currentRow = m_firstRow
    
End Sub
 
Public Property Get IDataSet_count() As Long
    IDataSet_count = m_numRecords
End Property
 
Public Property Get IDataSet_index() As Long
    IDataSet_index = m_currentRangeRow + m_rangeRowOffset
End Property
 
Public Sub IDataSet_initialize()
    
End Sub
 
Public Sub IDataSet_cl()
'closes the dataset
End Sub
