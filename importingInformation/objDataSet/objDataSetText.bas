Option Compare Database
Option Explicit


Private Const maxLines As Integer = 5
Private m_fileNo As Integer
Private m_recordPointer As Long


Public Property Get currentRecordNumber() As Long
    currentRecordNumber = m_numRecords
End Property

Public Property Get noMoreRecords() As Boolean
   noMoreRecords = EOF(m_fileNo)
End Property

Sub openSourceForReading(ByRef dataSource As objDataSource)
    '1. opens the text file for reading
    '2. moves file pointer to first record
    
    Dim rowStr As String
    m_fileNo = FreeFile 'Get first free file number
    Open dataSource.path For Input As #m_fileNo
    
    'look for the file header
    Line Input #m_fileNo, rowStr
    Do Until dataSource.isFileHeader(rowStr, vbTab)
        Line Input #m_fileNo, rowStr
    Loop
    m_recordPointer = 1
End Sub

Public Function getNextTabDelimitedString() As String
    Line Input #m_fileNo, getNextTabDelimitedString
    m_recordPointer = m_recordPointer + 1
End Function
