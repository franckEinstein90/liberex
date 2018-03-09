Option Compare Database
Option Explicit

Implements IDataSource

'March 2018
'Object Encapsulate concept of a data source
'-Can be txt, excel, or csv
Private m_filePath As String
Private m_fileName As String
Private m_folderPath As String

Private m_fileExists As Boolean
Public m_extension As String
Public m_file_type As String


Public Property Get IDataSource_fullPath() As String
    path = m_filePath
End Property
Public Property Get IDataSource_fileName() As String
    name = m_fileName
End Property
Public Property Get IDataSource_folderPath() As String
    folderPath = m_folderPath
End Property
Public Property Get IDataSource_exists() As Variant
    exists = m_fileExists
End Property

Private Sub deduce_file_type()
    Dim fileElements() As String
    fileElements = Split(m_fileName, ".")
    m_extension = LCase(fileElements(UBound(fileElements)))
    Select Case m_extension
            Case "txt"
                m_file_type = "text"
            Case "xls", "xlsx"
                m_file_type = "excel"
            Case Else
                m_file_type = "unknown"
    End Select
End Sub

Public Sub IDataSource_initialize(ByVal filePath As String)
    Dim pathElements() As String
       
    'fills in values variables related to path and file name
    m_filePath = filePath
    pathElements = Split(m_filePath, "\")
    m_fileName = pathElements(UBound(pathElements))
    m_folderPath = Left(m_filePath, Len(m_filePath) - Len(m_fileName))
    
    Call deduce_file_type 'extract extension
    'checks the file's existence
    m_fileExists = e(filePath:=m_filePath)
End Sub


Public Sub IDataSource_close()
End Sub

