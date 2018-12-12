VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objComputerFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IDataSource


'************************
'** objComputerFile
'** Wrapper for a computer file
'** Implements IDataSource
'** Requires IDataSource.cls

Private Type tComputerFile
    m_filePath As String
    m_fileExists As Boolean
End Type


Private m_fileName As String
Private m_folderPath As String
Public m_extension As String
Public m_file_type As String

Private computerFile As tComputerFile

Private Function initPrivateVars() As Boolean
    With computerFile
        If Dir(.m_filePath) <> "" Then .m_fileExists = True Else .m_fileExists = False
    End With
End Function

Public Property Get IDataSource_fullPath() As String
    path = m_filePath
End Property
Public Property Get IDataSource_fileName() As String
    name = m_fileName
End Property
Public Property Get IDataSource_folderPath() As String
    folderPath = m_folderPath
End Property
Public Property Get IDataSource_exists() As Boolean
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

Public Sub IDataSource_initDataSource(ByVal filePath As String)
    Dim pathElements() As String
    With computerFile
        'fills in values variables related to path and file name
        .m_filePath = filePath
        pathElements = Split(.m_filePath, "\")
        m_fileName = pathElements(UBound(pathElements))
        m_folderPath = Left(.m_filePath, Len(.m_filePath) - Len(m_fileName))
        Call deduce_file_type 'extract extension
        Call initPrivateVars
    End With
End Sub

Public Sub IDataSource_closeDataSource()
End Sub

Public Sub IDataSource_writeToFile(ByVal text As String)
    Dim fn As Integer
    With computerFile
        fn = FreeFile
        Open .m_filePath For Output As #fn
            Print #fn, text
        Close fn
        .m_fileExists = True
    End With
End Sub
