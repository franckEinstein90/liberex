VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'iDataSet is the interface for a variety
'of data readers for different formats
'The data readers are "plugged" in
'a dataSource object (objDataSource)
'which combines them with a datafilestructure
'object and uses that the combination of the two
'to do the migrating and the information
'processing of the concurrently
 
Private m_defaultDelimiter As String

Public Property Get defaultDelimiter()
    defaultDelimiter = m_defaultDelimiter
End Property


Public Property Get fieldValue(ByVal fieldName As String) As Variant
'returns the value of the field at
'the current record
End Property

Public Sub setDefaultDelimiter(ByVal strDelimiting As String)
End Sub

Public Property Get EOF() As Boolean
'Returns true if on last record, false otherwise
End Property
 
Public Sub moveNext() 'next
 
End Sub
 
Public Sub read(ByVal datasource As String)
'open the dataset and postitions the
'file pointer to the first record
End Sub
 
Public Property Get count() As Long
'returns the number of records in the dataset
End Property
 
Public Property Get index() As Long
'returns the position of the file pointer
End Property
 
Public Sub initialize()
'initializes the data set
End Sub
 
Public Sub cl()
'closes the dataset
End Sub
