VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'iDataSet is the interface for a variety
'of data readers for different formats
'The data readers are "plugged" in
'a dataSource object (objDataSource)
'which combines them with a datafilestructure
'object and uses that the combination of the two
'to do the migrating and the information
'processing of the concurrently

Public Property Get EOF() As Boolean
End Property

Public Sub moveNext() 
End Sub

Public Sub openForReading( _
    dataSource as Idatasource, _
    Optional dataStructure as IDataStructure)

end sub

Public Sub read(_
    datasource As IDataSource, _
    Optional dataStructure as IDataStructure) 

End Sub

Public Property Get count() As Long 
End Property


Public Sub initialize() 'initialize
End Sub

Public Sub close() 'close
End Sub

