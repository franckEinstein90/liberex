VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iDataSet"
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

Public Sub n() 'next
End Sub

Public Sub o(ByVal datasource As String) 'open
End Sub

Public Property Get c() As Long 'count
End Property


Public Sub i() 'initialize
End Sub

Public Sub x() 'close
End Sub

