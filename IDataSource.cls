VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Interface IDataSource

Public Property Get fullPath() As String
End Property

Public Property Get folderPath() As String
End Property

Public Property Get exists() As Boolean
End Property

Public Property Get fileName() As String
End Property

Public Sub writeToFile(ByVal text As String)

End Sub
Public Sub initDataSource(ByVal path As String)
End Sub

Public Sub closeDataSource()
End Sub

