VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objDataSetDAORecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Implements iDataSet


Private rs As DAO.Recordset

Public Property Get iDataSet_EOF() As Boolean
    EOF = rs.EOF
End Property

Public Property Get iDataSet_c() As Long
    rs.MoveLast
    iDataSet_c = rs.Count
End Property

Public Sub iDataSet_i()
End Sub

Public Sub iDataSet_o(ByVal datasource As String)
    Set rs = CurrentDb.OpenRecordset(datasource)
End Sub

Public Sub iDataSet_n()
    rs.MoveNext
End Sub


