VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Sub process(ByVal typeName As String, values As Collection)
    Dim i As Long
    For i = 1 To values.count
        Debug.Print values(i)
    Next i
End Sub
