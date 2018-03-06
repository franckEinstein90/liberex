Option Explicit


Dim fieldDict As Object
Dim csvRS As ADODB.recordset

Property Get recordset() As ADODB.recordset
    Set recordset = csvRS
End Property

Public Sub initialize(ByRef dataSource As objDataSource)
    Dim strTempCopyPath As String
    Dim key As Variant
    strTempCopyPath = copyFileToTempLocation(csvAttachment)
    
    Set fieldDict = CreateObject("Scripting.Dictionary")
    Set csvRS = extractEmailAttachementInformation(strTempCopyPath, fieldDict)
    
    With csvRS
        .MoveFirst
        While Not .EOF
            For Each key In fieldDict.keys
                Debug.Print .fields(key)
            Next key
            .MoveNext
        Wend
    End With
End Sub


Private Function extractEmailAttachementInformation(ByVal filePath As String, fieldDict As Object) As ADODB.recordset
  
    Dim directory As String
    Dim fileName As String
   
    directory = Left(filePath, InStrRev(filePath, "\"))
    fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
   
    Set extractEmailAttachementInformation = readCSVInformation(directory, fileName, fieldDict)
End Function



