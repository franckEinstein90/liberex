Option Explicit

Implements IDataSet
Dim fieldDict As Object

Private m_csvRS As ADODB.recordset
Private m_cn as ADODB.connection

public property get IDataSet_EOF()
    IDataSet_EOf = m_csvRS.EOF()
end property

public Sub IDataSet_readDataSource(byval dataSource as IDataSource)
   
    m_cN.Open ("Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & dataSource.Path & ";" & _
                   "Extended Properties=""text; HDR=Yes; FMT=Delimited; IMEX=1;""")
    m_csvRS.ActiveConnection = m_cn
    m_csvRS.Source = "select * from " & dataSource.fileName
end sub

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



