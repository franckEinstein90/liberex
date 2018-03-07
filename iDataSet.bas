option explicit
'iDataSet is the interface for a variety
'of data readers for different formats
'The data readers are "plugged" in
'a dataSource object (objDataSource)
'which combines them with a datafilestructure
'object and uses that the combination of the two
'to do the migrating and the information
'processing of the concurrently

public Property Get EOF() as boolean
end property

public Sub MoveNext()
end Sub

public Sub open(byval datasource as string)
End Sub

Public Property Get numRecords() as Long
end Property


Public Sub initialize(ByRef dataSource as objDataSource)
end sub

public Sub close()
end sub
