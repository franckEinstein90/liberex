Option Compare Database
Option Explicit

Private Type TExcelWorkbook 
	workbook As Excel.Workbook
	formulaRanges As Object
	sheets As Object
	values As Object
End Type	
Private this as TExcelWorkbook


Private type_list As String

public Sub initialize(bk As Excel.Workbook)
    Dim wks As Excel.Worksheet
    Set this.workbook = bk
    Set formula_ranges = CreateObject("scripting.dictionary")
    Set sheets = CreateObject("scripting.dictionary")
    Set values = CreateObject("scripting.dictionary")
    
    For Each wks In wkb.Worksheets
        this.sheets.Add key:=wks.name, Item:=0
    Next wks
End Sub

public Sub newSheet( _
    ByVal After As String, _
    ByVal SheetName As String)
    Dim wks As Excel.Worksheet
                    
    with this
       If .sheets.exists(SheetName) Then end sub
        .sheets.Add After:=.workbook.Worksheets(After)
        .sheets.Add key:=SheetName, Item:=0
        Set wks = .workbook.ActiveSheet
        wks.name = SheetName
    end with 	
End Sub

Sub addValue(ByVal Sheet As String, _
                ByVal Range As String, _
                ByVal Val As Variant)
    If values.exists(Range) Then
        values(Range) = Val
    Else
        values.Add key:=Range, Item:=Val
    End If
    wkb.sheets(Sheet).Range(Range).value = Val
End Sub

Function getValue( _
    ByVal range_name As String) As Variant
    get_value = values(range_name)
End Function

Sub addFormula(ByVal sheet_name As String, ByVal range_name As String, ByVal formula As String)
    If formula_ranges.exists(range_name) Then
    Else
        formula_ranges.Add key:=range_name, Item:=formula
        wkb.sheets(sheet_name).Range(range_name) = "=" & formula
    End If
End Sub


Sub clearTable(ByVal sheet_name As String, ByVal table_name As String)
    With wkb.Worksheets(sheet_name).ListObjects(table_name)
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
End Sub


