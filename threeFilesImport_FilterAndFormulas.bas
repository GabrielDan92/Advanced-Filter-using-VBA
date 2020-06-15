Sub main()

With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .EnableAnimations = False
    .DisplayAlerts = False
End With
    
'declare the sheets
Dim wbRaw As Worksheet, wbReworked As Worksheet, wbTemp As Worksheet, wbPivots As Worksheet, wb As Workbook
Set wbRaw = Sheet2
Set wbReworked = Sheet4
Set wbPivots = Sheet5
Set wbTemp = Sheet6
Set wbCriteria = Sheet7

'clear the sheets
Dim arr(2) As Variant, arrCopy As Variant
arr(0) = wbRaw.Name
arr(1) = wbReworked.Name
arr(2) = wbTemp.Name
For i = 0 To UBound(arr)
    If ThisWorkbook.Sheets(arr(i)).AutoFilterMode Then ThisWorkbook.Sheets(arr(i)).AutoFilterMode = False
    ThisWorkbook.Sheets(arr(i)).Range("A1").CurrentRegion.ClearFormats
    ThisWorkbook.Sheets(arr(i)).Range("A1").CurrentRegion.ClearContents
Next i
Erase arr
'==========================================================
'open the first file and import the raw data
Call openWb("Original File", wb)
If wb Is Nothing Then GoTo wrapUp
arrCopy = wb.Sheets(1).Range("A1").CurrentRegion
wb.Close (False)
Set wb = Nothing

wbRaw.Range("A1").Resize(UBound(arrCopy, 1), UBound(arrCopy, 2)) = arrCopy
Erase arrCopy
'==========================================================
'open the Consolidated file and import the raw data in the REWORKED sheet
Call openWb("Consolidated File", wb)
If wb Is Nothing Then GoTo wrapUp
With wb.Sheets(1)
    .Range("H:H").NumberFormat = "0"
    .Range("O:O").NumberFormat = "0"
    .Columns("P:P").Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    .Range("P1").Value = "LE-SN-IN"
    lastRow = .Range("A1").CurrentRegion.Rows.Count
    .Range("P2:P" & lastRow).FormulaR1C1 = "=CONCATENATE(RC[-14],""-"",RC[-8],""-"",RC[-1])"
    arrCopy = .Range("A1").CurrentRegion
End With
wb.Close (False)
Set wb = Nothing

With wbReworked
    .Range("A1").Resize(UBound(arrCopy, 1), UBound(arrCopy, 2)) = arrCopy
    .Range("A:O").Delete
    .Range("B:O").Delete
End With
Erase arrCopy
'==========================================================
'open the BI file and import the raw data in the REWORKED sheet
Call openWb("BI File", wb)
If wb Is Nothing Then GoTo wrapUp
With wb.Sheets(1)
    .Columns("P:P").Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    .Range("P1").Value = "Supplier Number - Site Name"
    lastRow = .Range("A1").CurrentRegion.Rows.Count
    .Range("P2:P" & lastRow).FormulaR1C1 = "=CONCATENATE(RC[-15],""-"",RC[-1])"
    arrCopy = .Range("A1").CurrentRegion
End With
wb.Close (False)
Set wb = Nothing

With wbCriteria
    .Range("D1").Resize(UBound(arrCopy, 1), UBound(arrCopy, 2)) = arrCopy
    .Range("D:R").Delete
    .Range("E:U").Delete
    .Range("F:BI").Delete
End With
Erase arrCopy
'==========================================================
'convert columns K and S to the 'Number' format
With wbRaw
    .Range("K:K").NumberFormat = "0"
    .Range("S:S").NumberFormat = "0"
    .Range("B1").Value = "Legal Entity"
End With
'build the advance filter in the 'temp' worksheet
Dim rng As Range
With wbTemp
    .Range("A1").Value = "Legal Entity"
    .Range("B1").Value = "Source_"
End With
With wbRaw
Set rng = .Range("A1").CurrentRegion
arrCopy = .Range(.Cells(1, 1), .Cells(1, rng.Columns.Count))
End With
wbTemp.Range("D1").Resize(UBound(arrCopy, 1), UBound(arrCopy, 2)) = arrCopy
Erase arrCopy
wbTemp.Range("A2").Value = "<>" & wbCriteria.Range("A2").Value
Call selection("B", 2, wbCriteria, wbTemp)

'set the advance filter ranges and run it
Set rngCriteria = wbTemp.Range("A1").CurrentRegion
Set rngOutput = wbTemp.Range("D1").CurrentRegion
Set rngData = wbRaw.Range("A1").CurrentRegion
rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput

'add 4 new columns and populate them with formulas
With wbTemp
    lastRowTemp = .Range("D1").CurrentRegion.Rows.Count
    .Range("AK1").Value = "LE - supplier number - invoice number"
    .Range("AK2:AK" & lastRowTemp).FormulaR1C1 = "=CONCATENATE(RC[-32],""-"",RC[-23],""-"",RC[-15])"
End With
With wbTemp
    .Range("AL1").Value = "Flag in DART"
    .Range("AL2:AL" & lastRowTemp).FormulaR1C1 = "=VLOOKUP(RC[-1],'AP200 month REWORKED'!C[-37]:C[-36],2,0)"
    .Range("AM1").Value = "supplier number - supplier site code"
    .Range("AM2:AM" & lastRowTemp).FormulaR1C1 = "=CONCATENATE(RC[-25],""-"",RC[-21])"
End With
With wbTemp
    .Range("AN1").Value = "Country"
    .Range("AN2:AN" & lastRowTemp).FormulaR1C1 = "=VLOOKUP(RC[-1],criteria!C[-36]:C[-35],2,0)"
    .Columns("R:R").NumberFormat = "@"
    Set rng = .Range("D1").CurrentRegion
    arrCopy = .Range(.Cells(1, 4), .Cells(1, rng.Columns.Count + 3))        'copy the header and save it for reapply
    .Range("AO1").Value = "Formula"
    .Range("AO2:AO" & lastRowTemp).FormulaR1C1 = "=AND(RC[-1]=""Italy"",RC[-36]=""411_OU"",IF(ISNUMBER(SEARCH(""DOMVAT"",RC[-23]))=FALSE,TRUE))"
    .Range("D1:AO" & lastRowTemp).AutoFilter Field:=38, Criteria1:="TRUE"
    .Range("D1").CurrentRegion.Delete shift:=xlUp
    .Rows("1:1").Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    .Range("D1").Resize(UBound(arrCopy, 1), UBound(arrCopy, 2)) = arrCopy
    .Range("AO:AO").Delete
    Erase arrCopy
    arrCopy = .Range("D1").CurrentRegion
    .Range("A1").Value = "a"
    .Range("B1").Value = "b"
    .Range("C1").Value = "c"
    .Range("A1").CurrentRegion = ClearFormats
    .Range("A1").CurrentRegion = ClearContents
End With
'delete the imported data once we no longer need it
With wbReworked
    .Range("A1").CurrentRegion.ClearFormats
    .Range("A1").CurrentRegion.ClearContents
    .Range("A1").Resize(UBound(arrCopy, 1), UBound(arrCopy, 2)) = arrCopy
    Erase arrCopy
End With
With wbCriteria
    .Range("D1").CurrentRegion.ClearFormats
    .Range("D1").CurrentRegion.ClearContents
End With
Dim pivotTable As pivotTable
For Each pivotTable In wbPivots.PivotTables
    pivotTable.ChangePivotCache ThisWorkbook.PivotCaches.Create _
    (SourceType:=xlDatabase, SourceData:=wbReworked.Range("A1").CurrentRegion)  'update the Pivot Source Data Range
    pivotTable.RefreshTable
Next

For Each pivotTable In ThisWorkbook.Sheets("PIVOT").PivotTables
    pivotTable.RefreshTable
Next

wrapUp:
With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
    .EnableAnimations = True
    .DisplayAlerts = True
End With

MsgBox "THE MACRO ROUTINE HAS BEEN COMPLETED!"
    
End Sub
Function openWb(filename As String, ByRef wb As Workbook) As Workbook

With Application.FileDialog(msoFileDialogOpen)
    .Title = "Open the " & filename & " file"
    .AllowMultiSelect = False
    .Show
    If Application.FileDialog(msoFileDialogOpen).SelectedItems.Count <> 0 Then
        file = .SelectedItems(1)
    Else
        Exit Function
    End If
End With
Set wb = Workbooks.Open(file, 2)
If wb.Sheets(1).AutoFilterMode Then wb.Sheets(1).AutoFilterMode = False

End Function
Function remove(ByVal letter As String, ByVal columnIndex As Integer, ByVal wbCriteria As Worksheet, ByVal wbTemp As Worksheet)
    lastRow = wbCriteria.Range(letter & "2", wbCriteria.Range(letter & "2").End(xlDown)).Count
    wbTemp.Range("A2").FormulaR1C1 = "=NOT(ISNUMBER(MATCH('AP200 month'!C[1],'criteria'!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",0)))"
End Function
Function selection(ByVal letter As String, ByVal columnIndex As Integer, ByVal wbCriteria As Worksheet, ByVal wbTemp As Worksheet)
    lastRow = wbCriteria.Range(letter & "2", wbCriteria.Range(letter & "2").End(xlDown)).Count
    wbTemp.Range("B2").FormulaR1C1 = "=COUNTIF('criteria'!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",'AP200 month'!RC[1])"
End Function
