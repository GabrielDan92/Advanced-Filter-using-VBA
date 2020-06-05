Public raw As Worksheet, dest As Worksheet, criteria As Worksheet, semiFinal As Worksheet, final As Worksheet, fsRaw As Worksheet
Public rng As Range, rngData As Range, rngCriteria As Range, rngOutput As Range
Public arr As Variant

Sub AdvFilter()

Dim lastRow As Long, lastRowFinal As Long, i As Long
Dim wb As Workbook, pass As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
pass = "protectedSheet"


    Set criteria = Sheet1                                                   'set the criteria sheet reference
    Set raw = Sheet2                                                        'set the raw data sheet reference
    Set dest = Sheet3                                                       'set the destination sheet reference
    Set semiFinal = Sheet4                                                  'set the semiFinal sheet reference
    Set fsRaw = Sheet5                                                      'set the FS raw data sheet
    Set final = Sheet6                                                      'set the final sheet reference
    final.Unprotect pass                                                    'unprotect the sheet
    
'clear previous values from the raw and destination sheets
    raw.Range("A1").CurrentRegion.ClearFormats
    raw.Range("A1").CurrentRegion.ClearContents
    fsRaw.Range("A1").CurrentRegion.ClearFormats
    fsRaw.Range("A1").CurrentRegion.ClearContents
    semiFinal.Range("A1").CurrentRegion.ClearFormats
    semiFinal.Range("A1").CurrentRegion.ClearContents
    If final.AutoFilterMode Then final.AutoFilterMode = False
    final.Range("A1").CurrentRegion.ClearFormats
    final.Range("A1").CurrentRegion.ClearContents
    
    Set rngCriteria = dest.Range("A1").CurrentRegion                        'set the criteria range
    Set rngOutput = dest.Range("G1").CurrentRegion                          'set the output range

'==================Import the raw data===================

    'open the first file
        With Application.FileDialog(msoFileDialogOpen)
            .Title = "Open the first file"
            .AllowMultiSelect = False
            .Show
            file = .SelectedItems(1)
        End With
        Set wb = Workbooks.Open(file, 2)
        
    'copy the raw data in the raw sheet
        If wb.Sheets(1).AutoFilterMode Then wb.Sheets(1).AutoFilterMode = False
        Set rng = wb.Sheets(1).Range("A1").CurrentRegion
        arr = rng
        raw.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
        Erase arr
        wb.Close (False)
        Set wb = Nothing
        
    'open the second file
        With Application.FileDialog(msoFileDialogOpen)
            .Title = "Open the second file"
            .AllowMultiSelect = False
            .Show
            file = .SelectedItems(1)
        End With
        Set wb = Workbooks.Open(file, 2)
        
    'copy the raw data in the raw sheet
        If wb.Sheets(1).AutoFilterMode Then wb.Sheets(1).AutoFilterMode = False
        Set rng = wb.Sheets(1).Range("A1").CurrentRegion
        arr = rng
        fsRaw.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
        Erase arr
        wb.Close (False)
        Set wb = Nothing

'==================Add the criteria filters===================
        
    Call clear(0)                                                           'clear the data range, criteria range, output range
    With dest
        Set rng = raw.Range("A1:CK1")
        arr = rng
        .Range("G1").Resize(UBound(arr, 1), UBound(arr, 2)) = arr           'copy the header from the BBB raw sheet
        Erase arr
        .Range("A2").Value2 = "="                                           'add the Supplier Inactive Date criteria (select blanks)
        Call supplierType("remove", "B", 2)                                 'add the Supplier Type criteria (remove)
        .Range("C2").Value2 = criteria.Range("C2") & "*"                    'add the Tax Registration Number criteria (select)
        Set rngCriteria = .Range("A1").CurrentRegion                        'set the criteria range
        Set rngOutput = .Range("G1").CurrentRegion                          'set the output range
        Set rngData = raw.Range("A1").CurrentRegion                         'set the data range
    End With
    
    rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput             'run the advanced filter
    lastRowFinal = copyData(1, semiFinal)                                   'copy the results in the semiFinal sheet
        
    With semiFinal                                                          'insert a new column in Q and concatenate the SUPPLIER NUMBER, SUPPLIER NAME, OPERATING UNIT values
        Set rng = .Range("A1").CurrentRegion
        lastRow = rng.Rows.Count
        .Columns("Q:Q").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("Q1").FormulaR1C1 = "Concatenate"
        .Range("Q2:Q" & lastRow).FormulaR1C1 = "=CONCATENATE(RC[-16],RC[-15],RC[-1])"
        Set rng = .Range("A1").CurrentRegion
        rng.RemoveDuplicates Columns:=17, Header:=xlYes                     'remove duplicates
        Set rng = .Range("A1").CurrentRegion
        lastRow = rng.Rows.Count
        .Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("D1").Value = "FS HS"                                        'add a new column in D and add the vlookup formula
        .Range("D2:D" & lastRow).FormulaR1C1 = "=VLOOKUP(RC[-3],'FS raw data'!C[-3]:C[-1],3,0)"
        Set rngData = .Range("A1").CurrentRegion
    End With
        
'==================Add the final criteria filters===================
        
    Call clear(0)                                                           'clear the data range, criteria range, output range
    With dest
        Set rng = semiFinal.Range("A1:CM1")
        arr = rng
        .Range("G1").Resize(UBound(arr, 1), UBound(arr, 2)) = arr           'copy the header from the semiFinal sheet
        Erase arr
        .Range("D2").Value = "#N/A"                                         'add the blanks FS HS criteria
        .Range("E2").Value = criteria.Range("G2")                           'add the country criteria
        Set rngCriteria = .Range("A1").CurrentRegion                        'set the criteria range
        Set rngOutput = .Range("G1").CurrentRegion                          'set the output range
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput         'run the advanced filter
        lastRowFinal = copyData(1, final)                                   'copy the results in the semiFinal sheet
    End With

'==================================================

    final.Activate
    final.Protect pass                                                      'protect the sheet

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Function supplierType(ByVal objective As String, ByVal letter As String, ByVal columnIndex As Integer)

    If objective = "select" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("B2").FormulaR1C1 = "=COUNTIF(criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",'raw data'!RC[6])"
    ElseIf objective = "remove" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("B2").FormulaR1C1 = "=NOT(ISNUMBER(MATCH('BBB raw data'!RC[6],criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",0)))"
    End If

End Function

Function copyData(ByVal lastRowFinal As Long, ByVal sheet As Worksheet) As Long

    'copy the results in the sheet passed as argument
        Set rng = dest.Range("G1").CurrentRegion
        arr = rng
        sheet.Range("A" & lastRowFinal).Resize(UBound(arr, 1), UBound(arr, 2)) = arr
        Erase arr
    'find the lastRow of the pasted data
        Set rng = sheet.Range("A1").CurrentRegion
        copyData = rng.Rows.Count
        
End Function

Function clear(ByVal i As Integer)

        rngCriteria.CurrentRegion.Offset(1).ClearFormats
        rngCriteria.CurrentRegion.Offset(1).ClearContents
        rngOutput.CurrentRegion.Offset(i).ClearFormats
        rngOutput.CurrentRegion.Offset(i).ClearContents
        
End Function
