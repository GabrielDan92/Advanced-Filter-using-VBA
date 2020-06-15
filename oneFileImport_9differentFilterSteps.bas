Public raw As Worksheet, dest As Worksheet, criteria As Worksheet, final As Worksheet
Public rng As Range, rngData As Range, rngCriteria As Range, rngOutput As Range

Sub AdvFilter()

Dim lastRow As Long, lastRowFinal As Long, i As Long
Dim arr As Variant, wb As Workbook, pass As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

pass = "protectedSheet"

'unprotect the sheets
    With ThisWorkbook
        .Sheets("destination").Unprotect pass
        .Sheets("final").Unprotect pass
        .Sheets("411-all").Unprotect pass
    End With

'set the raw data sheet reference
    Set raw = Sheet2
'set the destination sheet reference
    Set dest = Sheet3
'set the criteria sheet reference
    Set criteria = Sheet1
'set the final sheet reference
    Set final = Sheet4
'set the criteria range
    Set rngCriteria = dest.Range("A1").CurrentRegion
'set the output range
    Set rngOutput = dest.Range("Q1").CurrentRegion
'remove the filter from the raw sheet and clear previous values
    If raw.AutoFilterMode Then raw.AutoFilterMode = False
    Set rng = raw.Range("A1").CurrentRegion
    rng.Offset(1).ClearContents
'set the final sheet and clear previous values
    Set rng = final.Range("A1").CurrentRegion
    rng.Offset(1).ClearContents
    ThisWorkbook.Sheets("411-all").Range("A1").CurrentRegion.Offset(1).ClearContents
    
lastRowFinal = 1


'=============

    'open the source file
        With Application.FileDialog(msoFileDialogOpen)
            .Title = "Open the Source file"
            .AllowMultiSelect = False
            .Show
            file = .SelectedItems(1)
        End With
        Set wb = Workbooks.Open(file, 2, , , "password")
        
    'copy the raw data in the raw sheet
        Set rng = wb.Sheets(1).Range("A1").CurrentRegion.Offset(1)
        arr = rng
        raw.Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
        Erase arr
        
        wb.Close (False)
        Set wb = Nothing
        
    'set the data range
        Set rngData = raw.Range("A1").CurrentRegion

'=============

'==================Step 1: 404===================

     'clear the data range, criteria range, output range
        Call clear
    'add the company criteria (select)
        Call company("A", 1)
    'add the Category criteria (remove)
        Call category("B", True)
    'add the Buyer Name On PO criteria (remove)
        Call buyerNameOnPO("remove", "D", 4)
    'add the Hold Name criteria (select)
        Call holdName("C", 3)
        
    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    'copy the results in the final sheet
        lastRowFinal = copyData(lastRowFinal)
            
'==================Step 2: 404===================

    'clear the data range, criteria range, output range
        Call clear
    'add the company criteria (select)
        Call company("F", 6)
    'add the Category criteria (remove)
        Call category("G", True)
    'add the Hold Name criteria (select)
        Call holdName("H", 8)
    'add the Item Number criteria (remove, with wildcards)
        Call itemNumber("I", 9, True)
    'add the Buyer Name On PO criteria (select)
        Call buyerNameOnPO("select", "J", 10)
        
    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    'copy the data in the final sheet
        lastRowFinal = copyData(lastRowFinal)
        
'==================Step 3: 405===================

    'clear the data range, criteria range, output range
        Call clear
    'add the company criteria (select)
        Call company("L", 12)
    'add the Category criteria (remove)
        Call category("M", True)
    'add the Hold Name criteria (select)
        Call holdName("N", 14)
    'add the Item Number criteria (remove blanks)
        dest.Range("G1").Value = "ITEM NUMBER"
        dest.Range("G2").Value = "<>"
    'add the Buyer Name On PO criteria (remove)
        Call buyerNameOnPO("remove", "P", 16)
        
    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    'copy the data in the final sheet
        lastRowFinal = copyData(lastRowFinal)
        
'==================Step 4: 408===================

    'clear the data range, criteria range, output range
        Call clear
    'add the company criteria (select)
        Call company("R", 18)
    'add the Category criteria (remove + remove blanks)
        Call category("S", False)
    'add the Hold Name criteria (select)
        Call holdName("T", 20)
    'add the Invoice Source criteria (remove)
        Call invoiceSource("remove", "U", 21)
    'add the Buyer Name On PO criteria (remove)
        Call buyerNameOnPO("remove", "V", 22)

    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    'copy the data in the final sheet
        lastRowFinal = copyData(lastRowFinal)
        
'==================Step 5: 409===================

    'clear the data range, criteria range, output range
        Call clear
    'add the company criteria (select)
        Call company("X", 24)
    'add the Hold Name criteria (select)
        Call holdName("Y", 25)
    'add the Buyer Name On PO criteria (select)
        Call buyerNameOnPO("select", "Z", 26)

    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    'copy the data in the final sheet
        lastRowFinal = copyData(lastRowFinal)

'==================Step 6: 409===================

    'clear the data range, criteria range, output range
        Call clear
    'add the company criteria (select)
        Call company("AB", 28)
    'add the Category criteria (remove)
        Call category("AC", True)
    'add the Hold Name criteria (select)
        Call holdName("AD", 30)
    'add the PO Number criteria (select, range)
        dest.Range("J2").Value = "<" & Left(Str(criteria.Range("AE2").Value), 4) & "1000000"
        dest.Range("K2").Value = ">=" & Str(criteria.Range("AE2").Value) & "000000"
    'add the Item Number criteria (remove blanks)
        dest.Range("G1").Value = "ITEM NUMBER"
        dest.Range("G2").Value = "<>"
    'add the Buyer on Supplier Site criteria (remove)
        Call buyerOnSupplierSite("remove", "AG", 33)
    'add the SupplierName criteria (remove)
        Call supplierName("remove", "AH", 34)
    'add the Ship to Org Name criteria (remove)
        Call shipToOrgName("remove", "AI", 35)
    'add the Buyer Name On PO criteria (remove)
        Call buyerNameOnPO("remove", "AJ", 36)

    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    'copy the data in the final sheet
        lastRowFinal = copyData(lastRowFinal)

'==================Step 7: 410===================

    'clear the data range, criteria range, output range
        Call clear
    'add the company criteria (select)
        Call company("AL", 38)
    'add the Category criteria (remove + remove blanks)
        Call category("AM", False)
     'add the Hold Name criteria (select)
        Call holdName("AN", 40)
     'add the PO Number criteria (remove blanks)
        dest.Range("J2").Value = "<>"
    'add the Buyer on Supplier Site criteria (remove + remove blanks)
        Call buyerNameOnPO("remove", "AP", 42)
        dest.Range("O2").Value = "<>"
        
    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    'copy the data in the final sheet
        lastRowFinal = copyData(lastRowFinal)
        
'==================Step 8: 411===================
        
    'clear the data range, criteria range, output range
        Call clear
    'add the company criteria (select)
        Call company("AR", 44)
    'add the Category criteria (remove)
        Call category("AS", True)
    'add the Buyer Name On PO criteria (remove)
        Call buyerNameOnPO("remove", "AT", 46)
     'add the Hold Name criteria (select)
        Call holdName("AU", 47)
        
    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    'copy the results in sheet '411-all'
        Set rng = dest.Range("Q1").CurrentRegion.Offset(1)
        arr = rng
        ThisWorkbook.Sheets("411-all").Range("A2").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
        Erase arr
    'copy the data in the final sheet
        lastRowFinal = copyData(lastRowFinal)
        
        
 '==================Step 9: 411===================
 
    'clear the data range, criteria range, output range
        Call clear
    'add the company criteria (select)
        Call company("AY", 51)
    'add the Category criteria (remove)
        Call category("AZ", True)
    'add the Buyer Name On PO criteria (remove)
        Call buyerNameOnPO("remove", "BA", 53)
    'add the Hold Name criteria (select)
        Call holdName("BB", 54)
        
    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    'copy the data in the final sheet
        lastRowFinal = copyData(lastRowFinal)
        
    final.Activate

'protect the sheet
    With ThisWorkbook
        .Sheets("destination").Protect pass
        .Sheets("final").Protect pass
        .Sheets("411-all").Protect pass
    End With


Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Function copyData(ByVal lastRowFinal As Long) As Long

    'copy the results in the final sheet
        Set rng = dest.Range("Q1").CurrentRegion.Offset(1)
        arr = rng
        final.Range("A" & lastRowFinal + 1).Resize(UBound(arr, 1), UBound(arr, 2)) = arr
        Erase arr
    'find the lastRow of the pasted data
        Set rng = final.Range("A1").CurrentRegion
        copyData = rng.Rows.Count
        
End Function

Function clear()

        rngCriteria.CurrentRegion.Offset(1).ClearContents
        rngCriteria.CurrentRegion.Offset(1).ClearFormats
        rngOutput.CurrentRegion.Offset(1).ClearContents
        rngOutput.CurrentRegion.Offset(1).ClearFormats
        
End Function

Function category(ByVal letter As String, ByVal blanks As Boolean)

    'add the Category criteria (remove)
        dest.Range("A2").Value = "<>" & criteria.Range(letter & "2").Value & "*"
        
        If blanks = False Then
            dest.Range("B2").Value = "<>"
        End If
        
End Function

Function company(ByVal letter As String, ByVal columnIndex As Integer)

    'lastrow
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
    'add the Company criteria (select)
        dest.Range("E2").FormulaR1C1 = "=COUNTIF(criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",'raw data'!RC[-4])"

End Function

Function holdName(ByVal letter As String, ByVal columnIndex As Integer)

    'lastrow
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
    'add the Hold Name criteria (select)
        dest.Range("D2").FormulaR1C1 = "=COUNTIF(criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",'raw data'!RC[36])"

End Function

Function buyerNameOnPO(ByVal objective As String, ByVal letter As String, ByVal columnIndex As Integer)

    If objective = "select" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("C2").FormulaR1C1 = "=COUNTIF(criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",'raw data'!RC[11])"
    ElseIf objective = "remove" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("C2").FormulaR1C1 = "=NOT(ISNUMBER(MATCH('raw data'!RC[11],criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",0)))"
    End If

End Function

Function itemNumber(ByVal letter As String, ByVal columnIndex As Integer, ByVal wildcard As Boolean)

    If wildcard = True Then
        dest.Range("G1").Value = "ITEM NUMBER"
        dest.Range("G2").Value = "<>" & criteria.Range(letter & "2").Value & "*"
        dest.Range("H1").Value = "ITEM NUMBER"
        dest.Range("H2").Value = "<>" & criteria.Range(letter & "3").Value & "*"
    Else
        dest.Range("G1").Value = "ITEM_NUMBER"
        'lastrow
            lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        'add the Item Number criteria (remove)
            dest.Range("F2").FormulaR1C1 = "=NOT(ISNUMBER(MATCH('raw data'!RC[6],criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",0)))"
    End If

End Function

Function invoiceSource(ByVal objective As String, ByVal letter As String, ByVal columnIndex As Integer)

    If objective = "select" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("F2").FormulaR1C1 = "=COUNTIF(criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",'raw data'!RC[11])"
    ElseIf objective = "remove" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("F2").FormulaR1C1 = "=NOT(ISNUMBER(MATCH('raw data'!RC[11],criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",0)))"
    End If

End Function

Function buyerOnSupplierSite(ByVal objective As String, ByVal letter As String, ByVal columnIndex As Integer)

    If objective = "select" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("L2").FormulaR1C1 = "=COUNTIF(criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",'raw data'!RC[3])"
    ElseIf objective = "remove" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("L2").FormulaR1C1 = "=NOT(ISNUMBER(MATCH('raw data'!RC[3],criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",0)))"
    End If

End Function

Function supplierName(ByVal objective As String, ByVal letter As String, ByVal columnIndex As Integer)

    If objective = "select" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("I2").FormulaR1C1 = "=COUNTIF(criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",'raw data'!RC[-7])"
    ElseIf objective = "remove" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("I2").FormulaR1C1 = "=NOT(ISNUMBER(MATCH('raw data'!RC[-7],criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",0)))"
    End If

End Function

Function shipToOrgName(ByVal objective As String, ByVal letter As String, ByVal columnIndex As Integer)

    If objective = "select" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("M2").FormulaR1C1 = "=COUNTIF(criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",'raw data'!RC[37])"
    ElseIf objective = "remove" Then
        lastRow = criteria.Range(letter & "2", criteria.Range(letter & "2").End(xlDown)).Count
        dest.Range("M2").FormulaR1C1 = "=NOT(ISNUMBER(MATCH('raw data'!RC[37],criteria!R2C" & columnIndex & ":R" & 2 + (lastRow - 1) & "C" & columnIndex & ",0)))"
    End If

End Function
