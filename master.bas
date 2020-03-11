Option Explicit

Sub AdvFilter()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim rng As Range, rngData As Range, rngCriteria As Range, rngOutput As Range
Dim lastRow As Long, i As Long


'==================RETRIEVE THE EMPLOYEE ID FROM THE 'EMP' SHEET BASED ON THE MANAGER ID===================


    'remove the filter
        If ThisWorkbook.Sheets("EMP").AutoFilterMode Then ThisWorkbook.Sheets("EMP").AutoFilterMode = False
    
    'set the data range
        Set rngData = ThisWorkbook.Sheets("EMP").Range("A1").CurrentRegion
        
    'set the criteria range
        Set rngCriteria = ThisWorkbook.Sheets("MAN").Range("A1").CurrentRegion
    
    'set the output range
        Set rngOutput = ThisWorkbook.Sheets("MAN").Range("C1").CurrentRegion
        rngOutput.Offset(1).ClearContents
        rngOutput.Offset(1).ClearFormats
 
    'run the advanced filter
        rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput
    

'=========RETRIEVE THE 'DATA' SHEET RESULTS BASED ON THE EMPLOYEE ID ADDED AS A FILTER IN COLUMN 'B'=========


    'move the retrieved Employee IDs from column 'D' to column 'F'
        Set rng = ThisWorkbook.Sheets("MAN").Range("D1").CurrentRegion
        lastRow = rng.Rows.Count
        
        For i = 2 To lastRow
            ThisWorkbook.Sheets("MAN").Range("F" & i) = ThisWorkbook.Sheets("MAN").Range("D" & i)
        Next i

    'remove the filter
    If ThisWorkbook.Sheets("Data").AutoFilterMode Then ThisWorkbook.Sheets("Data").AutoFilterMode = False

   'set the data range
       Set rngData = ThisWorkbook.Sheets("Data").Range("A1").CurrentRegion
       
   'set the criteria range
       Set rngCriteria = ThisWorkbook.Sheets("MAN").Range("F1").CurrentRegion
   
   'set the output range
       Set rngOutput = ThisWorkbook.Sheets("Results").Range("A1").CurrentRegion
       rngOutput.Offset(1).ClearContents
       rngOutput.Offset(1).ClearFormats

   'run the advanced filter
       rngData.AdvancedFilter xlFilterCopy, rngCriteria, rngOutput


Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
