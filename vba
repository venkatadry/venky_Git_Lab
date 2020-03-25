

Sub IPMSReport()
    Call CheckWorkSheetExists
    Call ClearSheet
    Call getDataFromWbs
    Call findipmsinfo
    
End Sub


Sub getDataFromWbs()

Dim wb As Workbook, ws As Worksheet
Set fso = CreateObject("Scripting.FileSystemObject")
'This is where you put YOUR folder name
Set fldr = fso.GetFolder("C:\Users\Radha\Downloads\Temp\Temp")
'Set fldr = fso.GetFolder("D:\Temp\MM")

'Next available Row on Master Workbook
y = ThisWorkbook.Sheets("Consolidated_Won_Details").Cells(Rows.Count, 1).End(xlUp).Row + 1

'Loop through each file in that folder
For Each wbFile In fldr.Files
    
    'Make sure looping only through files ending in .xlsx (Excel files)
    If fso.GetExtensionName(wbFile.Name) = "xls" Then
      
      'Open current book
      Set wb = Workbooks.Open(wbFile.Path)
      
      'Loop through each sheet (ws)
      For Each ws In wb.Sheets
          'Last row in that sheet (ws)
          wsLR = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
          'Loop through each record (row 2 through last row)
          For x = 11 To wsLR
            'Put column 1,2,3 and 4 of current sheet (ws) into row y of master sheet, then increase row y to next row
            ThisWorkbook.Sheets("Consolidated_Won_Details").Cells(y, 1) = ws.Cells(x, 5) 'col 1
            ThisWorkbook.Sheets("Consolidated_Won_Details").Cells(y, 2) = ws.Cells(x, 1) 'col 1
            ThisWorkbook.Sheets("Consolidated_Won_Details").Cells(y, 3) = ws.Cells(x, 6) 'col 1
            ThisWorkbook.Sheets("Consolidated_Won_Details").Cells(y, 4) = ws.Cells(x, 7) 'col 1
            ThisWorkbook.Sheets("Consolidated_Won_Details").Cells(y, 5) = ws.Cells(x, 8) 'col 1
            y = y + 1
          Next x
          
          
      Next ws
      
      'Close current book
      wb.Close savechanges:=False
    End If

Next wbFile

End Sub







Sub findipmsinfo()
Dim rng As Range
Dim account As String
Dim rownumber As Long
Dim i As Integer
For i = 2 To 200
EMPID = Sheet1.Cells(i, 1)
Set rng = Sheet2.Columns("A:A").Find(What:=EMPID, _
    LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    rownumber = rng.Row
    Sheet1.Cells(i, 2).Value = Sheet2.Cells(rownumber, 2).Value
    Sheet1.Cells(i, 3).Value = Sheet2.Cells(rownumber, 3).Value
    Sheet1.Cells(i, 4).Value = Sheet2.Cells(rownumber, 4).Value
    Sheet1.Cells(i, 5).Value = Sheet2.Cells(rownumber, 5).Value
 
Next i
End Sub


'https://dedicatedexcel.com/vba-clear-entire-sheet-in-excel/







'Sub ClearSheet()

'Sheets("Sheet2").Delete

'End Sub


Sub ClearSheet()

Sheets("Sheet2").Cells.Delete

End Sub


Sub Report()
    Call CheckWorkSheetExists
    Call ClearSheet
    Call getDataFromWbs
    Call findipmsinfo
    
End Sub





Sub CheckWorkSheetExists()
    For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Consolidated_Won_Details" Then
        exists = True
    End If
Next i

If Not exists Then

    'Worksheets.Add.Name = "Consolidated_Won_Details"
    'mainWB.Sheets.Add(After:=mainWB.Sheets(mainWB.Sheets.Count)).Name = Consolidated_Won_Details
    
    Sheets.Add.Name = "Consolidated_Won_Details"
Worksheets("Consolidated_Won_Details").Move After:=Worksheets(Worksheets.Count)
End If
End Sub




Assuming your tables start with headers in respective sheets, you can simply enter following SUMIFS function in "B2" and copy it down

=SUMIFS(Sheet2!C:C,Sheet2!A:A,Sheet1!A2,Sheet2!B:B,"<>Nonbillable")

But since you are asking VBA ..

Sub Macro2()

    Sheets("Sheet1").Range("B2:B4").FormulaR1C1 = _
        "=SUMIFS(Sheet2!C3:C3,Sheet2!C1:C1,RC[-1],Sheet2!C2:C2,""<>Nonbillable"")"
    Worksheets("Sheet1").Calculate
    'you can add the following line to convert the formula to value
    Sheets("Sheet1").Range("B2:B4").Value = Sheets("Sheet1").Range("B2:B4").Value

End Sub
