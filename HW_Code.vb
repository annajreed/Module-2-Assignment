
Sub module2()
For Each Worksheet In Worksheets

Dim lastrow As Long
lastrow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
Dim ticker As String
Dim change As Double
Dim total As Double
Dim percent As Double
Dim start As Integer
Dim summary As Double

change = 0
total = 0
summary = 2

Worksheet.Cells(1, 9).Value = "Ticker"
Worksheet.Cells(1, 10).Value = "Yearly Change"
Worksheet.Cells(1, 11).Value = "Percent Change"
Worksheet.Cells(1, 12).Value = "Total Stock Value"


For i = 2 To lastrow
If Worksheet.Cells(i - 1, 1).Value <> Worksheet.Cells(i, 1).Value Then
    start = Worksheet.Cells(i, 3).Value
End If

If Worksheet.Cells(i + 1, 1).Value <> Worksheet.Cells(i, 1).Value Then
    
    ticker = Worksheet.Cells(i, 1).Value
    total = total + Worksheet.Cells(i, 7).Value
    Worksheet.Cells(summary, 12) = total
    change = Worksheet.Cells(i, 6).Value - start
    Worksheet.Cells(summary, 10).Value = change
    percent = round((change / start)*100, 2)
    
    
    
    Worksheet.Range("K" & summary).Value = percent
    Worksheet.Range("I" & summary).Value = ticker
    Worksheet.Range("L" & summary).Value = total
    Worksheet.Range("J" & summary).Value = change
    
    
    total = 0
    summary = summary + 1
     
Else
    total = total + Worksheet.Cells(i, 7).Value
    
    
End If
If Worksheet.Cells(i, 10).Value > 0 Then
    Worksheet.Cells(i, 10).Interior.ColorIndex = 4
    
    ElseIf Worksheet.Cells(i, 10).Value < 0 Then
    Worksheet.Cells(i, 10).Interior.ColorIndex = 3
    
    ElseIf Worksheet.Cells(i, 10).Value = 0 Then
    Worksheet.Cells(i, 10).Interior.ColorIndex = 6
   
End If

Next i
Next Worksheet

Worksheets("2018").Cells.EntireColumn.AutoFit
Worksheets("2019").Cells.EntireColumn.AutoFit
Worksheets("2020").Cells.EntireColumn.AutoFit


End Sub



'percent format from excelcampus.com 
'help finding "start" from AskBCS 
'help with summary table formatting from AskBCS
'tutoring session for help with percentage column/error "dividing by 0"
'worksheet loop help from excel-easy.com
