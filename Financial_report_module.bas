Attribute VB_Name = "Module1"
Sub TransferToWordTemplate()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim filePath As String
    
    ' Define the path to your Word template
    filePath = "C:\Path\To\Your\Template.docx"
    
    ' Set up Word application
    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    ' If Word is not running, create a new instance
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    
    ' Open the Word template
    Set wordDoc = wordApp.Documents.Open(filePath)
    
    ' Optional: make Word visible
    wordApp.Visible = True

    ' Assuming your Excel data is in Sheet1
    With ThisWorkbook.Sheets("Sheet1")
        ' Fill in the Word bookmarks with Excel data
        wordDoc.Bookmarks("Bookmark1").Range.Text = .Range("A1").Value
        wordDoc.Bookmarks("Bookmark2").Range.Text = .Range("B1").Value
        wordDoc.Bookmarks("Bookmark3").Range.Text = .Range("C1").Value
    End With
    
    ' Save the Word document
    wordDoc.SaveAs2 "C:\Path\To\Save\Output.docx"
    
    ' Clean up
    wordDoc.Close SaveChanges:=False
    wordApp.Quit
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

Sub AutomateFinancialReporting()
    Dim wsSources As Worksheet
    Dim wsData As Worksheet
    Dim wsMetrics As Worksheet
    
    ' Define sheets for data entry and consolidation
    Set wsSources = ThisWorkbook.Sheets("DataSources")
    Set wsData = ThisWorkbook.Sheets("RawData")
    Set wsMetrics = ThisWorkbook.Sheets("CalculatedMetrics")

    ' Data consolidation from listed file paths
    Dim folderPath As String, filename As String
    Dim lastRow As Long
    Dim currentRow As Long
    currentRow = 2 ' Starting to enter data from row 2

    folderPath = wsSources.Range("A2").Value
    Do While folderPath <> ""
        ' Open the workbook
        Dim wb As Workbook
        Set wb = Workbooks.Open(folderPath)
        
        ' Copy data from Sheet1 to RawData
        wb.Sheets(1).Range("A2:D100").Copy wsData.Cells(currentRow, 1)

        ' Increment currentRow for next data entry
        currentRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1
        
        ' Close the workbook
        wb.Close False
        
        ' Move to the next file path
        folderPath = wsSources.Cells(currentRow, 1).Value
    Loop

    ' Metrics Calculation
    CalculateMetrics wsData, wsMetrics

    ' Formatting the summary report
    FormatSummaryReport wsMetrics

    ' Create summary charts
    CreateCharts wsMetrics
    
    ' Finish up
    MsgBox "Monthly Financial Report automation complete!"
End Sub

Sub CalculateMetrics(wsData As Worksheet, wsMetrics As Worksheet)
    ' Example function to calculate financial metrics from RawData and output to CalculatedMetrics sheet
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Add headers for calculated metrics
    wsMetrics.Range("A1").Value = "Net Profit"
    wsMetrics.Range("B1").Value = "Revenue Growth (%)"
    wsMetrics.Range("C1").Value = "Profit Margin (%)"
    wsMetrics.Range("D1").Value = "Total Revenue"
    
    ' Loop to calculate net profit, revenue growth, profit margin, and total revenue
    Dim i As Long
    For i = 2 To lastRow
        ' Net Profit
        wsMetrics.Cells(i, 1).Value = wsData.Cells(i, 3).Value - wsData.Cells(i, 4).Value ' Revenue - Expenses
        
        ' Revenue Growth (%) - Assuming data is ordered by month
        If i > 2 Then
            wsMetrics.Cells(i, 2).Value = (wsData.Cells(i, 3).Value - wsData.Cells(i - 1, 3).Value) / wsData.Cells(i - 1, 3).Value * 100
        Else
            wsMetrics.Cells(i, 2).Value = "N/A" ' No previous data to calculate growth
        End If
        
        ' Profit Margin (%)
        If wsData.Cells(i, 3).Value <> 0 Then
            wsMetrics.Cells(i, 3).Value = (wsMetrics.Cells(i, 1).Value / wsData.Cells(i, 3).Value) * 100
        Else
            wsMetrics.Cells(i, 3).Value = "N/A"
        End If
        
        ' Total Revenue
        wsMetrics.Cells(i, 4).Value = wsData.Cells(i, 3).Value
    Next i
End Sub

Sub FormatSummaryReport(wsMetrics As Worksheet)
    ' Format the summary report in the CalculatedMetrics sheet
    With wsMetrics
        .Range("A1:D1").Font.Bold = True
        .Columns("A:D").AutoFit
        ' Add table formatting
        .ListObjects.Add(xlSrcRange, .Range("A1:D100"), , xlYes).TableStyle = "TableStyleMedium9"
    End With
End Sub

Sub CreateCharts(wsMetrics As Worksheet)
    ' Create a sample chart from the metrics
    Dim chart As ChartObject
    Set chart = wsMetrics.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    With chart.chart
        .SetSourceData Source:=wsMetrics.Range("A1:B10") ' Example range for chart data
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Monthly Revenue Growth"
    End With
End Sub

