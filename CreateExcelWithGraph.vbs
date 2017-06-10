Dim objXL
Set objXL = CreateObject("Excel.Application")

objXL.Visible = TRUE

objXL.WorkBooks.Add

objXL.Columns(1).ColumnWidth = 20
objXL.Columns(2).ColumnWidth = 30

objXL.Cells(1, 1).Value = "Month"
objXL.Cells(1, 2).Value = "Value"

objXL.Range("A1:B1").Select
objXL.Selection.Font.Bold = True
objXL.Selection.Interior.ColorIndex = 1
objXL.Selection.Interior.Pattern = 1 'xlSolid
objXL.Selection.Font.ColorIndex = 2

objXL.Cells(2, 1).Value = "Jan"
objXL.Cells(3, 1).Value = "Feb"
objXL.Cells(4, 1).Value = "Mar"
objXL.Cells(5, 1).Value = "Apr"
objXL.Cells(6, 1).Value = "May"
objXL.Cells(7, 1).Value = "Jun"
objXL.Cells(8, 1).Value = "Jul"
objXL.Cells(9, 1).Value = "Aug"
objXL.Cells(10, 1).Value = "Sep"
objXL.Cells(11, 1).Value = "Oct"
objXL.Cells(12, 1).Value = "Nov"
objXL.Cells(13, 1).Value = "Dec"

objXL.Cells(2, 2).Value = "100"
objXL.Cells(3, 2).Value = "200"
objXL.Cells(4, 2).Value = "159"
objXL.Cells(5, 2).Value = "250"
objXL.Cells(6, 2).Value = "90"
objXL.Cells(7, 2).Value = "300"
objXL.Cells(8, 2).Value = "400"
objXL.Cells(9, 2).Value = "500"
objXL.Cells(10, 2).Value = "400"
objXL.Cells(11, 2).Value = "300"
objXL.Cells(12, 2).Value = "200"
objXL.Cells(13, 2).Value = "100"

objXL.Range("A15").Select
dim rng
set rng = objXL.Sheets("Sheet1").Range("A1:B13")
objXL.Charts.Add
objXL.ActiveChart.ChartType = 1
objXL.ActiveChart.SetSourceData rng, 2
objXL.ActiveChart.Location 2, "Sheet1"
objXL.Application.CommandBars("Chart").Visible = False



