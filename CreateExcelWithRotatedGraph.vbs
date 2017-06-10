dim app
       set app = createobject("Excel.Application")

       ' Make it visible
       app.Visible = true

       ' Add a new workbook
       dim wb
       set wb = app.workbooks.add

       ' Fill array of values first...
       dim arr(19,9) ' Note: VBScript is zero-based
       for i = 1 to 20
          for j = 1 to 10
             arr(i-1,j-1) = i*j
          next
       next

       ' Declare a range object to hold our data
       dim rng
       set rng = wb.Activesheet.Range("A1").Resize(20,10)

       ' Now assign them all in one shot...
       rng.value = arr

       ' Add a new chart based on the data
       wb.Charts.Add
       wb.ActiveChart.ChartType = 70 'xl3dPieExploded
       wb.ActiveChart.SetSourceData rng, 2 ' xlColumns
       wb.ActiveChart.Location 2, "Sheet1" 'xlLocationAsObject

       ' Rotate it around...
       for i = 1 to 360 step 30
          wb.activechart.rotation = i
       next