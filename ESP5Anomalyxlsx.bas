Attribute VB_Name = "ESP5Anomalyxlsx"
Sub Pivot_Loopx()
  Dim xArray As Sheets
  Dim xWs As Worksheet
  Dim pt As PivotTable
  Dim pi As PivotItem
  Dim pf As PivotField
  Dim pi2 As PivotItem
  Set pt = Sheets("ESP5 Score Graph").PivotTables("PivotTable3")
  Set pf = pt.PivotFields("CENTRE CODE")
  
  'make an array of pages to be copied and pasted once filtered
  
  Set xArray = ActiveWorkbook.Sheets(Array("ESP5 Score Graph", "Progress vs Attainment", "Attainment & Progress(no rank)"))
  
  'go through pivot filter and filter sheets accordingly
  
  For Each pi2 In pf.PivotItems
      If pi2.Visible = True Then
        Sheets("Progress vs Attainment").PivotTables("PivotTable5").PivotFields("CENTRE CODE").ClearAllFilters
        Sheets("Progress vs Attainment").PivotTables("PivotTable5").PivotFields("CENTRE CODE").CurrentPage = pi2.Value
        Sheets("Attainment & Progress(no rank)").ListObjects("Attainment").HeaderRowRange(2).AutoFilter 2, pi2
        Sheets("Attainment & Progress(no rank)").ListObjects("Progress").HeaderRowRange(2).AutoFilter 2, pi2
      End If
   Next
   
   'Create a chart trendline
   
   Sheets("ESP5 Score Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines.Add
     With Sheets("ESP5 Score Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines(1)
        .Type = xlLinear
        .Format.Line.DashStyle = msoLineSysDot
        .Format.Line.Weight = 3
        .DisplayEquation = True
        .DataLabel.Font.Size = 24
        .DataLabel.Font.Color = vbBlack
   End With
   
   'add new workbook
  
  Set newBook = Workbooks.Add(xlWBATWorksheet)
  
  'go through xArray and copy and paste each filtered sheet to new workbook

  With xArray(1)
    .Range("A1:F30").SpecialCells(xlCellTypeVisible).Copy
    newBook.Activate
    ActiveSheet.Name = xArray(1).Name
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
    With ActiveSheet
        .Range("E1:F1,A5").Font.Size = 18
        .Range("A4").Font.Size = 24
        .Range("A19:C27").Font.Size = 14
        .Range("A1").ColumnWidth = 18.86
        .Range("B1").ColumnWidth = 10.71
        .Range("C1").ColumnWidth = 62
        .Range("D1").ColumnWidth = 11.86
        .Range("E1").ColumnWidth = 16.43
        .Range("F1").ColumnWidth = 27.86
        .Range("A1:A11").RowHeight = 30
        .Range("A12:A18").EntireRow.Delete
        .Range("A12:A20").RowHeight = 18.75
        .Range("C14:C15").Font.Bold = True
        .Range("A14").Value = ""
        With Range("A14:A15")
            .Merge
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With Range("B14:B15")
            .Merge
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range("A12:B12,A14:C15").Interior.Color = RGB(221, 235, 247)
    End With
  End With
   
   'copy and reformat textbox
    
   xArray(1).Shapes("TextBox 1").Copy
   newBook.Sheets("ESP5 Score Graph").Activate
   newBook.Sheets("ESP5 Score Graph").Range("A7").Select
   With ActiveSheet
    .Paste
    .Shapes("TextBox 1").TextFrame.Characters.Font.Size = 14
    .Shapes("TextBox 1").Top = Sheets("ESP5 Score Graph").Range("A7").Top + 0.5
    .Shapes("TextBox 1").Width = Sheets("ESP5 Score Graph").Range("A7:F7").Width
    .Shapes("TextBox 1").Height = Sheets("ESP5 Score Graph").Range("A7:A10").Height - 0.5
    .Shapes("TextBox 1").Fill.BackColor.RGB = RGB(255, 255, 255)
   End With
   
   'make chart
   
  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Set Ws = ActiveSheet
  Set Rang = Ws.Range("B14:C20")
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Source:=Rang, PlotBy:=xlRows
        .ChartType = xlXYScatterLines
        .ChartTitle.Text = "Value Added ESP5 Score"   'Title
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Color = vbBlack
        .HasLegend = False
        .PlotBy = xlColumns
        .Axes(xlCategory).MinimumScale = 2018    'Adjust scale
        .Axes(xlCategory).MaximumScale = 2022
        .Axes(xlCategory).MajorUnit = 1
        .Axes(xlValue).HasMajorGridlines = True  'Remove Gridlines
        .Axes(xlCategory).HasMajorGridlines = True
        .Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 12
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .Axes(xlCategory).TickLabels.Font.Color = vbBlack
        .Axes(xlCategory).TickLabels.Font.Size = 12
        .SeriesCollection(1).Trendlines.Add
         With .SeriesCollection(1).Trendlines(1)
            .Type = xlLinear
            .DisplayEquation = True
            .Format.Line.DashStyle = msoLineSysDot
            .Format.Line.Weight = 2.5
            .DataLabel.Font.Size = 18
            .DataLabel.Font.Color = vbBlack
        End With
        With .Parent
           .Left = Ws.Range("A23").Left
           .Top = Ws.Range("A23").Top
           .Width = Ws.Range("A23:F23").Width
           .Height = Ws.Range("A23:A48").Height
        End With
  End With

    With xArray(2)
        .Range("A1:E27").SpecialCells(xlCellTypeVisible).Copy
        newBook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = xArray(2).Name
        With ActiveSheet
            .Range("A1").ColumnWidth = 52.86
            .Range("B1").ColumnWidth = 10
            .Range("C1").ColumnWidth = 25.86
            .Range("D1").ColumnWidth = 40.43
            .Range("E1").ColumnWidth = 34.29
        End With
        ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
        With ActiveSheet
            .Range("A3").Font.Size = 24
            .Range("A4:A5").RowHeight = 18
            .Range("C6:E11").Font.Size = 14
            .Range("A13:A17").EntireRow.Delete
        End With
End With
 
    'copy and reformat textbox

    xArray(2).Shapes("TextBox").Copy
    newBook.Sheets("Progress vs Attainment").Paste
    With ActiveSheet.Shapes("TextBox")
        .TextFrame.Characters.Font.Size = 14
        .Top = ActiveSheet.Range("A14").Top + 0.5
        .Left = ActiveSheet.Range("A14").Left
        .Height = ActiveSheet.Range("A14:A17").Height - 0.5
        .Width = ActiveSheet.Range("A14:E14").Width
        .Fill.BackColor.RGB = RGB(255, 255, 255)
    End With
    
    'copy and reformat picture
    
    xArray(2).Shapes("Quadrant").Copy
    newBook.Sheets("Progress vs Attainment").Paste
    With ActiveSheet.Shapes("Quadrant")
        .Top = ActiveSheet.Range("A6").Top
        .Left = ActiveSheet.Range("A6").Left
        .Width = ActiveSheet.Range("A6:B6").Width
        .Height = ActiveSheet.Range("A6:A12").Height
    End With

    xArray(2).Range("A28:E37").SpecialCells(xlCellTypeVisible).Copy
    newBook.Sheets("Progress vs Attainment").Range("A19").PasteSpecial Paste:=xlPasteValues
    newBook.Sheets("Progress vs Attainment").Range("A19").PasteSpecial Paste:=xlPasteFormats
    With ActiveSheet
        .Range("C21:C26").Cut Range("E21:E26")
        .Range("C21:C26").Delete
        .Range("A19:D26").Font.Size = 14
        .Range("A19:B19,A21:D21").Interior.Color = RGB(221, 235, 247)
        .Range("A27:A28").RowHeight = 18
    End With
    
    'copy and reformat chart

    xArray(2).ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.ChartArea.Copy
    newBook.Sheets("Progress vs Attainment").Activate
    newBook.Sheets("Progress vs Attainment").Range("A29").Select
    ActiveSheet.Paste
    
    Set Rang = ActiveSheet.Range("C21:D26")

    With ActiveChart
        .SetSourceData Source:=Rang
        .ChartTitle.Font.Size = 18
        .Axes(xlValue).TickLabels.Font.Size = 12
        .Axes(xlCategory).TickLabels.Font.Size = 12
        .Axes(xlValue).AxisTitle.Font.Size = 12
        .Axes(xlCategory).AxisTitle.Font.Size = 12
        .SeriesCollection(1).Format.Line.Weight = 3
        .SeriesCollection(1).Format.Line.Visible = False
        .SeriesCollection(1).MarkerSize = 8
        .Parent.Top = ActiveSheet.Range("A29").Top
        .Parent.Left = ActiveSheet.Range("A29").Left
        .Parent.Width = ActiveSheet.Range("A29:E29").Width
        .Parent.Height = ActiveSheet.Range("A29:A56").Height
    End With
    
    
    With xArray(3)
        .Range("A1: L3000").SpecialCells(xlCellTypeVisible).Copy
        newBook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = xArray(3).Name
         With ActiveSheet
            .Range("A1").ColumnWidth = 8.14
            .Range("B1").ColumnWidth = 10.43
            .Range("C1").ColumnWidth = 34.71
            .Range("D1").ColumnWidth = 12.71
            .Range("E1").ColumnWidth = 11.57
            .Range("F1").ColumnWidth = 13.57
            .Range("G1").ColumnWidth = 14.57
            .Range("H1").ColumnWidth = 10.86
            .Range("J1").ColumnWidth = 11.43
            .Range("I1").ColumnWidth = 13.57
            .Range("K1").ColumnWidth = 13.43
            .Range("L1").ColumnWidth = 12.14
        End With
        ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
        ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
        last = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
        With ActiveSheet
            .Range("A5:L20").Font.Size = 14
            .Range("A4:L4,A14:L14").Font.Size = 12
            .Range("A2,A12").Font.Size = 24
            .Range("A14:L14").Borders.LineStyle = xlContinuous
            .Range("A2").RowHeight = 31.5
        End With
End With

   'Format sheets to be exported

   If Sheets("ESP5 Score Graph").Range("F1") = "Victoria" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\Victoria\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B16") & "-" & Sheets("ESP5 Score Graph").Range("B20") & ".xlsx"
    newBook.Close SaveChanges:=False
   

   ElseIf Sheets("ESP5 Score Graph").Range("F1") = "Caroni" Then
    newBook.SaveAs _
             Filename:="C:\Users\" & Environ("username") & "\Documents\Caroni\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B16") & "-" & Sheets("ESP5 Score Graph").Range("B20") & ".xlsx"
     newBook.Close SaveChanges:=False
   

    ElseIf Sheets("ESP5 Score Graph").Range("F1") = "North Eastern" Then
        newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\North Eastern\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B16") & "-" & Sheets("ESP5 Score Graph").Range("B20") & ".xlsx"
        newBook.Close SaveChanges:=False
       
    ElseIf Sheets("ESP5 Score Graph").Range("F1") = "South Eastern" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\South Eastern\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B16") & "-" & Sheets("ESP5 Score Graph").Range("B20") & ".xlsx"
    newBook.Close SaveChanges:=False
   
   ElseIf Sheets("ESP5 Score Graph").Range("F1") = "St George East" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\St. George East\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B16") & "-" & Sheets("ESP5 Score Graph").Range("B20") & ".xlsx"
    newBook.Close SaveChanges:=False

   ElseIf Sheets("ESP5 Score Graph").Range("F1") = "Port Of Spain" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\Port of Spain\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B16") & "-" & Sheets("ESP5 Score Graph").Range("B20") & ".xlsx"
    newBook.Close SaveChanges:=False

   ElseIf Sheets("ESP5 Score Graph").Range("F1") = "Tobago" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\Tobago\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B16") & "-" & Sheets("ESP5 Score Graph").Range("B20") & ".xlsx"
    newBook.Close SaveChanges:=False

   Else
   newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\St. Patrick\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B16") & "-" & Sheets("ESP5 Score Graph").Range("B20") & ".xlsx"
    newBook.Close SaveChanges:=False
   
   End If

   'delete trendline
   
   Sheets("ESP5 Score Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines(1).Delete
   
End Sub


