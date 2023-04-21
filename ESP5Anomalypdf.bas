Attribute VB_Name = "ESP5Anomalypdf"
Sub Anomaly()
  Dim sheetsArray As Sheets
  Dim xWs As Worksheet
  Dim pt As PivotTable
  Dim pi As PivotItem
  Dim pf As PivotField
  Dim pi2 As PivotItem
  Set pt = Sheets("ESP5 Score Graph").PivotTables("PivotTable3")
  Set pf = pt.PivotFields("CENTRE CODE")
  
  'go through pivot filter and filter other worksheets
  
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
        .DataLabel.Font.Size = 32
        .DataLabel.Font.Color = vbBlack
   End With
   
   'Format sheets to be exported
   
   For Each xWs In ActiveWorkbook.Worksheets
    With ActiveSheet.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
      End With
  Next
  On Error Resume Next
  
  'export to pdf and place in a folder based on district
  
   If Sheets("ESP5 Score Graph").Range("F1") = "Victoria" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\Victoria\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B23") & "-" & Sheets("ESP5 Score Graph").Range("B27") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("ESP5 Score Graph").Range("F1") = "Caroni" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\Caroni\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B23") & "-" & Sheets("ESP5 Score Graph").Range("B27") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
    ElseIf Sheets("ESP5 Score Graph").Range("F1") = "North Eastern" Then
       ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
       "C:\Users\" & Environ("username") & "\Documents\North Eastern\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B23") & "-" & Sheets("ESP5 Score Graph").Range("B27") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
    ElseIf Sheets("ESP5 Score Graph").Range("F1") = "South Eastern" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\South Eastern\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B23") & "-" & Sheets("ESP5 Score Graph").Range("B27") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("ESP5 Score Graph").Range("F1") = "St George East" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\St. George East\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B23") & "-" & Sheets("ESP5 Score Graph").Range("B27") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("ESP5 Score Graph").Range("F1") = "Port Of Spain" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\Port Of Spain\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B23") & "-" & Sheets("ESP5 Score Graph").Range("B27") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("ESP5 Score Graph").Range("F1") = "Tobago" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\Tobago\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B23") & "-" & Sheets("ESP5 Score Graph").Range("B27") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   Else
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\St. Patrick\" & Sheets("ESP5 Score Graph").Range("A4") & " ESP5 Report " & Sheets("ESP5 Score Graph").Range("B23") & "-" & Sheets("ESP5 Score Graph").Range("B27") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   End If
   
   'delete trendline
   
   Sheets("ESP5 Score Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines(1).Delete
   
End Sub


