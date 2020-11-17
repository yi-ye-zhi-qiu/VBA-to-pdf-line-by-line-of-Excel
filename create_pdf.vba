Sub create_pdf()

  Dim File As String
  Dim Folder As String
  Dim index As Long
  Dim pdfRng As Range
  Dim Rng As Range
  Dim Wks As Worksheet

  Folder = "C:\"

  Set Wks = Worksheets("Sheet_Name")

  Set Rng = Wks.Cells(1, "A").CurrentRegion

  Wks.Names.Add "Print_Area", Rng

  Set pdfRng = Intersect(Rng, Rng.Offset(1, 0))
  If pdfRng Is Nothing Then Exit Sub Else pdfRng.Rows.Hidden = True

  For index = 1 To pdfRng.Rows.Count
    With Worksheets("Sheet_Name").PageSetup
      .Orientation = xlLandscape
      .FitToPagesWide = 1
      .FitToPagesTall = False
      .Zoom = False
    End With

    pdfRng.Rows(index).Hidden = False

    File = Folder & pdfRng.Item(index, 1) & ".pdf"
    Wks.Names("Print_Area").RefersToRange.ExportAsFixedFormat xlTypePDF, File
    pdfRng.Rows(index).Hidden = True
  Next index

  Rng.Rows.Hidden = False

End Sub
