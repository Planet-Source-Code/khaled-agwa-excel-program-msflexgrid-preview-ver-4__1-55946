Attribute VB_Name = "Mform2"
Function FlexGridPrint(fgPrint As MSFlexGrid, Optional lOrientation As Long = vbPRORPortrait, Optional ByVal lMaxRowsPerPage As Long = -1, Optional lTopBorder As Long = 1000, Optional lLeftBorder As Long = 1000, Optional lRowsToRepeat As Long = 0) As Boolean
   Dim lRowsPrinted As Long, lRowsPerPage As Long
   Dim lThisRow As Long, lNumRows As Long, lImageHeight As Long, lLastImageTop As Long
   Dim lPrinterPageHeight As Long, lPagesPrinted As Long, lHeadingHeight As Long
   
   On Error GoTo ErrFailed
  ' Printer.Orientation = lOrientation
   lNumRows = fgPrint.Rows - 1
   lPrinterPageHeight = Printer.Height
   lRowsPerPage = lMaxRowsPerPage
   lRowsPrinted = lRowsToRepeat
   
   If lRowsToRepeat Then
       'Calculate the height of the heading row
       For lThisRow = 1 To lRowsToRepeat
           lHeadingHeight = lHeadingHeight + fgPrint.RowHeight(lThisRow)
       Next
   End If

   Do
       'Calculate the number of rows for this page
       lImageHeight = 0
       lRowsPerPage = 0
       For lThisRow = lRowsPrinted To lNumRows
           lImageHeight = lImageHeight + fgPrint.RowHeight(lThisRow)
           If lRowsPerPage > lMaxRowsPerPage And lMaxRowsPerPage <> -1 Then
               'Image has required number of rows, subtract height of current row
               lImageHeight = lImageHeight - fgPrint.RowHeight(lThisRow)
               Exit For
           ElseIf lImageHeight + lTopBorder * 2 + lHeadingHeight > lPrinterPageHeight Then           'Allow the same border at the bottom and top
               'Image is larger than page, subtract height of current row
               lImageHeight = lImageHeight - fgPrint.RowHeight(lThisRow)
               Exit For
           End If
           lRowsPerPage = lRowsPerPage + 1
       Next
       
       'Print this page
       lPagesPrinted = lPagesPrinted + 1
       If lRowsToRepeat Then
           'Print heading rows
           Printer.PaintPicture fgPrint.Picture, lLeftBorder, lTopBorder, , lHeadingHeight, , 0, , lHeadingHeight
           'Print data rows
           Printer.PaintPicture fgPrint.Picture, lLeftBorder, lTopBorder + lHeadingHeight, , lImageHeight + lHeadingHeight, , lLastImageTop + lHeadingHeight, , lImageHeight + lHeadingHeight
       Else
           'Print data rows
           Printer.PaintPicture fgPrint.Picture, lLeftBorder, lTopBorder, , lImageHeight, , lLastImageTop, , lImageHeight
       End If
       
       Printer.EndDoc
       
       'Store printer position
       lRowsPrinted = lRowsPrinted + lRowsPerPage
       lLastImageTop = lLastImageTop + lImageHeight + lHeadingHeight
   
   Loop While lRowsPrinted < lNumRows
   
   FlexGridPrint = True
   
   Exit Function
   
ErrFailed:
   'Failed to print grid
   FlexGridPrint = False
   Debug.Print "Error in FlexGridPrint: " & Err.Description
End Function




