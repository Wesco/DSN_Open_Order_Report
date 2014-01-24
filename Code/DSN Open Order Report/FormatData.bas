Attribute VB_Name = "FormatData"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : FormatDSNOOR
' Date : 1/15/2014
' Desc : Formats Doosans open order report
'---------------------------------------------------------------------------------------
Sub FormatDSNOOR()
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim i As Integer

    Sheets("DSN OOR").Select

    'Remove header data
    Rows("1:8").Delete
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Store correct column header order
    ColHeaders = Array("Source Type", _
                       "Part Number", _
                       "Description", _
                       "Order Number", _
                       "Buyer Name", _
                       "Release Number", _
                       "Line Number", _
                       "Shipment Number", _
                       "Expected Quantity", _
                       "Ordered Quantity", _
                       "Expected ASN Quantity", _
                       "Lead Time", _
                       "Due Date", _
                       "Supplier Number", _
                       "Supplier Name", _
                       "Supplier Site", _
                       "Planner")

    'Compare correct column headers to actual column headers
    For i = 0 To UBound(ColHeaders)
        If Cells(1, i + 1).Value <> ColHeaders(i) Then
            Err.Raise CustErr.INVALID_COLUMN_ORDER, "FormatDSNOOR", "The Doosan OOR column order is invalid."
        End If
    Next

    'Create UID column
    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = "=""'""&E2&""-""&G2&C2"
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value

    TotalCols = ActiveSheet.UsedRange.Columns.Count
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value

    If OORType = "production" Then
        ActiveSheet.UsedRange.AutoFilter 18, "=WESCO"
        Cells.Delete
        Rows(1).Insert
        Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : Format117
' Date : 1/15/2014
' Desc : Removes columns
'---------------------------------------------------------------------------------------
Sub Format117()
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim PartList As Variant
    Dim DescList As Variant
    Dim Result As Variant
    Dim i As Long
    Dim j As Long

    Sheets("117").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Remove report footer
    Rows(ActiveSheet.UsedRange.Rows.Count).Delete

    'Remove report header
    Rows(1).Delete

    'Remove all unneeded columns
    For i = TotalCols To 1 Step -1
        If Cells(1, i).Value <> "ORDER NO" And _
           Cells(1, i).Value <> "CUSTOMER REFERENCE NO" And _
           Cells(1, i).Value <> "CUSTOMER PART NUMBER" And _
           Cells(1, i).Value <> "LINE NO" And _
           Cells(1, i).Value <> "ITEM DESCRIPTION" And _
           Cells(1, i).Value <> "ORDER QTY" And _
           Cells(1, i).Value <> "AVAILABLE QTY" And _
           Cells(1, i).Value <> "QTY TO SHIP" And _
           Cells(1, i).Value <> "BO QTY" And _
           Cells(1, i).Value <> "QTY SHIPPED" And _
           Cells(1, i).Value <> "PO NUMBER" And _
           Cells(1, i).Value <> "PROMISE DATE" And _
           Cells(1, i).Value <> "SUPPLIER NUM" Then
            Columns(i).Delete
        End If
    Next

    'Load the correct column order into an array
    ColHeaders = Array("ORDER NO", _
                       "CUSTOMER REFERENCE NO", _
                       "CUSTOMER PART NUMBER", _
                       "LINE NO", _
                       "ITEM DESCRIPTION", _
                       "ORDER QTY", _
                       "AVAILABLE QTY", _
                       "QTY TO SHIP", _
                       "BO QTY", _
                       "QTY SHIPPED", _
                       "PO NUMBER", _
                       "PROMISE DATE", _
                       "SUPPLIER NUM")

    'Compare the correct column order to the actual column order
    For i = 0 To UBound(ColHeaders)
        If Cells(1, i + 1).Value <> ColHeaders(i) Then
            Err.Raise CustErr.INVALID_COLUMN_ORDER, "Format117", "The column order on the 117 report is incorrect."
        End If
    Next

    TotalCols = ActiveSheet.UsedRange.Columns.Count
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Remove spaces from "CUSTOMER REFERENCE NO" & "CUSTOMER PART NUMBER"
    Range("B2:C" & TotalRows).Replace "=""", ""
    Range("B2:C" & TotalRows).Replace """", ""
    Range("B2:C" & TotalRows).Replace " ", ""

    'Convert customer part numbers to strings
    Columns("C:C").Insert
    Range("C2:C" & TotalRows).Formula = "=""'"" & D2"
    Range("C2:C" & TotalRows).Value = Range("C2:C" & TotalRows).Value
    Columns("D:D").Delete

    'Remove extra spaces from part descriptions
    Columns(5).Insert
    Range("E1").Value = "ITEM DESCRIPTION"
    Range("E2:E" & TotalRows).Formula = "=TRIM(F2)"
    Range("E2:E" & TotalRows).Value = Range("E2:E" & TotalRows).Value
    Columns(6).Delete

    'Remove extra spaces from promise dates
    Columns(12).Insert
    Range("L1").Value = "PROMISE DATE"
    Range("L2:L" & TotalRows).Formula = "=SUBSTITUTE(M2,"" "","""")"
    Range("L2:L" & TotalRows).Value = Range("L2:L" & TotalRows).Value
    Range("L2:L" & TotalRows).NumberFormat = "yyyy/mm/dd"
    Columns(13).Delete

    'Remove extra spaces from supplier numbers
    Columns(13).Insert
    Range("M1").Value = "SUPPLIER NUM"
    Range("M2:M" & TotalRows).Formula = "=IF(SUBSTITUTE(N2,"" "","""")="""","""",""'""&SUBSTITUTE(N2,"" "",""""))"
    Range("M2:M" & TotalRows).Value = Range("M2:M" & TotalRows).Value
    Columns(14).Delete

    'Lookup missing part numbers by description on Master
    Columns(3).Insert
    Range("C1").Value = "CUSTOMER PART NUMBER"
    'If the part number is blank try to find it on the master using its description
    Range("C2:C" & TotalRows).Formula = "=IFERROR(IF(D2="""",""'"" & INDEX(Master!A:A,MATCH(F2,Master!C:C,0)),""'"" & D2),"""")"
    Range("C2:C" & TotalRows).Value = Range("C2:C" & TotalRows).Value

    Columns(4).Delete

    'Load part numbers into an array
    Sheets("Master").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    PartList = Range("A2:A" & TotalRows)

    'Load item descriptions into an array
    Sheets("117").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    DescList = Range("E2:E" & TotalRows)

    'Find part numbers in item descriptions
    For i = 1 To UBound(DescList)
        'If CUSTOMER PART NUMBER is blank
        If Cells(i + 1, 3).Value = "" Then
            'See if any part numbers are in the item description
            For j = 1 To UBound(PartList)
                Result = InStr(1, DescList(i, 1), PartList(j, 1))
                If Result <> 0 Then
                    Cells(i + 1, 3).Value = "'" & PartList(j, 1)
                    Exit For
                End If
            Next
        End If
    Next

    'Load customer part list into array
    Sheets("DSN OOR").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    PartList = Range("C2:C" & TotalRows)

    'Find part numbers in item descriptions
    Sheets("117").Select
    For i = 1 To UBound(DescList)
        'If CUSTOMER PART NUMBER is blank
        If Cells(i + 1, 3).Value = "" Then
            'See if any part numbers are in the item description
            For j = 1 To UBound(PartList)
                If PartList(j, 1) <> "" Then
                    Result = InStr(1, DescList(i, 1), PartList(j, 1))
                Else
                    Result = 0
                End If
                If Result <> 0 Then
                    Cells(i + 1, 3).Value = PartList(j, 1)
                    Exit For
                End If
            Next
        End If
    Next

    'Create UID column
    Columns(1).Insert
    Range("A1").Value = "UID"
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Range("A2:A" & TotalRows).Formula = "=""="""""" & C2 & D2 & """""""""
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value
End Sub

'---------------------------------------------------------------------------------------
' Proc : FormatReport
' Date : 1/20/2014
' Desc : Format the open order report
'---------------------------------------------------------------------------------------
Sub FormatReport(Sheet As Worksheet)
    Dim TotalCols As Integer

    Sheet.Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    With ActiveSheet.UsedRange
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.ColorIndex = xlAutomatic
        .Font.TintAndShade = 0
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Strikethrough = False
        .Font.Superscript = False
        .Font.Subscript = False
        .Font.OutlineFont = False
        .Font.Shadow = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.ColorIndex = xlAutomatic
        .Font.TintAndShade = 0
        .Font.ThemeFont = xlThemeFontMinor
        .Font.Bold = False
    End With

    ActiveSheet.ListObjects.Add(xlSrcRange, ActiveSheet.UsedRange, , xlYes).Name = "Table1"
    ActiveSheet.ListObjects(1).Unlist
    Range(Cells(1, 1), Cells(1, TotalCols)).Font.Color = RGB(255, 255, 255)
    Range("A1").Select
End Sub

Sub FormatDSNReport()
    Dim TotalRows As Integer
    Dim i As Integer

    Sheets("DSN Report").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    For i = 2 To TotalRows
        If Cells(i, 8).Value > 0 Then
            With Cells(i, 8).Interior
                .Pattern = xlSolid
                .PatternThemeColor = xlThemeColorAccent1
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0.799981688894314
            End With
        End If
    Next
    
    ActiveSheet.Columns.AutoFit
End Sub
