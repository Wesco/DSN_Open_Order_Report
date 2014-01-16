Attribute VB_Name = "FormatData"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : FormatOOR
' Date : 1/15/2014
' Desc : Formats Doosans open order report
'---------------------------------------------------------------------------------------
Sub FormatOOR()
    Dim TotalRows As Long
    
    Sheets("DSN OOR").Select
    
    'Determine if the report is aftermarket or production
    If Range("B1").Value = "PPZ Service Parts Inventory Org" Then
        OORType = "aftermarket"
    ElseIf Range("B1").Value = "STA Inventory Org" Then
        OORType = "production"
    Else
        Err.Raise CustErr.UNRECOGNIZED_REPORT, "FormatOOR", "The imported report is unrecognized."
    End If
    
    'Remove header data
    Rows("1:8").Delete
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    
    'Create UID column
    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = "=""'""&E2&""-""&G2&C2"
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value
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

    'Create UID column
    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = "=""="""""" & C2 & D2 & """""""""
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value
End Sub
