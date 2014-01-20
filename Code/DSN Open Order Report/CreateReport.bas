Attribute VB_Name = "CreateReport"
Option Explicit

Sub CreateOOR()
    Dim TotalRows As Long
    Dim PrevOORCols As Integer

    'Get UIDs from DSN OOR
    Sheets("DSN OOR").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    PrevOORCols = Sheets("Prev OOR").UsedRange.Columns.Count
    Range("A1:A" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("A1")

    Sheets("Open Order Report").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    AddColumn "Order Number", "=IFERROR(VLOOKUP(A2,'DSN OOR'!A:E,5,FALSE),"""")"
    AddColumn "Release Number", "=IFERROR(VLOOKUP(A2,'DSN OOR'!A:G,7,FALSE),"""")"
    AddColumn "Shipment Number", "=IFERROR(VLOOKUP(A2,'DSN OOR'!A:I,9,FALSE),"""")"
    AddColumn "Part Number", "=IFERROR(VLOOKUP(A2,'DSN OOR'!A:C,3,FALSE),"""")"
    AddColumn "Description", "=IFERROR(VLOOKUP(A2,'DSN OOR'!A:D,4,FALSE),"""")"
    AddColumn "Due Date", "=IFERROR(VLOOKUP(A2,'DSN OOR'!A:N,14,FALSE),"""")"
    AddColumn "Order Number", "=IFERROR(VLOOKUP(A2,117!A:B,2,FALSE),"""")"
    AddColumn "PO Number", "=IFERROR(VLOOKUP(A2,117!A:L,12,FALSE),"""")"
    AddColumn "Supplier", "=IFERROR(VLOOKUP(A2,117!A:N,14,FALSE),"""")"
    AddColumn "Promise Date", "=IFERROR(IF(VLOOKUP(A2,117!A:M,13,FALSE)=0,"""",VLOOKUP(A2,117!A:M,13,FALSE)),"""")", "mmm dd, yyyy"
    AddColumn "Ordered", "=IFERROR(VLOOKUP(A2,'DSN OOR'!A:K,11,FALSE),"""")"
    AddColumn "BO", "=IFERROR(VLOOKUP(A2,117!A:J,10,FALSE),"""")"
    AddColumn "RTS", "=IFERROR(VLOOKUP(A2,117!A:I,9,FALSE),"""")"
    AddColumn "Old Status", "=IFERROR(IF(VLOOKUP(A2,'Prev OOR'!A:Z," & PrevOORCols - 1 & ",FALSE)=0,"""",VLOOKUP(A2,'Prev OOR'!A:Z," & PrevOORCols - 1 & ",FALSE)),"""")"
    AddColumn "Status", "=IF(NOT(IFERROR(VLOOKUP(A2,117!A:A,1,FALSE),"""")="""")=TRUE,IF(IFERROR(VLOOKUP(A2,117!A:J,10,FALSE),0)>0,""B/O"",IF(L2=IFERROR(VLOOKUP(A2,117!A:I,9,FALSE),0),""RTS"",IF(IFERROR(VLOOKUP(A2,117!A:K,11,FALSE),0)=L2,""SHIPPED"",""CHECK""))),""NOO"")"
    'Add Note lookup
End Sub

Private Sub FillColumn(Rng As Range, Formula As String, Optional NumberFormat As String = "General")
    With Rng
        If .NumberFormat <> "General" Then .NumberFormat = "General"
        .Formula = Formula
        .NumberFormat = NumberFormat
        .Value = .Value
    End With
End Sub

Private Sub AddColumn(Header As String, Formula As String, Optional NumberFormat As String = "General")
    Dim TotalRows As Long
    Dim TotalCols As Integer

    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count + 1

    Cells(1, TotalCols).Value = Header

    With Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols))
        .Formula = Formula
        .NumberFormat = NumberFormat
        .Value = .Value
    End With
End Sub
