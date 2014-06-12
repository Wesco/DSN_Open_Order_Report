Attribute VB_Name = "CreateReport"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : CreateDSNReport
' Date : 1/24/2014
' Desc : Create an open order report to send to Doosan
'---------------------------------------------------------------------------------------
Sub CreateDSNReport()
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim Col As Integer
    Dim i As Integer

    Sheets("Prev OOR").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    For i = 1 To TotalCols
        Select Case Cells(1, i).Value
            Case "Order Number"
                Col = 2
            Case "Release Number"
                Col = 3
            Case "Shipment Number"
                Col = 4
            Case "Part Number"
                Col = 1
            Case "Description"
                Col = 5
            Case "Due Date"
                Col = 6
            Case "Ordered"
                Col = 7
            Case "BO"
                Col = 8
            Case "RTS"
                Col = 9
            Case "Status"
                Col = 10
            Case "Notes"
                Col = 11
            Case Else
                Col = 0
        End Select

        If Col > 0 Then
            Range(Cells(1, i), Cells(TotalRows, i)).Copy Destination:=Sheets("DSN Report").Cells(1, Col)
        End If
    Next
End Sub

Sub CreateOOR()
    Dim TotalRows As Long
    Dim PrevOORCols As Integer

    Sheets("DSN OOR").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    
    Sheets("Prev OOR").Select
    PrevOORCols = Columns(Columns.Count).End(xlToLeft).Column
    
    Sheets("DSN OOR").Select
    Range("A1:A" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("A1")    'UID
    Range("E1:E" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("B1")    'Order Number
    Range("G1:G" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("C1")    'Release Number
    Range("H1:H" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("D1")    'Release Number
    Range("I1:I" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("E1")    'Shipment Number
    Range("C1:C" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("F1")    'Part Number
    Range("D1:D" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("G1")    'Description
    Range("N1:N" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("H1")    'Due Date
    Range("Y1:Y" & TotalRows).Copy Destination:=Sheets("Open Order Report").Range("I1")    'Creation Date

    Sheets("Open Order Report").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    AddColumn "Wesco Order", "=IFERROR(VLOOKUP(A2,117!A:B,2,FALSE),"""")"
    AddColumn "Wesco PO", "=IFERROR(VLOOKUP(A2,117!A:L,12,FALSE),"""")"
    
    'F2 = Part Number
    AddColumn "SIM", "=IFERROR(IF(VLOOKUP(F2,Master!A:B,2,FALSE)=0,"""",""'""&VLOOKUP(F2,Master!A:B,2,FALSE)),"""")"
    
    AddColumn "Supplier", "=IFERROR(IF(VLOOKUP(A2,117!A:N,14,FALSE)=0,"""",""'""&VLOOKUP(A2,117!A:N,14,FALSE)),"""")"
    AddColumn "Promise Date", "=IFERROR(IF(VLOOKUP(A2,117!A:M,13,FALSE)=0,"""",VLOOKUP(A2,117!A:M,13,FALSE)),"""")", "m/d/yyyy"
    AddColumn "Ordered", "=IFERROR(VLOOKUP(A2,'DSN OOR'!A:K,11,FALSE),0)"
    AddColumn "BO", "=IFERROR(VLOOKUP(A2,117!A:J,10,FALSE),0)"
    AddColumn "RTS", "=IFERROR(VLOOKUP(A2,117!A:I,9,FALSE),0)"
    AddColumn "Old Status", "=IFERROR(IF(VLOOKUP(A2,'Prev OOR'!A:Z," & PrevOORCols - 1 & ",FALSE)=0,"""",VLOOKUP(A2,'Prev OOR'!A:Z," & PrevOORCols - 1 & ",FALSE)),"""")"
    AddColumn "Status", "=IF(NOT(IFERROR(VLOOKUP(A2,117!A:A,1,FALSE),"""")="""")=TRUE,IF(IFERROR(VLOOKUP(A2,117!A:J,10,FALSE),0)>0,""B/O"",IF(IFERROR(VLOOKUP(A2,'DSN OOR'!A:K,11,FALSE),0)=IFERROR(VLOOKUP(A2,117!A:I,9,FALSE),0),""RTS"",IF(IFERROR(VLOOKUP(A2,117!A:K,11,FALSE),0)=IFERROR(VLOOKUP(A2,'DSN OOR'!A:K,11,FALSE),0),""SHIPPED"",""CHECK""))),""NOO"")"
    AddColumn "Notes", "=IFERROR(IF(VLOOKUP(A2,'Prev OOR'!A:Z," & PrevOORCols & ",FALSE)=0,"""",VLOOKUP(A2,'Prev OOR'!A:Z," & PrevOORCols & ",FALSE)),"""")", "mmm dd, yyyy"
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
