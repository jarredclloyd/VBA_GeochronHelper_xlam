Attribute VB_Name = "Tools"
'Module for general tool functions
Option Explicit
    
Dim rngFirstRow As Range
Dim Rng As Range
Dim rngCol As Long
Dim ws As Worksheet
Dim LastCol As Long
Dim LastRow As Long
Dim Msg As String
Dim Ans As Variant
Private Sub Sorter()
'Created By Jarred Lloyd on 03-10-2018
'Modified on 08-02-2020 to improve performance and flexibility
'Feel free to modify but give credit and do not sell any version of this, modified or not. It is to remain free for those who need it
'This code will sort any data in a sheet within individual columns in ascending order (safe for headers in row one)
    Application.ScreenUpdating = False
       Set ws = ActiveSheet
    'Determine number of columns and define range
        LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        Set rngFirstRow = ws.Range("A1", Cells(1, LastCol))
    'Sort each column
        For Each Rng In rngFirstRow
            rngCol = Rng.Column
            LastRow = Cells(Rows.Count, rngCol).End(xlUp).Row
            With ws.Sort
                .SortFields.Clear
                .SortFields.Add Key:=Rng, Order:=xlAscending
                .SetRange ws.Range(Rng, Cells(LastRow, rngCol))
                .Header = xlYes
                .Apply
            End With
        Next Rng
    'Move to clean procedure to remove non numeric characters
        Ans = MsgBox("Do you want to remove text values from row two and below?", vbYesNo, "Text clean")
            Select Case Ans
                Case vbYes
                    Call Clean
                Case vbNo
                    GoTo Quit:
            End Select
Quit:
    Application.ScreenUpdating = True
End Sub
Private Sub Clean()
'This will remove non numeric values from from cells
    Dim CleanRng As Range
        Set ws = ActiveSheet
    'Determine number of columns and define range
        With ActiveSheet
            LastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
            LastRow = .UsedRange.Rows.Count
            Set CleanRng = .Range("A2", Cells(LastRow, LastCol))
        End With
    'Perform clean of non-numeric values starting from row 2
        On Error GoTo NoText
            CleanRng.SpecialCells(xlCellTypeConstants, xlTextValues).ClearContents
            Application.CutCopyMode = False
        Exit Sub
NoText: MsgBox ("No text values found in range (A2:LastRow,LastColumn)")
Application.CutCopyMode = False
End Sub
Sub SortConfirm(control As IRibbonControl)
'This is to prevent accidentally ruining an excel spreadsheet
    Msg = "Does this sheet contain data in columns you want to sort individually?" & vbCrLf & vbCrLf & "Have you saved the workbook?"
    Ans = MsgBox(Msg, vbYesNo, "Column sorter and text clean")
    Select Case Ans
        Case vbYes
            Call Sorter
        Case vbNo
            Exit Sub
    End Select
End Sub
