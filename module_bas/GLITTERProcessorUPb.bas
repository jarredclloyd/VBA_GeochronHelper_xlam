Attribute VB_Name = "GlitterProcessorUPb"
'Module for handling Glitter U-Pb geochronology data outputs
Option Explicit
Dim Msg As String
Dim Ans As Variant

Private Sub GLITTERZirconArranger()

'Created By Jarred Lloyd on 03-10-2018, modified from Ben's arranger (Ben Wade)
'Feel free to modify but give credit and do not sell any version of this, modified or not. It is to remain free for those who need it
                   
Application.ScreenUpdating = False
    
    'Rename and add sheets
        ActiveSheet.Name = "Original Data"
        Sheets.Add.Name = "Rearranged Data"
        Sheets.Add.Name = "Concordia Data"
    
    Sheets("Original Data").Activate
    
    'Copy Analysis Number
        ActiveSheet.Range("A7").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("A1")
    
    'Copy Isotope Ratios
        '207/206
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("B1")
        
        '206/238
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("D1")
        
        '207/235
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("F1")
        
        '208/232
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("H1")
        
     'Copy Isotope Ratio Error(1 sigma)
        '207/206
        Sheets("Original Data").Activate
        Selection.End(xlDown).Select
        Selection.Offset(4, -3).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("C2")
              
        '206/238
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("E2")
              
        '207/235
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("G2")
        
        '208/232
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("I2")
        
    'Copy Age Estimates
        '207/206
        Sheets("Original Data").Activate
        Selection.End(xlDown).Select
        Selection.Offset(3, -3).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("K1")
        
        '206/238
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("M1")
        
        '207/235
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("O1")
        
        '208/232
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("Q1")
        
    'Copy Age Estimate Error (1 Sigma)
        '207/206
        Sheets("Original Data").Activate
        Selection.End(xlDown).Select
        Selection.Offset(4, -3).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("L2")
        
        '206/238
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("N2")
        
        '207/235
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("P2")
        
        '208/232
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("R2")
        
    'Copy CPS
        '204
        Sheets("Original Data").Activate
        Selection.End(xlDown).Select
        Selection.Offset(3, -3).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("T1")
        
        '206
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("U1")
        
        '207
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("V1")
        
        '208
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("W1")
        
        '232
        Sheets("Original Data").Activate
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("X1")
                      
        '238
        Selection.Offset(0, 1).Select
        Selection.Copy Destination:=Sheets("Rearranged Data").Range("Y1")

    
    'Copy Analysis #, 207/235, 206/238 for concordia
        'Analysis
        Sheets("Rearranged Data").Activate
        ActiveSheet.Range("A1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Concordia Data").Range("A1")
        
        '207/235
        Sheets("Rearranged Data").Activate
        ActiveSheet.Range("F1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Concordia Data").Range("B1")
        
        '207/235 (1 Sigma)
        Sheets("Rearranged Data").Activate
        ActiveSheet.Range("G2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Concordia Data").Range("C2")
        
        '206/238
        Sheets("Rearranged Data").Activate
        ActiveSheet.Range("D1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Concordia Data").Range("D1")
        
        '206/238 (1 Sigma)
        Sheets("Rearranged Data").Activate
        ActiveSheet.Range("E2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Concordia Data").Range("E2")
                
    'Calculate Rho
        Sheets("Concordia Data").Activate
        ActiveSheet.Range("F1").Value = "Rho"
        
        Sheets("Rearranged Data").Activate
            Dim LastRow As Long
                LastRow = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row
                
        ActiveSheet.Range("AA1").Value = "Rho"
        ActiveSheet.Range("AA2:AA" & LastRow).Formula = "=((((G2/F2)^2)+((E2/D2)^2)-((C2/B2)^2))*(F2*D2))/(2*G2*E2)"
        ActiveSheet.Range("AA2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Sheets("Concordia Data").Activate
        ActiveSheet.Range("F2").Select
        Selection.PasteSpecial Paste:=xlPasteValues
                
    'Copy 07/06, 06/38, 07/35 Age and Errors to concordia
        Sheets("Rearranged Data").Activate
        ActiveSheet.Range("K1:P1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Destination:=Sheets("Concordia Data").Range("H1")
    
    'Calculate Concordancies
        '07/35 06/38
        Sheets("Concordia Data").Activate
        ActiveSheet.Range("N2:N" & LastRow).Formula = "=ROUND((J2/L2)*100,0)"
        ActiveSheet.Range("N1").Value = "Concordancy [07/35][06/38]"
        
        '07/06 06/38
        ActiveSheet.Range("O2:O" & LastRow).Formula = "=ROUND((H2/J2)*100,0)"
        ActiveSheet.Range("O1").Value = "Concordancy [07/06][06/38]"
        
    'Save as xlsx
    Sheets("Concordia Data").Activate
    Sheets("Concordia Data").Range("A1").Select
    ActiveWorkbook.SaveAs Filename:=Left(Application.ActiveWorkbook.FullName, Len(Application.ActiveWorkbook.FullName) - 4), FileFormat:=51
      
Application.CutCopyMode = False
Application.ScreenUpdating = True
        
End Sub

Sub ProcessBatch()
'This will batch process a folder's CSV files, based on the host file opened
    'Variable Declaration
    Dim FolderName As String
    Dim Fname As String
    
    FolderName = Application.ActiveWorkbook.Path
    If Right(FolderName, 1) <> Application.PathSeparator Then FolderName = FolderName & Application.PathSeparator
    Fname = Dir(FolderName & "*.csv")

    'loop through the files
    Do While Len(Fname)

        With Workbooks.Open(FolderName & Fname)
           Call GLITTERZirconArranger
           ActiveWorkbook.Close
        End With
        ' Go to the next file in the folder
        Fname = Dir
    Loop
End Sub

Sub GLITTERBatchConfirm(control As IRibbonControl)
''This is to prevent accidentally ruining an excel spreadsheet

    Msg = "Do you wish to batch process GLITTER CSV output files in the host folder?"
    Ans = MsgBox(Msg, vbYesNo)
    Select Case Ans
        Case vbYes
        Call ProcessBatch
        Case vbNo
        GoTo Quit:
    End Select

Quit:
End Sub
Sub GLITTERConfirm(control As IRibbonControl)
'This is to prevent accidentally ruining an excel spreadsheet

    Msg = "Do you wish to process this GLITTER CSV output file?"
    Ans = MsgBox(Msg, vbYesNo)
    Select Case Ans
        Case vbYes
        Call GLITTERZirconArranger
        Case vbNo
        GoTo Quit:
    End Select

Quit:
End Sub

