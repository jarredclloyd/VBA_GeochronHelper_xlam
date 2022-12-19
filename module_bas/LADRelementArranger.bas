Attribute VB_Name = "LADRelementArranger"
'Module for handling LADR elemental concentration data outputs
'Last modified on 2022-12-19
'Feel free to modify but give credit and do not sell any version of this, modified or not. It is to remain free for those who need it

Option Explicit
Option Compare Text
'Variable declaration
        Dim Msg As String
        Dim Ans As Long
        Dim FolderName As String
        Dim Fname As String
    'Standards name variable
        Dim NumStandards As Long
        Dim Standard1 As String
        Dim Standard2 As String
        Dim Standard3 As String
        Dim Standard4 As String
        Dim Standard5 As String
    'Reported uncertainy level
        Dim StandardErrorLevel As Variant
    'Varibles for copying of original data
        Dim ODLastRowA As Long
        Dim ODStartRowA As Long 'includes headers
        Dim ODLastRowB As Long
        Dim ODStartRowB As Long 'no headers
        Dim ODLastCol As Long
        Dim EDNextCol As Long
    'Variables for headers in original data
        Dim FirstMassRow As String
        Dim FirstMass As String
        Dim LastMass As String
        Dim SourceFileCol As Long
        Dim SampleCol As Long
        Dim AnalysisCol As Long
        Dim ALNumCol As Long
        Dim CommentsCol As Long
        Dim ElementTotalCol As Long
        Dim EleStartCol As Variant
        Dim EleEndCol As Long
        Dim TraceElementDataPresent As Boolean
        Dim HeaderRange As Range
        Dim HeaderRangeEle As Range
        Dim CommentsColEle As Long
        Dim EleUnStartCol As Long
        Dim EleUnEndCol As Long
    'Variables for sorting
        Dim SFColDel As Long
        Dim EDLastRow As Long
        Dim EDLastCol As Long
        Dim EDSLastRow As Long
        Dim EDSLastCol As Long
        Dim EDRange As Range
        Dim EDSRange As Range
    'Variables for number formatting or rounding
        Dim cell As Range
        Dim EleConFirstCol As Long
        Dim EleConLastCol As Long
        Dim EleConUncFirstCol As Long
        Dim EleConUncLastCol As Long
    'Variables for sample and analysis label correction
        Dim SourceFile As String
        Dim Sample As String
        Dim Analysis As String
        Dim n As Long
        Dim NumberingFormat As String
    'Variables for splitting standards and unknowns (f=first row, l = last row)
        Dim Standard1f As Long
        Dim Standard1l As Long
        Dim Standard2f As Long
        Dim Standard2l As Long
        Dim Standard3f As Long
        Dim Standard3l As Long
        Dim Standard4f As Long
        Dim Standard4l As Long
        Dim Standard5f As Long
        Dim Standard5l As Long
    'Variables for formatting
        Dim LastCol As Long
        Dim LastRow As Long
        Dim ComCol As Long
Sub LADRElementalConfirm(control As IRibbonControl)
'This is to prevent accidentally running this procedure
    Msg = "This will rearrange and format a LADR CSV output file for LA-ICP-MS U-Pb elemental data." & vbCrLf & vbCrLf & "It will save the result as an XLSX with the same name as the input CSV." & vbCrLf & vbCrLf & "It will not overwrite the CSV or any existing XLSX file of the same name in the directory of the CSV." & vbCrLf & vbCrLf & "Do you want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR Elemental arranger")
    Select Case Ans
        Case vbYes
            Call LADRprocessorElemental
        Case vbNo
            Exit Sub
    End Select
End Sub
Private Sub LADRprocessorElemental()
'This procedure transforms CSV output from LADR into a human readable arrangment. This specific procedure handles elemental data.

    'Define number and name of standards
DefineStandards:                NumStandards = Application.InputBox("How many different standards were used?", "LADRElemental_Arranger", 2, Type:=1)
        Select Case NumStandards
            Case 1
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADRWetherill_Arranger", Type:=2)
            Case 2
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
            Case 3
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
            Case 4
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
            Case 5
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
                Standard5 = Application.InputBox("What is the sample name of the fifth standard as it is shown in the output CSV?", "LADRElemental_Arranger", Type:=2)
            Case Else
            Ans = MsgBox("Please enter the number of standards, this has to be a value of 1 to 4." & vbCrLf & "Do you want to continue?", vbYesNo, "LADRElemental_Arranger")
            Select Case Ans
                Case vbYes
                    GoTo DefineStandards
                Case vbNo
                    Exit Sub
            End Select
        End Select
    'Check Standards are correctly set
        Select Case NumStandards
            Case 1
                If IsEmpty(Standard1) = True Then
                    GoTo StandardCheckError
                Else
                    GoTo Start
                End If
            Case 2
                If IsEmpty(Standard1) = True Or IsEmpty(Standard2) = True Then
                    GoTo StandardCheckError
                Else
                    GoTo Start
                End If
            Case 3
                If IsEmpty(Standard1) = True Or IsEmpty(Standard2) = True Or IsEmpty(Standard3) = True Then
                    GoTo StandardCheckError
                Else: GoTo Start
                End If
            Case 4
                If IsEmpty(Standard1) = True Or IsEmpty(Standard2) = True Or IsEmpty(Standard3) = True Or IsEmpty(Standard4) = True Then
                    GoTo StandardCheckError
                Else
                    GoTo Start
                End If
            Case 5
                If IsEmpty(Standard1) = True Or IsEmpty(Standard2) = True Or IsEmpty(Standard3) = True Or IsEmpty(Standard4) = True Or IsEmpty(Standard5) = True Then
                    GoTo StandardCheckError
                Else
                    GoTo Start
                End If
        End Select
    'Standard variable declaration error handling
StandardCheckError:      Ans = MsgBox("Standards variables are not set correctly." & vbCrLf & "Do you want to continue?", vbYesNo, "LADRElemental_Arranger")
        Select Case Ans
            Case vbYes
                GoTo DefineStandards
            Case vbNo
                Exit Sub
        End Select
Start:
    'Start Timer
        Dim startTime As Double
        startTime = Timer
    'General performance optimisations
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    'Rename orignial sheet
        ActiveSheet.Name = "Original Data"
        Sheets("Original Data").Activate
    'Define standard deviation level
        StandardErrorLevel = Range("A:A").Find(what:="Reported Uncertainty", MatchCase:=True, lookat:=xlWhole).Row
        StandardErrorLevel = Cells(StandardErrorLevel, 2).Value
        StandardErrorLevel = Left(StandardErrorLevel, 1)
    'Find filtered results section and define variables
        With Sheets("Original Data")
            FirstMassRow = Range("A:A").Find(what:="Mass", MatchCase:=True, lookat:=xlWhole).Row
            FirstMassRow = FirstMassRow + 1
            FirstMass = Range("A" & FirstMassRow).Value
            LastMass = Range("B" & FirstMassRow).End(xlDown).Offset(0, -1).Value
            ODStartRowA = Range("A:A").Find(what:="FilteredConcentration_PPM", MatchCase:=True, lookat:=xlPart).Row
            ODStartRowA = ODStartRowA + 2
            ODLastRowA = Range("C" & ODStartRowA).End(xlDown).Row
            ODLastCol = Cells(ODStartRowA, Columns.Count).End(xlToLeft).Column
            Set HeaderRange = Range(Cells(ODStartRowA, 1), Cells(ODStartRowA, ODLastCol))
        End With
    'Find error section and set variables
        With Sheets("Original Data")
            ODStartRowB = Range("A:A").Find(what:="Uncertainty_PPM", MatchCase:=True, lookat:=xlPart).Row
            ODStartRowB = ODStartRowB + 3
            ODLastRowB = Range("C" & ODStartRowB).End(xlDown).Row
        End With
    'Define column variables
        With Sheets("Original Data")
            ALNumCol = HeaderRange.Find(what:="AL#", MatchCase:=False).Column
            SourceFileCol = HeaderRange.Find(what:="Source Filename", MatchCase:=False).Column
            SampleCol = HeaderRange.Find(what:="Sample", MatchCase:=False).Column
            AnalysisCol = HeaderRange.Find(what:="Analysis", MatchCase:=False).Column
            CommentsCol = HeaderRange.Find(what:="Comment", MatchCase:=False).Column
            'Check and define presence of trace element data
                Set EleStartCol = HeaderRange.Find(what:=FirstMass, MatchCase:=False)
                If Not EleStartCol Is Nothing Then
                    ElementTotalCol = HeaderRange.Find(what:="Element Total", MatchCase:=False).Column
                    EleStartCol = HeaderRange.Find(what:=FirstMass, MatchCase:=False).Column
                    EleEndCol = HeaderRange.Find(what:=LastMass, MatchCase:=False).Column
                Else
                    GoTo ElementalDataMissing
                End If
        End With
    'Add new sheets
        Sheets.Add.Name = "Elemental Data"
        Sheets("Original Data").Activate
    'Copy AL#, sample name and analysis number, and element total
        Range(Cells(ODStartRowA, ALNumCol), Cells(ODLastRowA, ALNumCol)).Copy Destination:=Sheets("Elemental Data").Range("A1")
        Sheets("Elemental Data").Range("A1").Value = "ALnum"
        Range(Cells(ODStartRowA, SampleCol), Cells(ODLastRowA, AnalysisCol)).Copy Destination:=Sheets("Elemental Data").Range("B1")
    'Copy elemental data and uncertainties
        EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
        EDNextCol = EDLastCol + 1
        Range(Cells(ODStartRowA, ElementTotalCol), Cells(ODLastRowA, ElementTotalCol)).Copy Destination:=Sheets("Elemental Data").Cells(1, EDNextCol)
        EDNextCol = EDNextCol + 1
        For n = EleStartCol To EleEndCol
            Range(Cells(ODStartRowA, n), Cells(ODLastRowA, n)).Copy Destination:=Sheets("Elemental Data").Cells(1, EDNextCol)
            EDNextCol = EDNextCol + 1
            Sheets("Elemental Data").Cells(1, EDNextCol).Value = Sheets("Elemental Data").Cells(1, EDNextCol - 1).Value & "_" & StandardErrorLevel & "SE"
            Range(Cells(ODStartRowB, n), Cells(ODLastRowB, n)).Copy Destination:=Sheets("Elemental Data").Cells(2, EDNextCol)
            EDNextCol = EDNextCol + 1
        Next n
        EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
        EDNextCol = EDLastCol + 1
    Sheets("Original Data").Activate 'copying comments, source filename, elemental uncertainties
    'Copy Comments
        Range(Cells(ODStartRowA, CommentsCol), Cells(ODLastRowA, CommentsCol)).Copy Destination:=Sheets("Elemental Data").Cells(1, EDNextCol)
        EDNextCol = EDNextCol + 1
    'Copy Source Filename
        Range(Cells(ODStartRowA, SourceFileCol), Cells(ODLastRowA, SourceFileCol)).Copy Destination:=Sheets("Elemental Data").Cells(1, EDNextCol)
        SFColDel = EDNextCol
    Sheets("Elemental Data").Activate 'Correct sample label and trailing number for correct sorting in Excel
            EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
            EDLastRow = Sheets("Elemental Data").Cells(Rows.Count, 1).End(xlUp).Row
            If Left(Cells(2, SFColDel).Value, 2) = "1-" Or Left(Cells(2, SFColDel).Value, 2) = "1 -" Then
                NumberingFormat = "NewWave"
                Else
                NumberingFormat = "GeoStar"
            End If
            Select Case NumberingFormat
                Case "GeoStar"
                    For n = 2 To EDLastRow
                        SourceFile = Cells(n, SFColDel).Value
                        SourceFile = Left(SourceFile, InStrRev(SourceFile, ".") - 1)
                        If Mid(SourceFile, InStrRev(SourceFile, "-") - 1, 3) = " - " Then
                            Sample = Left(SourceFile, InStrRev(SourceFile, "-") - 2)
                        Else
                            Sample = Left(SourceFile, InStrRev(SourceFile, "-") - 1)
                        End If
                        Range("B" & n).Value = Sample
                        Analysis = Right(SourceFile, Len(SourceFile) - Len(Sample) - 2)
                        Analysis = Format(Analysis, "000")
                        Range("C" & n).Value = Sample & "-" & Analysis
                    Next n
                Case "NewWave"
                    For n = 2 To EDLastRow
                        SourceFile = Cells(n, SFColDel).Value
                        SourceFile = Left(SourceFile, InStrRev(SourceFile, ".") - 1)
                        If Mid(SourceFile, InStr(1, SourceFile, "-") - 1, 3) = " - " Then
                            Sample = Right(SourceFile, Len(SourceFile) - (InStr(1, SourceFile, "-") + 1))
                        Else
                            Sample = Right(SourceFile, Len(SourceFile) - InStr(1, SourceFile, "-"))
                        End If
                        Range("B" & n).Value = Sample
                        If Mid(SourceFile, InStr(1, SourceFile, "-") - 1, 3) = " - " Then
                            Analysis = Left(SourceFile, InStr(1, SourceFile, "-") - 2)
                        Else
                            Analysis = Left(SourceFile, InStr(1, SourceFile, "-") - 1)
                        End If
                        Analysis = Format(Analysis, "000")
                        Range("C" & n).Value = Analysis & "-" & Sample
                    Next n
            End Select
    'Sort prior to cut - performance optimisation
        'Delete sourcefile column
            Columns(SFColDel).Delete
        'Set ED Last Column and sort elemental data
            EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
            EDLastRow = Sheets("Elemental Data").Cells(Rows.Count, 1).End(xlUp).Row
            Set EDRange = Sheets("Elemental Data").Range("A1", Cells(EDLastRow, EDLastCol))
            With EDRange
                .Sort Key1:=Range("B1"), order1:=xlAscending, Header:=xlYes
            End With
     'Separate Standards and Unknowns
        'Add standards worksheets
            Sheets.Add.Name = "Elemental Data - Standards"
            Sheets("Elemental Data - Standards").Move after:=Sheets("Elemental Data")
        'Define ED variables
            EDLastRow = Sheets("Elemental Data").Range("A" & Rows.Count).End(xlUp).Row
            EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
    'Cut standards from elemental data
        Sheets("Elemental Data").Activate
        Range("A1", Cells(1, EDLastCol)).Copy Destination:=Sheets("Elemental Data - Standards").Range("A1")
        'Define variables for cut ranges
            With ActiveSheet
                On Error Resume Next
                Standard1f = .Range("B:B").Find(what:=Standard1).Row
                On Error Resume Next
                Standard1l = .Range("B:B").Find(what:=Standard1, searchdirection:=xlPrevious).Row
                On Error Resume Next
                Standard2f = .Range("B:B").Find(what:=Standard2).Row
                On Error Resume Next
                Standard2l = .Range("B:B").Find(what:=Standard2, searchdirection:=xlPrevious).Row
                On Error Resume Next
                Standard3f = .Range("B:B").Find(what:=Standard3).Row
                On Error Resume Next
                Standard3l = .Range("B:B").Find(what:=Standard3, searchdirection:=xlPrevious).Row
                On Error Resume Next
                Standard4f = .Range("B:B").Find(what:=Standard4).Row
                On Error Resume Next
                Standard4l = .Range("B:B").Find(what:=Standard4, searchdirection:=xlPrevious).Row
                On Error Resume Next
                Standard5f = .Range("B:B").Find(what:=Standard5).Row
                On Error Resume Next
                Standard5l = .Range("B:B").Find(what:=Standard5, searchdirection:=xlPrevious).Row
            End With
        'Cut Standard1
            On Error Resume Next
            Range("A" & Standard1f, Cells(Standard1l, EDLastCol)).Cut Destination:=Sheets("Elemental Data - Standards").Range("A" & Standard1f)
        'Cut Standard2
            On Error Resume Next
            Range("A" & Standard3f, Cells(Standard3l, EDLastCol)).Cut Destination:=Sheets("Elemental Data - Standards").Range("A" & Standard3f)
        'Cut  Standard3
            On Error Resume Next
            Range("A" & Standard2f, Cells(Standard2l, EDLastCol)).Cut Destination:=Sheets("Elemental Data - Standards").Range("A" & Standard2f)
        'Cut  Standard4
            On Error Resume Next
            Range("A" & Standard4f, Cells(Standard4l, EDLastCol)).Cut Destination:=Sheets("Elemental Data - Standards").Range("A" & Standard4f)
        'Cut  Standard5
            On Error Resume Next
            Range("A" & Standard5f, Cells(Standard5l, EDLastCol)).Cut Destination:=Sheets("Elemental Data - Standards").Range("A" & Standard5f)
        'Sort after cut
            'Redefine elemental data last rows and columns
                EDLastRow = Sheets("Elemental Data").Range("A" & Rows.Count).End(xlUp).Row
                EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
                EDSLastRow = Sheets("Elemental Data - Standards").Range("A" & Rows.Count).End(xlUp).Row
                EDSLastCol = Sheets("Elemental Data - Standards").Cells(1, Columns.Count).End(xlToLeft).Column
            'Elemental Data
                Set EDRange = Sheets("Elemental Data").Range("A1", Cells(EDLastRow, EDLastCol))
                With EDRange
                    .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
                End With
            'Elemental Data - Standards
                Sheets("Elemental Data - Standards").Activate
                Set EDSRange = Sheets("Elemental Data - Standards").Range("A1", Cells(EDSLastRow, EDSLastCol))
                With EDSRange
                    .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
                End With
        'Rename unknowns sheet
            Sheets("Elemental Data").Name = "Elemental Data - Unknowns"
    'Format and reset performance optimisations
        'Enable screen updating for pane freeze
            Application.ScreenUpdating = True
            Application.EnableEvents = True
        'Elemental Data Standards
                    Sheets("Elemental Data - Standards").Activate
                    With ActiveSheet
                        Cells.NumberFormat = "General"
                        LastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
                        LastRow = .Range("A" & Rows.Count).End(xlUp).Row
                        .Range("A1", Cells(1, LastCol)).Columns.AutoFit
                        .Range("A:C").Columns.AutoFit
                        ComCol = Range("A1", Cells(1, LastCol)).Find(what:="Comment", MatchCase:=False).Column
                        .Range(Cells(1, ComCol), Cells(LastRow, ComCol)).Columns.AutoFit
                        .Range("A1", Cells(1, LastCol)).Font.Bold = True
                        .Range("A1", Cells(LastRow, LastCol)).Font.Size = 8
                    End With
                    With ActiveWindow
                        .SplitColumn = 3
                        .SplitRow = 1
                        .FreezePanes = True
                    End With
                'Elemental Data Unknowns
                    Sheets("Elemental Data - Unknowns").Activate
                    With ActiveSheet
                        Cells.NumberFormat = "General"
                        LastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
                        LastRow = .Range("A" & Rows.Count).End(xlUp).Row
                        .Range("A1", Cells(1, LastCol)).Columns.AutoFit
                        .Range("A:C").Columns.AutoFit
                        ComCol = Range("A1", Cells(1, LastCol)).Find(what:="Comment", MatchCase:=False).Column
                        .Range(Cells(1, ComCol), Cells(LastRow, ComCol)).Columns.AutoFit
                        .Range("A1", Cells(1, LastCol)).Font.Bold = True
                        .Range("A1", Cells(LastRow, LastCol)).Font.Size = 8
                    End With
                    With ActiveWindow
                        .SplitColumn = 3
                        .SplitRow = 1
                        .FreezePanes = True
                    End With
'Reset calculation status and trigger calculation
        Application.Calculation = xlCalculationAutomatic
    'Display elapsed time
        MsgBox "Completed in " & Format(Timer - startTime, "00.00") & " seconds"
    'Save as xlsx
        On Error GoTo SaveError
        ActiveWorkbook.SaveAs FileName:=Left(Application.ActiveWorkbook.FullName, Len(Application.ActiveWorkbook.FullName) - 4) & ".xlsx", FileFormat:=51, ConflictResolution:=xlOtherSessionChanges
        Exit Sub
    'Explicit error handling
ElementalDataMissing: MsgBox "Elemental data is missing from this sheet.", , "LADR Elemental Arranger"
    Exit Sub
SaveError:     MsgBox "There was an error saving this file during execution of the add-in." & vbCrLf & "This is likely due to a file of the same name already existing" & vbCrLf & vbCrLf & "Please remember to save your file if you haven't already", , "LADR Elemental Arranger"
End Sub
