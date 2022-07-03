Attribute VB_Name = "LADRprocessorSIsotopes"
'Module for handling LADR S isotope ratio data outputs
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
        Dim UncertaintyLevel As Variant
        Dim StandardErrorLevel As Variant
    'Varibles for copying of original data
        Dim ODLastRowA As Long
        Dim ODStartRowA As Long
        Dim ODLastRowB As Long
        Dim ODStartRowB As Long
        Dim ODLastCol As Long
        Dim RDNextCol As Long
        Dim EDNextCol As Long
    'Variables for headers in original data
        Dim FMrow As String
        Dim FirstMass As String
        Dim LastMass As String
        Dim SourceFileCol As Long
        Dim SampleCol As Long
        Dim AnalysisCol As Long
        Dim ALNumCol As Long
        Dim CommentsCol As Long
        Dim EleStartCol As Variant
        Dim EleEndCol As Long
        Dim TraceElementDataPresent As Boolean
        Dim RatioS34S32Col As Variant
        Dim HeaderRange As Range
        Dim HeaderRangeEle As Range
        Dim CommentsColEle As Long
        Dim EleUnStartCol As Long
        Dim EleUnEndCol As Long
    'Variables for sorting
        Dim SFColDel As Long
        Dim RDLastRow As Long
        Dim RDLastCol As Long
        Dim EDLastRow As Long
        Dim EDLastCol As Long
        Dim RDSLastRow As Long
        Dim RDSLastCol As Long
        Dim EDSLastRow As Long
        Dim EDSLastCol As Long
        Dim EDRange As Range
        Dim EDSRange As Range
        Dim RDRange As Range
        Dim RDSRange As Range
    'Variables for age/ratio rounding
        Dim cell As Range
        Dim RDAgeI As Long
        Dim RDAgeF As Long
        Dim RatioFirstCol As Long
        Dim RatioLastCol As Long
    'Variables for sample and analysis label correction
        Dim SourceFile As String
        Dim Sample As String
        Dim Analysis As String
        Dim n As Long
    'Variables for splitting standards and unknowns
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
Sub LADRSIsotopeArranger()
'Created By Jarred Lloyd on 2022-06-23
'Last modified on 2022-06-23
'Feel free to modify but give credit and do not sell any version of this, modified or not. It is to remain free for those who need it

    'Define number and name of standards
DefineStandards:            NumStandards = Application.InputBox("How many different standards were used?", "LADR_S_Isotope_Arranger", 4, Type:=1)
        Select Case NumStandards
            Case 1
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
            Case 2
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
            Case 3
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
            Case 4
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
            Case 5
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
                Standard5 = Application.InputBox("What is the sample name of the fifth standard as it is shown in the output CSV?", "LADR_S_Isotope_Arranger", Type:=2)
            Case Else
                Ans = MsgBox("Please enter the number of standards, this has to be a value of 1 to 5." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_S_Isotope_Arranger")
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
                    GoTo StdCheckError
                Else
                    GoTo Start
                End If
            Case 2
                If IsEmpty(Standard1) = True Or IsEmpty(Standard2) = True Then
                    GoTo StdCheckError
                Else
                    GoTo Start
                End If
            Case 3
                If IsEmpty(Standard1) = True Or IsEmpty(Standard2) = True Or IsEmpty(Standard3) = True Then
                    GoTo StdCheckError
                Else
                    GoTo Start
                End If
            Case 4
                If IsEmpty(Standard1) = True Or IsEmpty(Standard2) = True Or IsEmpty(Standard3) = True Or IsEmpty(Standard4) = True Then
                    GoTo StdCheckError
                Else
                    GoTo Start
                End If
            Case 5
                If IsEmpty(Standard1) = True Or IsEmpty(Standard2) = True Or IsEmpty(Standard3) = True Or IsEmpty(Standard4) = True Or IsEmpty(Standard5) = True Then
                    GoTo StdCheckError
                Else
                    GoTo Start
                End If
        End Select
    'Standard variable declaration error handling
StdCheckError:      Ans = MsgBox("Standards variables are not set correctly." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_S_Isotope_Arranger")
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
            FMrow = Range("A:A").Find(what:="Mass", MatchCase:=True, lookat:=xlWhole).Row
            FMrow = FMrow + 1
            FirstMass = Range("A" & FMrow).Value
            LastMass = Range("B" & FMrow).End(xlDown).Offset(0, -1).Value
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
            RatioS34S32Col = HeaderRange.Find(what:="34S->66/32S->64", MatchCase:=False).Column
            'Check and define presence of trace element data
                Set EleStartCol = HeaderRange.Find(what:=FirstMass, MatchCase:=False)
                If Not EleStartCol Is Nothing Then
                    EleStartCol = HeaderRange.Find(what:=FirstMass, MatchCase:=False).Column
                    EleEndCol = HeaderRange.Find(what:=LastMass, MatchCase:=False).Column
                    TraceElementDataPresent = True
                Else
                    TraceElementDataPresent = False
                End If
        End With
    'Add new sheets
        Sheets.Add.Name = "Ratio Data"
        Select Case TraceElementDataPresent
            Case True
                Sheets.Add.Name = "Elemental Data"
            Case False
        End Select
        Sheets("Original Data").Activate
    'Copy AL#, sample name and analysis number
        Range(Cells(ODStartRowA, ALNumCol), Cells(ODLastRowA, ALNumCol)).Copy Destination:=Sheets("Ratio Data").Range("A1")
        Range(Cells(ODStartRowA, SampleCol), Cells(ODLastRowA, AnalysisCol)).Copy Destination:=Sheets("Ratio Data").Range("B1")
        Select Case TraceElementDataPresent
            Case True
                Range(Cells(ODStartRowA, ALNumCol), Cells(ODLastRowA, ALNumCol)).Copy Destination:=Sheets("Elemental Data").Range("A1")
                Range(Cells(ODStartRowA, SampleCol), Cells(ODLastRowA, AnalysisCol)).Copy Destination:=Sheets("Elemental Data").Range("B1")
            Case False
        End Select
    'Copy Ratio Data, suitable for isoplotR input
        'Define RDNextCol, RDLastRow
            RDLastCol = Sheets("Ratio Data").Cells(1, Columns.Count).End(xlToLeft).Column
            RDNextCol = RDLastCol + 1
            RDLastRow = Sheets("Ratio Data").Cells(Rows.Count, 1).End(xlUp).Row
        'Copy ratios and uncertainties, label uncertainty columns
        'S34/S32
            Range(Cells(ODStartRowA, RatioS34S32Col), Cells(ODLastRowA, RatioS34S32Col)).Copy Destination:=Sheets("Ratio Data").Cells(1, RDNextCol)
            Sheets("Ratio Data").Cells(1, RDNextCol).Value = "S34/S32"
            RDNextCol = RDNextCol + 1
            Sheets("Ratio Data").Cells(1, RDNextCol).Value = "Uncertainty[S34/S32] " & StandardErrorLevel & "SE"
            Range(Cells(ODStartRowB, RatioS34S32Col), Cells(ODLastRowB, RatioS34S32Col)).Copy Destination:=Sheets("Ratio Data").Cells(2, RDNextCol)
            RDNextCol = RDNextCol + 1
        'Copy elemental concentrations and uncertainty data
            Select Case TraceElementDataPresent
                Case True
                    EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
                    EDNextCol = EDLastCol + 1
                    For n = EleStartCol To EleEndCol
                        Range(Cells(ODStartRowA, n), Cells(ODLastRowA, n)).Copy Destination:=Sheets("Elemental Data").Cells(1, EDNextCol)
                        EDNextCol = EDNextCol + 1
                        Sheets("Elemental Data").Cells(1, EDNextCol).Value = Sheets("Elemental Data").Cells(1, EDNextCol - 1).Value & " " & StandardErrorLevel & "SE"
                        Range(Cells(ODStartRowB, n), Cells(ODLastRowB, n)).Copy Destination:=Sheets("Elemental Data").Cells(2, EDNextCol)
                        EDNextCol = EDNextCol + 1
                    Next n
                    EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
                    EDNextCol = EDLastCol + 1
                Case False
            End Select
        'Copy Comments
            Range(Cells(ODStartRowA, CommentsCol), Cells(ODLastRowA, CommentsCol)).Copy Destination:=Sheets("Ratio Data").Cells(1, RDNextCol)
            RDNextCol = RDNextCol + 1
            Select Case TraceElementDataPresent
                Case True
                    Range(Cells(ODStartRowA, CommentsCol), Cells(ODLastRowA, CommentsCol)).Copy Destination:=Sheets("Elemental Data").Cells(1, EDNextCol)
                    EDNextCol = EDNextCol + 1
                Case False
            End Select
        'Copy Source Filename
            Range(Cells(ODStartRowA, SourceFileCol), Cells(ODLastRowA, SourceFileCol)).Copy Destination:=Sheets("Ratio Data").Cells(1, RDNextCol)
            SFColDel = RDNextCol
    
    Sheets("Ratio Data").Activate 'Correct sample label and trailing number for correct sorting in Excel
            RDLastCol = Sheets("Ratio Data").Cells(1, Columns.Count).End(xlToLeft).Column
            RDLastRow = Sheets("Ratio Data").Cells(Rows.Count, 1).End(xlUp).Row
            For n = 2 To RDLastRow
                SourceFile = Cells(n, SFColDel).Value
                SourceFile = Left(SourceFile, InStrRev(SourceFile, ".") - 1)
                Sample = Left(SourceFile, InStrRev(SourceFile, "-") - 2)
                Range("B" & n).Value = Sample
                Analysis = Right(SourceFile, Len(SourceFile) - Len(Sample) - 2)
                Analysis = Format(Analysis, "000")
                Range("C" & n).Value = Sample & " - " & Analysis
            Next n
        'Copy corrected sample and analysis labels to elemental data
           Select Case TraceElementDataPresent
                Case True
                    Range("B1", "C" & RDLastRow).Copy Destination:=Sheets("Elemental Data").Range("B1")
                Case False
            End Select
    'Sort prior to cut - performance optimisation
        'Delete sourcefile column
            Columns(SFColDel).Delete
        'Set GD Last Column
            RDLastCol = Sheets("Ratio Data").Cells(1, Columns.Count).End(xlToLeft).Column
        'Ratio Data
            Set RDRange = Sheets("Ratio Data").Range("A1", Cells(RDLastRow, RDLastCol))
                With RDRange
                .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
            End With
        'Elemental Data
            Select Case TraceElementDataPresent
            Case True
                Sheets("Elemental Data").Activate
                EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
                EDLastRow = Sheets("Elemental Data").Cells(Rows.Count, 1).End(xlUp).Row
                Set EDRange = Sheets("Elemental Data").Range("A1", Cells(EDLastRow, EDLastCol))
                With EDRange
                    .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
                End With
            Case False
            End Select
    'Separate Standards and Unknowns
        'Add standards worksheets
            Sheets.Add.Name = "Ratio Data - Standards"
            Select Case TraceElementDataPresent
                Case True
                    Sheets.Add.Name = "Elemental Data - Standards"
                    Sheets("Elemental Data - Standards").Move after:=Sheets("Elemental Data")
                Case False
            End Select
        'Define GD and ED variables
            RDLastRow = Sheets("Ratio Data").Range("A" & Rows.Count).End(xlUp).Row
            RDLastCol = Sheets("Ratio Data").Cells(1, Columns.Count).End(xlToLeft).Column
            Select Case TraceElementDataPresent
                Case True
                    EDLastRow = Sheets("Elemental Data").Range("A" & Rows.Count).End(xlUp).Row
                    EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
                Case False
            End Select
        Sheets("Ratio Data").Activate 'Cut standards from Ratio Data
            Range("A1", Cells(1, RDLastCol)).Copy Destination:=Sheets("Ratio Data - Standards").Range("A1") 'copy headers
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
                Range("A" & Standard1f, Cells(Standard1l, RDLastCol)).Cut Destination:=Sheets("Ratio Data - Standards").Range("A" & Standard1f)
            'Cut Standard2
                On Error Resume Next
                Range("A" & Standard3f, Cells(Standard3l, RDLastCol)).Cut Destination:=Sheets("Ratio Data - Standards").Range("A" & Standard3f)
            'Cut  Standard3
                On Error Resume Next
                Range("A" & Standard2f, Cells(Standard2l, RDLastCol)).Cut Destination:=Sheets("Ratio Data - Standards").Range("A" & Standard2f)
            'Cut  Standard4
                On Error Resume Next
                Range("A" & Standard4f, Cells(Standard4l, RDLastCol)).Cut Destination:=Sheets("Ratio Data - Standards").Range("A" & Standard4f)
            'Cut  Standard5
                On Error Resume Next
                Range("A" & Standard5f, Cells(Standard5l, RDLastCol)).Cut Destination:=Sheets("Ratio Data - Standards").Range("A" & Standard5f)
            'Sort after cut
                'Redefine geochronology last rows and columnss
                    RDLastRow = Sheets("Ratio Data").Range("A" & Rows.Count).End(xlUp).Row
                    RDLastCol = Sheets("Ratio Data").Cells(1, Columns.Count).End(xlToLeft).Column
                    RDSLastRow = Sheets("Ratio Data - Standards").Range("A" & Rows.Count).End(xlUp).Row
                    RDSLastCol = Sheets("Ratio Data - Standards").Cells(1, Columns.Count).End(xlToLeft).Column
                'Ratio Data
                    Set RDRange = Sheets("Ratio Data").Range("A1", Cells(RDLastRow, RDLastCol))
                    With RDRange
                        .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
                    End With
                'Ratio Data - Standards
                    Sheets("Ratio Data - Standards").Activate
                    Set RDSRange = Sheets("Ratio Data - Standards").Range("A1", Cells(RDSLastRow, RDSLastCol))
                    With RDSRange
                        .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
                    End With
            'Rename unknowns sheet
                Sheets("Ratio Data").Name = "Ratio Data - Unknowns"
                Sheets("Ratio Data - Unknowns").Move after:=Sheets("Ratio Data - Standards")
        'Cut standards from elemental data
            Select Case TraceElementDataPresent
                Case True
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
                        Sheets("Elemental Data - Unknowns").Move after:=Sheets("Elemental Data - Standards")
                Case False
            End Select
    'Format and reset performance optimisations
        'Enable screen updating for pane freeze
            Application.ScreenUpdating = True
            Application.EnableEvents = True
        'Format
            Select Case TraceElementDataPresent 'Elemental Data
                Case True
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
                Case False
            End Select
            'Ratio Data Standards
                Sheets("Ratio Data - Standards").Activate
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
            'Ratio Data Unknowns
                Sheets("Ratio Data - Unknowns").Activate
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
SaveError:                     MsgBox "There was an error saving this file during execution of the add-in." & vbCrLf & "This is likley due to a file of the same name already existing" & vbCrLf & vbCrLf & "Please remember to save your file if you haven't already", , "LADR_S_Isotope_Arranger"
End Sub
Sub LADRSIsotopeArrangerConfirm(control As IRibbonControl)
'This is to prevent accidentally running this procedure
    Msg = "This will rearrange and format a LADR CSV output file for QQQ S Isotope quantification and elemental data." & vbCrLf & vbCrLf & "It will save the result as an XLSX with the same name as the input CSV." & vbCrLf & vbCrLf & "It will not overwrite the CSV or any existing XLSX file of the same name in the directory of the CSV." & vbCrLf & vbCrLf & "Do you want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR_S_Isotope_Arranger")
    Select Case Ans
        Case vbYes
        Call LADRSIsotopeArranger
        Case vbNo
        GoTo Quit:
    End Select
Quit:
End Sub
