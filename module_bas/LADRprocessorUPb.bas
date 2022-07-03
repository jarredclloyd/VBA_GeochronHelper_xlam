Attribute VB_Name = "LADRprocessorUPb"
'Module for handling LADR U-Pb geochronology data outputs
'Created By Jarred Lloyd on 2020-03-08
'Last modified on 2022-07-02
'Feel free to modify but give credit and do not sell any version of this, modified or not. It is to remain free for those who need it

Option Explicit
Option Compare Text
'Variable declaration
        Dim Msg As String
        Dim Ans As Long
        Dim ConcordiaPlotType As String
        Dim FolderName As String
        Dim Fname As String
    'Standards name variable
        Dim NumStandards As Long
        Dim Standard1 As String
        Dim Standard2 As String
        Dim Standard3 As String
        Dim Standard4 As String
        Dim Standard5 As String
    'Error correlation calculation
        Dim SignalPrecision As Boolean
        Dim RhoCalc As Boolean
        Dim Rho206Pb238Uvs207Pb235UCol As Variant
        Dim Rho207Pb206Pbvs238U206PbCol As Variant
    'Reported uncertainy level
        Dim UncertaintyLevel As Variant
        Dim StandardErrorLevel As Variant
    'Varibles for copying of original data
        Dim ODLastRowA As Long
        Dim ODStartRowA As Long 'includes headers
        Dim ODLastRowB As Long
        Dim ODStartRowB As Long 'no headers
        Dim ODLastCol As Long
        Dim GDNextCol As Long
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
        Dim EleStartCol As Variant
        Dim EleEndCol As Long
        Dim TraceElementDataPresent As Boolean
        Dim Ratio208232Col As Variant
        Dim Ratio208206Col As Variant
        Dim Ratio207206Col As Variant
        Dim Ratio204206Col As Variant
        Dim Ratio207235calcCol As Variant
        Dim Ratio207235Col As Variant
        Dim Ratio206238Col As Variant
        Dim Ratio238206Col As Variant
        Dim AgeEstimate208232Col As Variant
        Dim AgeEstimate207206Col As Variant
        Dim AgeEstimate207235calcCol As Variant
        Dim AgeEstimate207235Col As Variant
        Dim AgeEstimate206238Col As Variant
        Dim AgeEstimate238206Col As Variant
        Dim HeaderRange As Range
        Dim HeaderRangeEle As Range
        Dim CommentsColEle As Long
        Dim EleUnStartCol As Long
        Dim EleUnEndCol As Long
    'Variables for checking ratio/age presence
        Dim Check208206 As Boolean
        Dim Check204206 As Boolean
        Dim Check238206 As Boolean
        Dim Check207235calc As Boolean
        Dim Check207235 As Boolean
        Dim Check207206 As Boolean
        Dim Check206238 As Boolean
        Dim Check208232 As Boolean
    'Variables for sorting
        Dim SFColDel As Long
        Dim GDLastRow As Long
        Dim GDLastCol As Long
        Dim EDLastRow As Long
        Dim EDLastCol As Long
        Dim GDSLastRow As Long
        Dim GDSLastCol As Long
        Dim EDSLastRow As Long
        Dim EDSLastCol As Long
        Dim EDRange As Range
        Dim EDSRange As Range
        Dim GDRange As Range
        Dim GDSRange As Range
    'Variables for concordance(logratio distance)
        Dim RCIndexRatio207206 As Long
        Dim RCIndexRatio238206 As Long
        Dim RCIndexAge206238 As Long
        Dim RCIndexAge207206 As Long
    'Variables for number formatting or rounding
        Dim cell As Range
        Dim AgeEstimateFirstCol As Long
        Dim AgeEstimateLastCol As Long
        Dim RatioFirstCol As Long
        Dim RatioLastCol As Long
        Dim EleConFirstCol As Long
        Dim EleConLastCol As Long
        Dim EleConUncFirstCol As Long
        Dim EleConUncLastCol As Long
    'Variables for sample and analysis label correction
        Dim SourceFile As String
        Dim Sample As String
        Dim Analysis As String
        Dim n As Long
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
Sub LADRUPbWetherillBatchProcess()
'This will batch process a folder's CSV files, based on the host file opened
    'Variable declaration
    Dim FolderName As String
    Dim Fname As String
    FolderName = Application.ActiveWorkbook.Path
    If Right(FolderName, 1) <> Application.PathSeparator Then FolderName = FolderName & Application.PathSeparator
    Fname = Dir(FolderName & "*.csv")
    'loop through the files
    Do While Len(Fname)
        With Workbooks.Open(FolderName & Fname)
            ConcordiaPlotType = "Wetherill"
            Call LADRprocessorUPb
            ActiveWorkbook.Close
        End With
        ' Go to the next file in the folder
        Fname = Dir
    Loop
End Sub
Sub LADRUPbTeraWasserburgBatchProcess()
'This will batch process a folder's CSV files, based on the host file opened
    FolderName = Application.ActiveWorkbook.Path
    If Right(FolderName, 1) <> Application.PathSeparator Then FolderName = FolderName & Application.PathSeparator
    Fname = Dir(FolderName & "*.csv")
    'loop through the files
    Do While Len(Fname)
        With Workbooks.Open(FolderName & Fname)
           ConcordiaPlotType = "TeraWasserburg"
           Call LADRprocessorUPb
           ActiveWorkbook.Close
        End With
        ' Go to the next file in the folder
        Fname = Dir
    Loop
End Sub
Sub LADRUPbWetherillBatchConfirm(control As IRibbonControl)
'This is to prevent accidentally ruining an excel spreadsheet
    Msg = "Do you wish to batch process all CSV output files (LADR-with elemental data) in the host folder as Tera-Wasserburg format?" & vbCrLf & vbCrLf & "It is a good idea to move the CSV files you want to process into a separate folder as this will process all CSV files in the host folder and result in execution errors" & vbCrLf & vbCrLf & "Continue?"
    Ans = MsgBox(Msg, vbYesNo)
    Select Case Ans
        Case vbYes
            Call LADRUPbWetherillBatchProcess
        Case vbNo
            Exit Sub
    End Select
End Sub
Sub LADRUPbTeraWasserburgBatchConfirm(control As IRibbonControl)
'This is to prevent accidentally ruining an excel spreadsheet
    Msg = "Do you wish to batch process all CSV output files (LADR-with elemental data) in the host folder as Tera-Wasserburg format?" & vbCrLf & vbCrLf & "It is a good idea to move the CSV files you want to process into a separate folder as this will process all CSV files in the host folder and result in execution errors" & vbCrLf & vbCrLf & "Continue?"
    Ans = MsgBox(Msg, vbYesNo)
    Select Case Ans
        Case vbYes
            Call LADRUPbTeraWasserburgBatchProcess
        Case vbNo
            Exit Sub
    End Select
End Sub
Sub LADRUPbTeraWasserburgConfirm(control As IRibbonControl)
'This is to prevent accidentally running this procedure
    Msg = "This will rearrange and format a LADR CSV output file for LA-ICP-MS U-Pb geochronology and trace element data in Tera-Wasserburg format." & vbCrLf & vbCrLf & "It will save the result as an XLSX with the same name as the input CSV." & vbCrLf & vbCrLf & "It will not overwrite the CSV or any existing XLSX file of the same name in the directory of the CSV." & vbCrLf & vbCrLf & "Do you want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR U-Pb arranger:TeraWasserburg")
    Select Case Ans
        Case vbYes
            ConcordiaPlotType = "TeraWasserburg"
            Call LADRprocessorUPb
        Case vbNo
            Exit Sub
    End Select
End Sub
Sub LADRUPbWetherillConfirm(control As IRibbonControl)
'This is to prevent accidentally running this procedure
    Msg = "This will rearrange and format a LADR CSV output file for LA-ICP-MS U-Pb geochronology and trace element data in Wetherill format." & vbCrLf & vbCrLf & "It will save the result as an XLSX with the same name as the input CSV." & vbCrLf & vbCrLf & "It will not overwrite the CSV or any existing XLSX file of the same name in the directory of the CSV." & vbCrLf & vbCrLf & "Do you want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR U-Pb arranger:Wetherill")
    Select Case Ans
        Case vbYes
            ConcordiaPlotType = "Wetherill"
            Call LADRprocessorUPb
        Case vbNo
            Exit Sub
    End Select
End Sub
Private Sub LADRprocessorUPb()
'This procedure transforms CSV output from LADR into a human readable, IsoplotR interoperable arrangment. This specific procedure handles U-Pb data, dynamically determining presence of trace element data, ratio and age estimates, error correlations (or estimates if signal precision, and three system calculation is possible)

    'Define number and name of standards
DefineStandards:                NumStandards = Application.InputBox("How many different standards were used?", "LADR_Wetherill_Arranger", 4, Type:=1)
        Select Case NumStandards
            Case 1
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
            Case 2
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
            Case 3
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
            Case 4
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
            Case 5
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
                Standard5 = Application.InputBox("What is the sample name of the fifth standard as it is shown in the output CSV?", "LADR_Wetherill_Arranger", Type:=2)
            Case Else
            Ans = MsgBox("Please enter the number of standards, this has to be a value of 1 to 4." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_Wetherill_Arranger")
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
StandardCheckError:      Ans = MsgBox("Standards variables are not set correctly." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_Wetherill_Arranger")
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
    'Determine if data is signal precision (for estimation of error correlation)
        UncertaintyLevel = Range("A:A").Find(what:="Reported Uncertainty Level", MatchCase:=True, lookat:=xlWhole).Row
        UncertaintyLevel = Cells(UncertaintyLevel, 2).Value
        If UncertaintyLevel = "Signal Precision" Then
            SignalPrecision = True
        Else
            SignalPrecision = False
        End If
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
                    EleStartCol = HeaderRange.Find(what:=FirstMass, MatchCase:=False).Column
                    EleEndCol = HeaderRange.Find(what:=LastMass, MatchCase:=False).Column
                    TraceElementDataPresent = True
                Else
                    TraceElementDataPresent = False
                End If
            'Check and define U238/Pb206 ratio
                Set Ratio238206Col = HeaderRange.Find(what:="238U/206Pb", MatchCase:=False, lookat:=xlWhole)
                If Not Ratio238206Col Is Nothing Then
                    Ratio238206Col = HeaderRange.Find(what:="238U/206Pb", MatchCase:=False).Column
                    Check238206 = True
                Else
                    Check238206 = False
                End If
            'Check and define Pb208/Pb206 ratio
                Set Ratio208206Col = HeaderRange.Find(what:="208Pb/206Pb", MatchCase:=False, lookat:=xlWhole)
                If Not Ratio208206Col Is Nothing Then
                    Ratio208206Col = HeaderRange.Find(what:="208Pb/206Pb", MatchCase:=False).Column
                    Check208206 = True
                Else
                    Check208206 = False
                End If
            'Check and define Pb207/U235 (calc) ratio
                Set Ratio207235calcCol = HeaderRange.Find(what:="207Pb/235U(calc)", MatchCase:=False, lookat:=xlWhole)
                    If Not Ratio207235calcCol Is Nothing Then
                        Ratio207235calcCol = HeaderRange.Find(what:="207Pb/235U(calc)", MatchCase:=False, lookat:=xlWhole).Column
                        Check207235calc = True
                    Else
                        Check207235calc = False
                    End If
            'Check and define Pb207/U235 ratio
                Set Ratio207235Col = HeaderRange.Find(what:="207Pb/235U", MatchCase:=False, lookat:=xlWhole)
                If Not Ratio207235Col Is Nothing Then
                    Ratio207235Col = HeaderRange.Find(what:="207Pb/235U", MatchCase:=False, lookat:=xlWhole).Column
                    Check207235 = True
                Else
                    Check207235 = False
                End If
            'Check and define Pb206/U238 ratio
                Set Ratio206238Col = HeaderRange.Find(what:="206Pb/238U", MatchCase:=False, lookat:=xlWhole)
                If Not Ratio206238Col Is Nothing Then
                    Ratio206238Col = HeaderRange.Find(what:="206Pb/238U", MatchCase:=False).Column
                    Check206238 = True
                Else
                    Check206238 = False
                End If
            'Check and define Pb207/Pb206 ratio
                Set Ratio207206Col = HeaderRange.Find(what:="207Pb/206Pb", MatchCase:=False, lookat:=xlWhole)
                If Not Ratio207206Col Is Nothing Then
                    Ratio207206Col = HeaderRange.Find(what:="207Pb/206Pb", MatchCase:=False).Column
                    Check207206 = True
                Else
                    Check207206 = False
                End If
            'Check and define Pb207/Pb206 ratio
                Set Ratio204206Col = HeaderRange.Find(what:="204Pb/206Pb", MatchCase:=False, lookat:=xlWhole)
                If Not Ratio204206Col Is Nothing Then
                    Ratio204206Col = HeaderRange.Find(what:="204Pb/206Pb", MatchCase:=False).Column
                    Check204206 = True
                Else
                    Check204206 = False
                End If
            'Check and define Pb208/Th232 ratio
                Set Ratio208232Col = HeaderRange.Find(what:="208Pb/232Th", MatchCase:=False, lookat:=xlWhole)
                If Not Ratio208232Col Is Nothing Then
                    Ratio208232Col = HeaderRange.Find(what:="208Pb/232Th", MatchCase:=False, lookat:=xlWhole).Column
                    Check208232 = True
                Else
                    Check208232 = False
                End If
            'Check and define 238/206 AgeEstimate
                Select Case Check238206
                    Case True
                        AgeEstimate238206Col = HeaderRange.Find(what:="238U/206Pb Age (Ma)", MatchCase:=False, lookat:=xlWhole).Column
                    Case False
                End Select
            'Check and define 207/235 (calc) AgeEstimate
                Select Case Check207235calc
                    Case True
                        AgeEstimate207235calcCol = HeaderRange.Find(what:="207Pb/235U(calc) Age (Ma)", MatchCase:=False).Column
                    Case False
                End Select
            'Check and define 207/235 AgeEstimate
                Select Case Check207235
                    Case True
                        AgeEstimate207235Col = HeaderRange.Find(what:="207Pb/235U Age (Ma)", MatchCase:=False).Column
                    Case False
                End Select
            'Check and define 206/238 AgeEstimate
                Select Case Check206238
                    Case True
                        AgeEstimate206238Col = HeaderRange.Find(what:="206Pb/238U Age (Ma)", MatchCase:=False, lookat:=xlWhole).Column
                    Case False
                End Select
            'Check and define 208/232 AgeEstimate
                Select Case Check208232
                    Case True
                        AgeEstimate208232Col = HeaderRange.Find(what:="208Pb/232Th Age (Ma)", MatchCase:=False, lookat:=xlWhole).Column
                    Case False
                End Select
            'Check and define 207/206 AgeEstimate
                Select Case Check207206
                    Case True
                        AgeEstimate207206Col = HeaderRange.Find(what:="207Pb/206Pb Age (Ma)", MatchCase:=False, lookat:=xlWhole).Column
                    Case False
                End Select
            'Check and define error correlations
                Set Rho206Pb238Uvs207Pb235UCol = HeaderRange.Find(what:="Rho: 206/238 vs 207/235", MatchCase:=False, lookat:=xlWhole)
                    If Not Rho206Pb238Uvs207Pb235UCol Is Nothing Then
                        RhoCalc = False
                        Rho206Pb238Uvs207Pb235UCol = HeaderRange.Find(what:="Rho: 206/238 vs 207/235", MatchCase:=False, lookat:=xlWhole).Column
                    ElseIf SignalPrecision = True Then
                        RhoCalc = True
                    Else
                        RhoCalc = False
                    End If
                Set Rho207Pb206Pbvs238U206PbCol = HeaderRange.Find(what:="Rho: 207/206 vs 238/206", MatchCase:=False, lookat:=xlWhole)
                If Not Rho207Pb206Pbvs238U206PbCol Is Nothing Then
                    Rho207Pb206Pbvs238U206PbCol = HeaderRange.Find(what:="Rho: 207/206 vs 238/206", MatchCase:=False, lookat:=xlWhole).Column
                ElseIf SignalPrecision = True Then
                    RhoCalc = True
                Else
                    RhoCalc = False
                End If
        End With
    'Add new sheets
        Sheets.Add.Name = "Geochronology Data"
        Select Case TraceElementDataPresent
            Case True
                Sheets.Add.Name = "Elemental Data"
            Case False
        End Select
        Sheets("Original Data").Activate
    'Copy AL#, sample name and analysis number
        Range(Cells(ODStartRowA, ALNumCol), Cells(ODLastRowA, ALNumCol)).Copy Destination:=Sheets("Geochronology Data").Range("A1")
        Range(Cells(ODStartRowA, SampleCol), Cells(ODLastRowA, AnalysisCol)).Copy Destination:=Sheets("Geochronology Data").Range("B1")
        Select Case TraceElementDataPresent
            Case True
                Range(Cells(ODStartRowA, ALNumCol), Cells(ODLastRowA, ALNumCol)).Copy Destination:=Sheets("Elemental Data").Range("A1")
                Range(Cells(ODStartRowA, SampleCol), Cells(ODLastRowA, AnalysisCol)).Copy Destination:=Sheets("Elemental Data").Range("B1")
            Case False
        End Select
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
    'Copy geochronology data, suitable for isoplotR input
        'Define GDNextCol, GDLastRow
            GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
            GDNextCol = GDLastCol + 1
            GDLastRow = Sheets("Geochronology Data").Cells(Rows.Count, 1).End(xlUp).Row
    Select Case ConcordiaPlotType 'arrange ratios, ratio uncertainties, age estimates, age estimate uncertainties, copy/calculation of error correlations, calculations of concordance based on prefered output type for concordia plot
        Case "Wetherill" '(Wetherill format)
            'Copy ratios and uncertainties, label uncertainty columns
                '238/206 ratio
                    Select Case Check238206
                        Case True
                            Range(Cells(ODStartRowA, Ratio238206Col), Cells(ODLastRowA, Ratio238206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            RCIndexRatio238206 = GDNextCol
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[38/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio238206Col), Cells(ODLastRowB, Ratio238206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '208/206 ratio
                    Select Case Check208206
                        Case True
                            Range(Cells(ODStartRowA, Ratio208206Col), Cells(ODLastRowA, Ratio208206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[08/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio208206Col), Cells(ODLastRowB, Ratio208206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '204/206 ratio
                    Select Case Check204206
                        Case True
                            Range(Cells(ODStartRowA, Ratio204206Col), Cells(ODLastRowA, Ratio204206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[04/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio204206Col), Cells(ODLastRowB, Ratio204206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '208/232 ratio
                    Select Case Check208232
                        Case True
                            Range(Cells(ODStartRowA, Ratio208232Col), Cells(ODLastRowA, Ratio208232Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[08/32] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio208232Col), Cells(ODLastRowB, Ratio208232Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '207/235 (calc) ratio (one 207/235 variant is required for Wetherill format)
                    Select Case Check207235calc
                        Case True
                            Range(Cells(ODStartRowA, Ratio207235calcCol), Cells(ODLastRowA, Ratio207235calcCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/35](calc) " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio207235calcCol), Cells(ODLastRowB, Ratio207235calcCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            If Check207235 = False Then
                            GoTo MissingRequiredRatio
                            Else
                            End If
                    End Select
                '207/235 ratio (one 207/235 variant is required for Wetherill format)
                    Select Case Check207235
                        Case True
                            Range(Cells(ODStartRowA, Ratio207235Col), Cells(ODLastRowA, Ratio207235Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/35] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio207235Col), Cells(ODLastRowB, Ratio207235Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            If Check207235calc = False Then
                            GoTo MissingRequiredRatio
                            Else
                            End If
                    End Select
                '206/238 ratio (required for Wetherill format)
                    Select Case Check206238
                        Case True
                            Range(Cells(ODStartRowA, Ratio206238Col), Cells(ODLastRowA, Ratio206238Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            Select Case Check238206
                                Case False
                                    RCIndexRatio238206 = GDNextCol
                                Case True
                            End Select
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[06/38] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio206238Col), Cells(ODLastRowB, Ratio206238Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            GoTo MissingRequiredRatio
                    End Select
                '207/206 ratio (required for Wetherill format)
                    Select Case Check207206
                        Case True
                            Range(Cells(ODStartRowA, Ratio207206Col), Cells(ODLastRowA, Ratio207206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            RCIndexRatio207206 = GDNextCol
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio207206Col), Cells(ODLastRowB, Ratio207206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            GoTo MissingRequiredRatio
                    End Select
            'Copy/calculate error correlation (rho), approximations based on the mathematics of Schmitz et al 2007
                'Write headers, shift back to original position for rest of function
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rho[206/238][207/235]"
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rho[207/206][206/238]"
                    GDNextCol = GDNextCol - 1
                'Dynamic determination of copy/calculate error correlation (rho)
                    Select Case RhoCalc
                        Case False
                            If IsNumeric(Rho206Pb238Uvs207Pb235UCol) Then
                                Range(Cells(ODStartRowA + 1, Rho206Pb238Uvs207Pb235UCol), Cells(ODLastRowA, Rho206Pb238Uvs207Pb235UCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                                GDNextCol = GDNextCol + 1
                                If IsNumeric(Rho207Pb206Pbvs238U206PbCol) Then
                                    Range(Cells(ODStartRowA + 1, Rho207Pb206Pbvs238U206PbCol), Cells(ODLastRowA, Rho207Pb206Pbvs238U206PbCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                                    GDNextCol = GDNextCol + 1
                                Else
                                    GDNextCol = GDNextCol + 1
                                End If
                            ElseIf IsNumeric(Rho207Pb206Pbvs238U206PbCol) Then
                                GDNextCol = GDNextCol + 1
                                Range(Cells(ODStartRowA + 1, Rho207Pb206Pbvs238U206PbCol), Cells(ODLastRowA, Rho207Pb206Pbvs238U206PbCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                                GDNextCol = GDNextCol + 1
                            Else
                                GDNextCol = GDNextCol + 2
                            End If
                        Case True
                            If IsNumeric(Rho206Pb238Uvs207Pb235UCol) Then
                                Range(Cells(ODStartRowA + 1, Rho206Pb238Uvs207Pb235UCol), Cells(ODLastRowA, Rho206Pb238Uvs207Pb235UCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                                GDNextCol = GDNextCol + 1
                                If IsNumeric(Rho207Pb206Pbvs238U206PbCol) Then
                                    Range(Cells(ODStartRowA + 1, Rho207Pb206Pbvs238U206PbCol), Cells(ODLastRowA, Rho207Pb206Pbvs238U206PbCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                                    GDNextCol = GDNextCol + 1
                                Else
                                    Sheets("Geochronology Data").Activate
                                    '[06/38][07/06]
                                    Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=(((RC[-4]/RC[-5])^2)+((RC[-2]/RC[-3])^2)-((RC[-6]/RC[-7])^2))/(2*(RC[-4]/RC[-5])*(RC[-2]/RC[-3]))"
                                    GDNextCol = GDNextCol + 1
                                    Sheets("Original Data").Activate
                                End If
                            Else
                                '[07/35][06/38]
                                Sheets("Geochronology Data").Activate
                                Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=(((RC[-5]/RC[-6])^2)+((RC[-3]/RC[-4])^2)-((RC[-1]/RC[-2])^2))/(2*(RC[-5]/RC[-6])*(RC[-3]/RC[-4]))"
                                GDNextCol = GDNextCol + 1
                                Sheets("Original Data").Activate
                                If IsNumeric(Rho207Pb206Pbvs238U206PbCol) Then
                                    Range(Cells(ODStartRowA + 1, Rho207Pb206Pbvs238U206PbCol), Cells(ODLastRowA, Rho207Pb206Pbvs238U206PbCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                                    GDNextCol = GDNextCol + 1
                                Else
                                    '[06/38][07/06]
                                    Sheets("Geochronology Data").Activate
                                    Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=(((RC[-4]/RC[-5])^2)+((RC[-2]/RC[-3])^2)-((RC[-6]/RC[-7])^2))/(2*(RC[-4]/RC[-5])*(RC[-2]/RC[-3]))"
                                    GDNextCol = GDNextCol + 1
                                    Sheets("Original Data").Activate
                                End If
                            End If
                    End Select
            'Copy age estimates and uncertainties, label uncertainty columns
                '238/206 age estimate
                    Select Case Check238206
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate238206Col), Cells(ODLastRowA, AgeEstimate238206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[38/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate238206Col), Cells(ODLastRowB, AgeEstimate238206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '208/232 age estimate
                    Select Case Check208232
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate208232Col), Cells(ODLastRowA, AgeEstimate208232Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[08/32] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate208232Col), Cells(ODLastRowB, AgeEstimate208232Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '207/235 (calc) age estimate (one 207/235 variant is required for Wetherill Format)
                    Select Case Check207235calc
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate207235calcCol), Cells(ODLastRowA, AgeEstimate207235calcCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/35](calc) " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate207235calcCol), Cells(ODLastRowB, AgeEstimate207235calcCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            If Check207235 = False Then
                            GoTo MissingRequiredAgeEstimate
                            Else
                            End If
                    End Select
                '207/235 age estimate (one 207/235 variant is required for Wetherill Format)
                    Select Case Check207235
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate207235Col), Cells(ODLastRowA, AgeEstimate207235Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/35] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate207235Col), Cells(ODLastRowB, AgeEstimate207235Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            If Check207235calc = False Then
                            GoTo MissingRequiredAgeEstimate
                            Else
                            End If
                    End Select
                '206/238 age estimate (required for Wetherill format)
                    Select Case Check206238
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate206238Col), Cells(ODLastRowA, AgeEstimate206238Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            RCIndexAge206238 = GDNextCol
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[06/38] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate206238Col), Cells(ODLastRowB, AgeEstimate206238Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            GoTo MissingRequiredAgeEstimate
                    End Select
                '207/206 age estimate required for Wetherill format)
                    Select Case Check207206
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate207206Col), Cells(ODLastRowA, AgeEstimate207206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            RCIndexAge207206 = GDNextCol
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate207206Col), Cells(ODLastRowB, AgeEstimate207206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            GoTo MissingRequiredAgeEstimate
                    End Select
            'Concordance calculations
                Sheets("Geochronology Data").Activate
                    With ActiveSheet
                        '[06/38][07/06]
                            Cells(1, GDNextCol).Value = "Concordance [06/38][07/06]"
                            Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=ROUND((RC[-4]/RC[-2])*100,0)"
                            GDNextCol = GDNextCol + 1
                        '[06/38][07/35]
                            Cells(1, GDNextCol).Value = "Concordance [06/38][07/35]"
                            Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=ROUND((RC[-5]/RC[-7])*100,0)"
                            GDNextCol = GDNextCol + 1
                        '[07/35][07/06]
                            Cells(1, GDNextCol).Value = "Concordance [07/35][07/06]"
                            Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=ROUND((RC[-8]/RC[-4])*100,0)"
                            GDNextCol = GDNextCol + 1
                        'Logratio Distance (Aitchison Distance)
                            Cells(1, GDNextCol).Value = "Concordance [logratio distance]"
                            RCIndexRatio207206 = RCIndexRatio207206 - GDNextCol
                            RCIndexRatio238206 = RCIndexRatio238206 - GDNextCol
                            RCIndexAge206238 = RCIndexAge206238 - GDNextCol
                            RCIndexAge207206 = RCIndexAge207206 - GDNextCol
                            Select Case Check238206
                                Case True
                                    Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=100*(LN(RC[" & RCIndexRatio238206 & "])-LN(EXP(0.000155125*RC[" & RCIndexAge207206 & "])-1))*SIN(ATAN((LN(RC[" & RCIndexRatio207206 & "])-LN((1/137.818)*(EXP(0.00098485*RC[" & RCIndexAge206238 & "])-1)/(EXP(0.000155125*RC[" & RCIndexAge206238 & "])-1)))/(LN(RC[" & RCIndexRatio238206 & "])-LN(EXP(0.000155125*RC[" & RCIndexAge207206 & "])-1))))"
                                    GDNextCol = GDNextCol + 1
                                Case False
                                    Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=100*(LN(1/RC[" & RCIndexRatio238206 & "])-LN(EXP(0.000155125*RC[" & RCIndexAge207206 & "])-1))*SIN(ATAN((LN(RC[" & RCIndexRatio207206 & "])-LN((1/137.818)*(EXP(0.00098485*RC[" & RCIndexAge206238 & "])-1)/(EXP(0.000155125*RC[" & RCIndexAge206238 & "])-1)))/(LN(1/RC[" & RCIndexRatio238206 & "])-LN(EXP(0.000155125*RC[" & RCIndexAge207206 & "])-1))))"
                                    GDNextCol = GDNextCol + 1
                            End Select
                    End With
        Case "TeraWasserburg" '(TeraWasserburg Format)
            'Copy ratios and uncertainties, label uncertainty columns
                '208/206 ratio
                    Select Case Check208206
                        Case True
                            Range(Cells(ODStartRowA, Ratio208206Col), Cells(ODLastRowA, Ratio208206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[08/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio208206Col), Cells(ODLastRowB, Ratio208206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '204/206 ratio
                    Select Case Check204206
                        Case True
                            Range(Cells(ODStartRowA, Ratio204206Col), Cells(ODLastRowA, Ratio204206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[04/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio204206Col), Cells(ODLastRowB, Ratio204206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '208/232 ratio
                    Select Case Check208232
                        Case True
                            Range(Cells(ODStartRowA, Ratio208232Col), Cells(ODLastRowA, Ratio208232Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[08/32] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio208232Col), Cells(ODLastRowB, Ratio208232Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '207/235 (calc) ratio
                    Select Case Check207235calc
                        Case True
                            Range(Cells(ODStartRowA, Ratio207235calcCol), Cells(ODLastRowA, Ratio207235calcCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/35](calc) " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio207235calcCol), Cells(ODLastRowB, Ratio207235calcCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '207/235 ratio
                    Select Case Check207235
                        Case True
                            Range(Cells(ODStartRowA, Ratio207235Col), Cells(ODLastRowA, Ratio207235Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/35] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio207235Col), Cells(ODLastRowB, Ratio207235Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '206/238 ratio
                    Select Case Check206238
                        Case True
                            Range(Cells(ODStartRowA, Ratio206238Col), Cells(ODLastRowA, Ratio206238Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[06/38] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio206238Col), Cells(ODLastRowB, Ratio206238Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '238/206 ratio (required for Tera-Wasserburg format)
                    Select Case Check238206
                        Case True
                            Range(Cells(ODStartRowA, Ratio238206Col), Cells(ODLastRowA, Ratio238206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            RCIndexRatio238206 = GDNextCol
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[38/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio238206Col), Cells(ODLastRowB, Ratio238206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            GoTo MissingRequiredRatio
                    End Select
                '207/206 ratio (required for Tera-Wasserburg format)
                    Select Case Check207206
                        Case True
                            Range(Cells(ODStartRowA, Ratio207206Col), Cells(ODLastRowA, Ratio207206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            RCIndexRatio207206 = GDNextCol
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, Ratio207206Col), Cells(ODLastRowB, Ratio207206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            GoTo MissingRequiredRatio
                    End Select
            'Copy/calculate error correlation (rho), write header, approximations based on the mathematics of Schmitz et al 2007
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rho[38/06][07/06]"
                    Select Case RhoCalc
                        Case False
                            If IsNumeric(Rho207Pb206Pbvs238U206PbCol) Then
                                Range(Cells(ODStartRowA + 1, Rho207Pb206Pbvs238U206PbCol), Cells(ODLastRowA, Rho207Pb206Pbvs238U206PbCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                                GDNextCol = GDNextCol + 1
                            Else
                                GDNextCol = GDNextCol + 1
                            End If
                        Case True
                            If IsNumeric(Rho207Pb206Pbvs238U206PbCol) Then
                                Range(Cells(ODStartRowA + 1, Rho207Pb206Pbvs238U206PbCol), Cells(ODLastRowA, Rho207Pb206Pbvs238U206PbCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                                GDNextCol = GDNextCol + 1
                            Else
                                Select Case Check207235
                                    Case True
                                        Sheets("Geochronology Data").Activate
                                        '[38/06][07/06] (requires 207/235)
                                        Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=(((RC[-3]/RC[-4])^2)+((RC[-1]/RC[-2])^2)-((RC[-7]/RC[-8])^2))/(2*(RC[-3]/RC[-4])*(RC[-1]/RC[-2]))"
                                        GDNextCol = GDNextCol + 1
                                        Sheets("Original Data").Activate
                                    Case False
                                        GDNextCol = GDNextCol + 1
                                End Select
                            End If
                    End Select
            'Copy age uncertainty estimates and uncertainties, label uncertainty columns
                '208/232 age estimate
                    Select Case Check208232
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate208232Col), Cells(ODLastRowA, AgeEstimate208232Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[08/32] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate208232Col), Cells(ODLastRowB, AgeEstimate208232Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '207/235 (calc) age estimate
                    Select Case Check207235calc
                    Case True
                        Range(Cells(ODStartRowA, AgeEstimate207235calcCol), Cells(ODLastRowA, AgeEstimate207235calcCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                        GDNextCol = GDNextCol + 1
                        Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/35](calc) " & StandardErrorLevel & "SE"
                        Range(Cells(ODStartRowB, AgeEstimate207235calcCol), Cells(ODLastRowB, AgeEstimate207235calcCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                        GDNextCol = GDNextCol + 1
                    Case False
                    End Select
                '207/235 age estimate
                    Select Case Check207235
                    Case True
                        Range(Cells(ODStartRowA, AgeEstimate207235Col), Cells(ODLastRowA, AgeEstimate207235Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                        GDNextCol = GDNextCol + 1
                        Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/35] " & StandardErrorLevel & "SE"
                        Range(Cells(ODStartRowB, AgeEstimate207235Col), Cells(ODLastRowB, AgeEstimate207235Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                        GDNextCol = GDNextCol + 1
                    Case False
                    End Select
                '206/238 age estimate
                    Select Case Check206238
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate206238Col), Cells(ODLastRowA, AgeEstimate206238Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            RCIndexAge206238 = GDNextCol
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[06/38] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate206238Col), Cells(ODLastRowB, AgeEstimate206238Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                    End Select
                '238/206 age estimate
                    Select Case Check238206
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate238206Col), Cells(ODLastRowA, AgeEstimate238206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            Select Case Check206238
                                Case False
                                    RCIndexAge206238 = GDNextCol
                                Case True
                            End Select
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[38/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate238206Col), Cells(ODLastRowB, AgeEstimate238206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            GoTo MissingRequiredAgeEstimate
                    End Select
                '207/206 age estimate
                    Select Case Check207206
                        Case True
                            Range(Cells(ODStartRowA, AgeEstimate207206Col), Cells(ODLastRowA, AgeEstimate207206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                            RCIndexAge207206 = GDNextCol
                            GDNextCol = GDNextCol + 1
                            Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[07/06] " & StandardErrorLevel & "SE"
                            Range(Cells(ODStartRowB, AgeEstimate207206Col), Cells(ODLastRowB, AgeEstimate207206Col)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Case False
                            GoTo MissingRequiredAgeEstimate
                    End Select
            'Concordance calculations
                Sheets("Geochronology Data").Activate
                    With ActiveSheet
                    '[38/06][07/06]
                        Cells(1, GDNextCol).Value = "Concordance [38/06][07/06]"
                        Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=ROUND((RC[-4]/RC[-2])*100,0)"
                        GDNextCol = GDNextCol + 1
                    'Logratio Distance (Aitchison Distance)
                        Cells(1, GDNextCol).Value = "Concordance [logratio distance]"
                        RCIndexRatio207206 = RCIndexRatio207206 - GDNextCol
                        RCIndexRatio238206 = RCIndexRatio238206 - GDNextCol
                        RCIndexAge206238 = RCIndexAge206238 - GDNextCol
                        RCIndexAge207206 = RCIndexAge207206 - GDNextCol
                        Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=100*(LN(RC[" & RCIndexRatio238206 & "])-LN(EXP(0.000155125*RC[" & RCIndexAge207206 & "])-1))*SIN(ATAN((LN(RC[" & RCIndexRatio207206 & "])-LN((1/137.818)*(EXP(0.00098485*RC[" & RCIndexAge206238 & "])-1)/(EXP(0.000155125*RC[" & RCIndexAge206238 & "])-1)))/(LN(RC[" & RCIndexRatio238206 & "])-LN(EXP(0.000155125*RC[" & RCIndexAge207206 & "])-1))))"
                        GDNextCol = GDNextCol + 1
                    End With
    End Select
    Sheets("Original Data").Activate 'copying comments, source filename, elemental uncertainties
        'Copy Comments
            Range(Cells(ODStartRowA, CommentsCol), Cells(ODLastRowA, CommentsCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
            GDNextCol = GDNextCol + 1
            Select Case TraceElementDataPresent
                Case True
                    Range(Cells(ODStartRowA, CommentsCol), Cells(ODLastRowA, CommentsCol)).Copy Destination:=Sheets("Elemental Data").Cells(1, EDNextCol)
                    EDNextCol = EDNextCol + 1
                Case False
            End Select
        'Copy Source Filename
            Range(Cells(ODStartRowA, SourceFileCol), Cells(ODLastRowA, SourceFileCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
            SFColDel = GDNextCol
    Sheets("Geochronology Data").Activate 'Correct sample label and trailing number for correct sorting in Excel
            GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
            GDLastRow = Sheets("Geochronology Data").Cells(Rows.Count, 1).End(xlUp).Row
            For n = 2 To GDLastRow
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
                    Range("B1", "C" & GDLastRow).Copy Destination:=Sheets("Elemental Data").Range("B1")
                Case False
            End Select
    'Sort prior to cut - performance optimisation
        'Delete sourcefile column
            Columns(SFColDel).Delete
        'Set GD Last Column
            GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
        'Sort geochronology data
            Set GDRange = Sheets("Geochronology Data").Range("A1", Cells(GDLastRow, GDLastCol))
            With GDRange
                .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
            End With
        'Set ED Last Column and sort elemental data
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
            Sheets.Add.Name = "Geochronology Data - Standards"
            Select Case TraceElementDataPresent
                Case True
                    Sheets.Add.Name = "Elemental Data - Standards"
                    Sheets("Elemental Data - Standards").Move after:=Sheets("Elemental Data")
                Case False
            End Select
        'Define GD and ED variables
            GDLastRow = Sheets("Geochronology Data").Range("A" & Rows.Count).End(xlUp).Row
            GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
            Select Case TraceElementDataPresent
                Case True
                    EDLastRow = Sheets("Elemental Data").Range("A" & Rows.Count).End(xlUp).Row
                    EDLastCol = Sheets("Elemental Data").Cells(1, Columns.Count).End(xlToLeft).Column
                Case False
            End Select
        Sheets("Geochronology Data").Activate 'Cut standards from geochronology data
            Range("A1", Cells(1, GDLastCol)).Copy Destination:=Sheets("Geochronology Data - Standards").Range("A1") 'copy headers
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
                Range("A" & Standard1f, Cells(Standard1l, GDLastCol)).Cut Destination:=Sheets("Geochronology Data - Standards").Range("A" & Standard1f)
            'Cut Standard2
                On Error Resume Next
                Range("A" & Standard3f, Cells(Standard3l, GDLastCol)).Cut Destination:=Sheets("Geochronology Data - Standards").Range("A" & Standard3f)
            'Cut  Standard3
                On Error Resume Next
                Range("A" & Standard2f, Cells(Standard2l, GDLastCol)).Cut Destination:=Sheets("Geochronology Data - Standards").Range("A" & Standard2f)
            'Cut  Standard4
                On Error Resume Next
                Range("A" & Standard4f, Cells(Standard4l, GDLastCol)).Cut Destination:=Sheets("Geochronology Data - Standards").Range("A" & Standard4f)
            'Cut  Standard5
                On Error Resume Next
                Range("A" & Standard5f, Cells(Standard5l, GDLastCol)).Cut Destination:=Sheets("Geochronology Data - Standards").Range("A" & Standard5f)
            'Sort after cut
                'Redefine geochronology last rows and columnss
                    GDLastRow = Sheets("Geochronology Data").Range("A" & Rows.Count).End(xlUp).Row
                    GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
                    GDSLastRow = Sheets("Geochronology Data - Standards").Range("A" & Rows.Count).End(xlUp).Row
                    GDSLastCol = Sheets("Geochronology Data - Standards").Cells(1, Columns.Count).End(xlToLeft).Column
                'Geochronology Data
                    Set GDRange = Sheets("Geochronology Data").Range("A1", Cells(GDLastRow, GDLastCol))
                    With GDRange
                        .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
                    End With
                'Geochronology Data - Standards
                    Sheets("Geochronology Data - Standards").Activate
                    Set GDSRange = Sheets("Geochronology Data - Standards").Range("A1", Cells(GDSLastRow, GDSLastCol))
                    With GDSRange
                        .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
                    End With
            'Rename and move unknowns sheet
                Sheets("Geochronology Data").Name = "Geochronology Data - Unknowns"
                Sheets("Geochronology Data - Unknowns").Move before:=Sheets("Geochronology Data - Standards")
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
            'Geochronology  Data Standards
                Sheets("Geochronology Data - Standards").Activate
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
                    AgeEstimateFirstCol = Range("A1", Cells(1, LastCol)).Find(what:="Age (Ma)", MatchCase:=True, lookat:=xlPart).Column
                    AgeEstimateLastCol = Range("A1", Cells(1, LastCol)).Find(what:="Concordance", MatchCase:=True, lookat:=xlPart).Column - 1
                    .Range(Cells(2, AgeEstimateFirstCol), Cells(LastRow, AgeEstimateLastCol)).NumberFormat = "0.0"
                End With
                With ActiveWindow
                    .SplitColumn = 3
                    .SplitRow = 1
                    .FreezePanes = True
                End With
            'Geochronology  Data Unknowns
                Sheets("Geochronology Data - Unknowns").Activate
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
                    AgeEstimateFirstCol = Range("A1", Cells(1, LastCol)).Find(what:="Age (Ma)", MatchCase:=True, lookat:=xlPart).Column
                    AgeEstimateLastCol = Range("A1", Cells(1, LastCol)).Find(what:="Concordance", MatchCase:=True, lookat:=xlPart).Column - 1
                    .Range(Cells(2, AgeEstimateFirstCol), Cells(LastRow, AgeEstimateLastCol)).NumberFormat = "0.0"
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
MissingRequiredRatio:     MsgBox "The LADR output is missing required ratios for this arrangment of data" & vbCrLf & "Wetherill requires 207/235, 206/238, and 207/206" & vbCrLf & "Tera-Wasserburg requires 207/206, 238/206", , "LADR U-Pb Arranger"
        Exit Sub
MissingRequiredAgeEstimate:     MsgBox "The LADR output is missing required ratios for this arrangment of data" & vbCrLf & "Wetherill requires 207/235, 206/238, and 207/206" & vbCrLf & "Tera-Wasserburg requires 207/206, 238/206", , "LADR U-Pb Arranger"
        Exit Sub
SaveError:     MsgBox "There was an error saving this file during execution of the add-in." & vbCrLf & "This is likely due to a file of the same name already existing" & vbCrLf & vbCrLf & "Please remember to save your file if you haven't already", , "LADR U-Pb Arranger"
End Sub
