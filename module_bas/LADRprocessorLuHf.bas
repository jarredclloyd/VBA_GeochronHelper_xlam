Attribute VB_Name = "LADRprocessorLuHf"
'Module for handling LADR Lu-Hf geochronology data outputs
'Created By Jarred Lloyd on 2022-06-14
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
    'Error correlation calculation
        Dim SignalPrecision As Boolean
        Dim RhoCalc As Boolean
        Dim RhoLuHfCol As Variant
    'Reported uncertainy level
        Dim UncertaintyLevel As Variant
        Dim StandardErrorLevel As Variant
    'Varibles for copying of original data
        Dim ODLastRowA As Long
        Dim ODStartRowA As Long
        Dim ODLastRowB As Long
        Dim ODStartRowB As Long
        Dim ODLastRowC As Long
        Dim ODStartRowC As Long
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
        Dim ElementTotalCol As Long
        Dim EleStartCol As Variant
        Dim EleEndCol As Long
        Dim TraceElementDataPresent As Boolean
        Dim Ratio176Lu177HfCol As Variant
        Dim Ratio176Hf177HfCol As Variant
        Dim Ratio176Lu176HfCol As Variant
        Dim Ratio177Hf176HfCol As Variant
        Dim CPScol1 As Variant
        Dim CPScol2 As Variant
        Dim CPScol3 As Variant
        Dim CommonIsotope As String
        Dim HeaderRange As Range
        Dim HeaderRangeEle As Range
        Dim CommentsColEle As Long
        Dim EleUnStartCol As Long
        Dim EleUnEndCol As Long
    'Variables for ratio pair checking
        Dim ElementSymNumOrder As Variant
        Dim CheckRatio176Lu177Hf As Boolean
        Dim CheckRatio176Hf177Hf As Boolean
        Dim CheckRatio176Lu176Hf As Boolean
        Dim CheckRatio177Hf176Hf As Boolean
        Dim RatioPairNormalPresent As Boolean
        Dim RatioPairInversePresent As Boolean
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
    'Variables for age/ratio rounding
        Dim cell As Range
        Dim GDAgeI As Long
        Dim GDAgeF As Long
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
Sub LADRLuHfArranger()

    'Define number and name of standards
DefineStandards:            NumStandards = Application.InputBox("How many different standards were used?", "LADR_LuHf_Arranger", 4, Type:=1)
        Select Case NumStandards
            Case 1
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case 2
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case 3
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case 4
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case 5
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard5 = Application.InputBox("What is the sample name of the fifth standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case Else
                Ans = MsgBox("Please enter the number of standards, this has to be a value of 1 to 5." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_LuHf_Arranger")
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
StdCheckError:      Ans = MsgBox("Standards variables are not set correctly." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_LuHf_Arranger")
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
    'Find GSubCPS and set variables
        With Sheets("Original Data")
            ODStartRowC = Range("A:A").Find(what:="GBSub_CPS", MatchCase:=True, lookat:=xlPart).Row
            ODStartRowC = ODStartRowC + 3
            ODLastRowC = Range("C" & ODStartRowC).End(xlDown).Row
        End With
    'Define element symbol and number order
            If Cells(FirstMassRow, 1).Value Like "#*" Then
                ElementSymNumOrder = "NumSym"
            ElseIf Cells(FirstMassRow, 1).Value Like "[A-Z]*" Then
                ElementSymNumOrder = "SymNum"
            End If
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
                    TraceElementDataPresent = True
                Else
                    TraceElementDataPresent = False
                End If
        
        Select Case ElementSymNumOrder
            Case "NumSym"
                'CPS Columns
                    Select Case TraceElementDataPresent
                        Case True
                            CPScol1 = HeaderRange.Find(what:="175Lu", MatchCase:=False, lookat:=xlPart).Column
                            CPScol2 = HeaderRange.Find(what:="176Hf", MatchCase:=False, lookat:=xlPart).Column
                            CPScol3 = HeaderRange.Find(what:="178Hf", MatchCase:=False, lookat:=xlPart).Column
                        Case False
                    End Select
                'Check and define Lu176/Hf177 ratio
                    Set Ratio176Lu177HfCol = HeaderRange.Find(what:="175Lu/178Hf->260", MatchCase:=False)
                    If Not Ratio176Lu177HfCol Is Nothing Then
                        Ratio176Lu177HfCol = HeaderRange.Find(what:="175Lu/178Hf->260", MatchCase:=False).Column
                        CheckRatio176Lu177Hf = True
                    Else
                        Set Ratio176Lu177HfCol = HeaderRange.Find(what:="175Lu/178Hf", MatchCase:=False)
                        If Not Ratio176Lu177HfCol Is Nothing Then
                            Ratio176Lu177HfCol = HeaderRange.Find(what:="175Lu/178Hf", MatchCase:=False).Column
                            CheckRatio176Lu177Hf = True
                        Else
                            CheckRatio176Lu177Hf = False
                        End If
                    End If
                'Check and define Hf176/Hf177 ratio
                    Set Ratio176Hf177HfCol = HeaderRange.Find(what:="176Hf->258/178Hf->260", MatchCase:=False)
                    If Not Ratio176Hf177HfCol Is Nothing Then
                        Ratio176Hf177HfCol = HeaderRange.Find(what:="176Hf->258/178Hf->260", MatchCase:=False).Column
                        CheckRatio176Hf177Hf = True
                    Else
                        Set Ratio176Hf177HfCol = HeaderRange.Find(what:="176Hf/178Hf", MatchCase:=False)
                        If Not Ratio176Hf177HfCol Is Nothing Then
                            Ratio176Hf177HfCol = HeaderRange.Find(what:="176Hf/178Hf", MatchCase:=False).Column
                            CheckRatio176Hf177Hf = True
                        Else
                            CheckRatio176Hf177Hf = False
                        End If
                    End If
                'Check and define Lu176/Hf176 Ratio
                    Set Ratio176Lu176HfCol = HeaderRange.Find(what:="175Lu/176Hf->258", MatchCase:=False)
                    If Not Ratio176Lu176HfCol Is Nothing Then
                        Ratio176Lu176HfCol = HeaderRange.Find(what:="175Lu/176Hf->258", MatchCase:=False).Column
                        CheckRatio176Lu176Hf = True
                    Else
                        Set Ratio176Lu176HfCol = HeaderRange.Find(what:="175Lu/176Hf", MatchCase:=False)
                        If Not Ratio176Lu176HfCol Is Nothing Then
                            Ratio176Lu176HfCol = HeaderRange.Find(what:="175Lu/176Hf", MatchCase:=False).Column
                            CheckRatio176Lu176Hf = True
                        Else
                            CheckRatio176Lu176Hf = False
                        End If
                    End If
                'Check and define Hf177/Hf176 ratio
                    Set Ratio177Hf176HfCol = HeaderRange.Find(what:="178Hf->260/176Hf->258", MatchCase:=False)
                    If Not Ratio177Hf176HfCol Is Nothing Then
                        Ratio177Hf176HfCol = HeaderRange.Find(what:="178Hf->260/176Hf->258", MatchCase:=False).Column
                        CheckRatio177Hf176Hf = True
                    Else
                        Set Ratio177Hf176HfCol = HeaderRange.Find(what:="178Hf/176Hf", MatchCase:=False)
                        If Not Ratio177Hf176HfCol Is Nothing Then
                            Ratio177Hf176HfCol = HeaderRange.Find(what:="178Hf/176Hf", MatchCase:=False).Column
                            CheckRatio177Hf176Hf = True
                        Else
                            CheckRatio177Hf176Hf = False
                        End If
                    End If
                Case "SymNum"
                'CPS Columns
                    Select Case TraceElementDataPresent
                        Case True
                            CPScol1 = HeaderRange.Find(what:="Lu175", MatchCase:=False, lookat:=xlPart).Column
                            CPScol2 = HeaderRange.Find(what:="Hf176", MatchCase:=False, lookat:=xlPart).Column
                            CPScol3 = HeaderRange.Find(what:="Hf178", MatchCase:=False, lookat:=xlPart).Column
                        Case False
                    End Select
                'Check and define Lu176/Hf177 ratio
                    Set Ratio176Lu177HfCol = HeaderRange.Find(what:="Lu175/Hf178->260", MatchCase:=False)
                    If Not Ratio176Lu177HfCol Is Nothing Then
                        Ratio176Lu177HfCol = HeaderRange.Find(what:="Lu175/Hf178->260", MatchCase:=False).Column
                        CheckRatio176Lu177Hf = True
                    Else
                        Set Ratio176Lu177HfCol = HeaderRange.Find(what:="Lu175/Hf178", MatchCase:=False)
                        If Not Ratio176Lu177HfCol Is Nothing Then
                            Ratio176Lu177HfCol = HeaderRange.Find(what:="Lu175/Hf178", MatchCase:=False).Column
                            CheckRatio176Lu177Hf = True
                        Else
                            CheckRatio176Lu177Hf = False
                        End If
                    End If
                'Check and define Hf176/Hf177 ratio
                    Set Ratio176Hf177HfCol = HeaderRange.Find(what:="Hf176->258/Hf178->260", MatchCase:=False)
                    If Not Ratio176Hf177HfCol Is Nothing Then
                        Ratio176Hf177HfCol = HeaderRange.Find(what:="Hf176->258/Hf178->260", MatchCase:=False).Column
                        CheckRatio176Hf177Hf = True
                    Else
                        Set Ratio176Hf177HfCol = HeaderRange.Find(what:="Hf176/Hf178", MatchCase:=False)
                        If Not Ratio176Hf177HfCol Is Nothing Then
                            Ratio176Hf177HfCol = HeaderRange.Find(what:="Hf176/Hf178", MatchCase:=False).Column
                            CheckRatio176Hf177Hf = True
                        Else
                            CheckRatio176Hf177Hf = False
                        End If
                    End If
                'Check and define Lu176/Hf176 Ratio
                    Set Ratio176Lu176HfCol = HeaderRange.Find(what:="Lu175/Hf176->258", MatchCase:=False)
                    If Not Ratio176Lu176HfCol Is Nothing Then
                        Ratio176Lu176HfCol = HeaderRange.Find(what:="Lu175/Hf176->258", MatchCase:=False).Column
                        CheckRatio176Lu176Hf = True
                    Else
                        Set Ratio176Lu176HfCol = HeaderRange.Find(what:="Lu175/Hf176", MatchCase:=False)
                        If Not Ratio176Lu176HfCol Is Nothing Then
                            Ratio176Lu176HfCol = HeaderRange.Find(what:="Lu175/Hf176", MatchCase:=False).Column
                            CheckRatio176Lu176Hf = True
                        Else
                            CheckRatio176Lu176Hf = False
                        End If
                    End If
                'Check and define Hf177/Hf176 ratio
                    Set Ratio177Hf176HfCol = HeaderRange.Find(what:="Hf178->260/Hf176->258", MatchCase:=False)
                    If Not Ratio177Hf176HfCol Is Nothing Then
                        Ratio177Hf176HfCol = HeaderRange.Find(what:="Hf178->260/Hf176->258", MatchCase:=False).Column
                        CheckRatio177Hf176Hf = True
                    Else
                        Set Ratio177Hf176HfCol = HeaderRange.Find(what:="Hf178/Hf176", MatchCase:=False)
                        If Not Ratio177Hf176HfCol Is Nothing Then
                            Ratio177Hf176HfCol = HeaderRange.Find(what:="Hf178/Hf176", MatchCase:=False).Column
                            CheckRatio177Hf176Hf = True
                        Else
                            CheckRatio177Hf176Hf = False
                        End If
                    End If
                End Select
            'Check that at least one ratio pair (normal or inverse isochron) is present
                If CheckRatio176Lu177Hf = True And CheckRatio176Hf177Hf = True Then
                    RatioPairNormalPresent = True
                Else
                    RatioPairNormalPresent = False
                End If
                If CheckRatio176Lu176Hf = True And CheckRatio177Hf176Hf = True Then
                    RatioPairInversePresent = True
                Else
                    RatioPairInversePresent = False
                End If
                If RatioPairNormalPresent = False And RatioPairInversePresent = False Then
                    MsgBox ("At least one pair of ratios is required for this procedure to continue." & vbCrLf & vbCrLf & "Please check that either Lu176/Hf177 AND Hf176/Hf177 OR Lu176/Hf177 AND Hf177/Hf176 are present in the data. These can be in alias forms like 85Rb/87Sr->103." & vbCrLf & vbCrLf & "Procedure ended.")
                    Exit Sub
                Else
                End If
            'Check and define error correlations (U/Pb systems used to force LADR to calculate error correlation, otherwise two system approach is used but is likely to be unstable)
                Set RhoLuHfCol = HeaderRange.Find(what:="Rho: 176Lu/177Hf vs 176Hf/177Hf", MatchCase:=False, lookat:=xlPart)
                If Not RhoLuHfCol Is Nothing Then
                    RhoCalc = False
                    RhoLuHfCol = HeaderRange.Find(what:="Rho: 176Lu/176Hf vs 177Hf/176Hf", MatchCase:=False, lookat:=xlPart).Column
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
        Sheets("Geochronology Data").Range("A1").Value = "ALnum"
        Range(Cells(ODStartRowA, SampleCol), Cells(ODLastRowA, AnalysisCol)).Copy Destination:=Sheets("Geochronology Data").Range("B1")
        Select Case TraceElementDataPresent
            Case True
                Range(Cells(ODStartRowA, ALNumCol), Cells(ODLastRowA, ALNumCol)).Copy Destination:=Sheets("Elemental Data").Range("A1")
                Sheets("Elemental Data").Range("A1").Value = "ALnum"
                Range(Cells(ODStartRowA, SampleCol), Cells(ODLastRowA, AnalysisCol)).Copy Destination:=Sheets("Elemental Data").Range("B1")
            Case False
        End Select
    'Copy elemental concentrations and uncertainty data
        Select Case TraceElementDataPresent
            Case True
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
            Case False
        End Select
    'Copy geochronology data, suitable for isoplotR input
        'Define GDNextCol, GDLastRow
            GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
            GDNextCol = GDLastCol + 1
            GDLastRow = Sheets("Geochronology Data").Cells(Rows.Count, 1).End(xlUp).Row
        'Copy Important CPS columns
            Select Case TraceElementDataPresent
            Case True
                'Lu175
                    Range(Cells(ODStartRowC, CPScol1), Cells(ODLastRowC, CPScol1)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Lu175_CPS"
                    GDNextCol = GDNextCol + 1
                'Hf176
                    Range(Cells(ODStartRowC, CPScol2), Cells(ODLastRowC, CPScol2)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Hf176_CPS"
                    GDNextCol = GDNextCol + 1
                'Hf178
                    Range(Cells(ODStartRowC, CPScol3), Cells(ODLastRowC, CPScol3)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Hf178_CPS"
                    GDNextCol = GDNextCol + 1
                Case False
            End Select
        'Copy ratios and uncertainties, label uncertainty columns
        'Normal isochron ratios
            Select Case RatioPairNormalPresent
            Case True
                'Lu176/Hf177
                    Range(Cells(ODStartRowA, Ratio176Lu177HfCol), Cells(ODLastRowA, Ratio176Lu177HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Lu176Hf177"
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Lu176Hf177_" & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio176Lu177HfCol), Cells(ODLastRowB, Ratio176Lu177HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
                'Hf176/Hf177
                    Range(Cells(ODStartRowA, Ratio176Hf177HfCol), Cells(ODLastRowA, Ratio176Hf177HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Hf176Hf177"
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Hf176Hf177_" & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio176Hf177HfCol), Cells(ODLastRowB, Ratio176Hf177HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
            'Copy/calculate error correlation (rho)
                Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rho_Lu176Hf177_Hf176Hf177"
                Select Case RhoCalc
                    Case False
                        If IsNumeric(RhoLuHfCol) Then
                            Range(Cells(ODStartRowA + 1, RhoLuHfCol), Cells(ODLastRowA, RhoLuHfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Else
                            GDNextCol = GDNextCol + 1
                        End If
                    Case True
                        'Two system approximation, but may break down as equation is "unstable" - values will be <-1,>1 if this is the case
                        Sheets("Geochronology Data").Activate
                        With ActiveSheet
                            Sheets("Geochronology Data").Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=(RC[-1]/RC[-2])/(RC[-3]/RC[-4])"
                            GDNextCol = GDNextCol + 1
                        End With
                        Sheets("Original Data").Activate
                End Select
            Case False
            End Select
        'Inverse isochron ratios
            Select Case RatioPairInversePresent
            Case True
                'Lu176/Hf176
                    Range(Cells(ODStartRowA, Ratio176Lu176HfCol), Cells(ODLastRowA, Ratio176Lu176HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Lu176Hf176"
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Lu176Hf176_" & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio176Lu176HfCol), Cells(ODLastRowB, Ratio176Lu176HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
                'Hf177/Hf176
                    Range(Cells(ODStartRowA, Ratio177Hf176HfCol), Cells(ODLastRowA, Ratio177Hf176HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Hf177Hf176"
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Hf177Hf176_" & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio177Hf176HfCol), Cells(ODLastRowB, Ratio177Hf176HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
            'Copy/calculate error correlation (rho)
                Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rho_Lu176Hf176_Hf177Hf176"
                Select Case RhoCalc
                    Case False
                        If IsNumeric(RhoLuHfCol) Then
                            Range(Cells(ODStartRowA + 1, RhoLuHfCol), Cells(ODLastRowA, RhoLuHfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Else
                            GDNextCol = GDNextCol + 1
                        End If
                    Case True
                        'Two system approximation, but may break down as equation is "unstable" - values will be <-1,>1 if this is the case
                        Sheets("Geochronology Data").Activate
                        With ActiveSheet
                            Sheets("Geochronology Data").Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=(RC[-1]/RC[-2])/(RC[-3]/RC[-4])"
                            GDNextCol = GDNextCol + 1
                        End With
                        Sheets("Original Data").Activate
                End Select
            Case False
            End Select
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
        'Geochronology Data
            Set GDRange = Sheets("Geochronology Data").Range("A1", Cells(GDLastRow, GDLastCol))
                With GDRange
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
            'Rename unknowns sheet
                Sheets("Geochronology Data").Name = "Geochronology Data - Unknowns"
                Sheets("Geochronology Data - Unknowns").Move after:=Sheets("Geochronology Data - Standards")
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
                        Select Case TraceElementDataPresent
                        Case True
                            CPScol1 = Range("A1", Cells(1, LastCol)).Find(what:="Lu175_CPS", MatchCase:=True).Column
                            CPScol2 = Range("A1", Cells(1, LastCol)).Find(what:="Hf178_CPS", MatchCase:=True).Column
                            .Range(Cells(2, CPScol1), Cells(LastRow, CPScol2)).NumberFormat = "0"
                        Case False
                        End Select
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
                        Select Case TraceElementDataPresent
                        Case True
                            CPScol1 = Range("A1", Cells(1, LastCol)).Find(what:="Lu175_CPS", MatchCase:=True).Column
                            CPScol2 = Range("A1", Cells(1, LastCol)).Find(what:="Hf178_CPS", MatchCase:=True).Column
                            .Range(Cells(2, CPScol1), Cells(LastRow, CPScol2)).NumberFormat = "0"
                        Case False
                        End Select
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
SaveError:                     MsgBox "There was an error saving this file during execution of the add-in." & vbCrLf & "This is likley due to a file of the same name already existing" & vbCrLf & vbCrLf & "Please remember to save your file if you haven't already", , "LADR_LuHf_Arranger"
End Sub

Sub LADRLuHfArrangerBatch()
'This will batch process a folder's CSV files, based on the host file opened
    FolderName = Application.ActiveWorkbook.Path
    If Right(FolderName, 1) <> Application.PathSeparator Then FolderName = FolderName & Application.PathSeparator
    Fname = Dir(FolderName & "*.csv")
    'loop through the files
    Do While Len(Fname)
        With Workbooks.Open(FolderName & Fname)
           Call LADRLuHfArranger
           ActiveWorkbook.Close
        End With
        ' Go to the next file in the folder
        Fname = Dir
    Loop
End Sub

Sub LADRLuHfArrangerBatchConfirm(control As IRibbonControl)
''This is to prevent accidentally ruining an excel spreadsheet
    Msg = "Do you wish to batch process all CSV output files (LADR Lu-Hf geochronology and elemental data) in the host folder?" & vbCrLf & vbCrLf & "It is a good idea to move the CSV files you want to process into a separate folder as this will process all CSV files in the host folder and result in execution errors" & vbCrLf & vbCrLf & "Continue?"
    Ans = MsgBox(Msg, vbYesNo)
    Select Case Ans
        Case vbYes
        Call LADRLuHfArrangerBatch
        Case vbNo
        GoTo Quit:
    End Select
Quit:
End Sub
Sub LADRLuHfArrangerConfirm(control As IRibbonControl)
'This is to prevent accidentally running this procedure
    Msg = "This will rearrange and format a LADR CSV output file for QQQ Lu-Hf geochronology and elemental data." & vbCrLf & vbCrLf & "It will save the result as an XLSX with the same name as the input CSV." & vbCrLf & vbCrLf & "It will not overwrite the CSV or any existing XLSX file of the same name in the directory of the CSV." & vbCrLf & vbCrLf & "Do you want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR_LuHf_Arranger")
    Select Case Ans
        Case vbYes
        Call LADRLuHfArranger
        Case vbNo
        GoTo Quit:
    End Select
Quit:
End Sub
Sub LADRLuHfArrangerRhoConfirm(control As IRibbonControl)
'This is to prevent accidentally running this procedure
    Msg = "This will rearrange and format a LADR CSV output file for QQQ Lu-Hf (proxied as UPb for rho) geochronology." & vbCrLf & vbCrLf & "It will save the result as an XLSX with the same name as the input CSV." & vbCrLf & vbCrLf & "It will not overwrite the CSV or any existing XLSX file of the same name in the directory of the CSV." & vbCrLf & vbCrLf & "Do you want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR_LuHf_Arranger_Rho")
    Select Case Ans
        Case vbYes
        Call LADRLuHfArrangerRho
        Case vbNo
        GoTo Quit:
    End Select
Quit:
End Sub
Sub LADRLuHfArrangerRho()
    'Define number and name of standards
DefineStandards:            NumStandards = Application.InputBox("How many different standards were used?", "LADR_LuHf_Arranger", 4, Type:=1)
        Select Case NumStandards
            Case 1
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case 2
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case 3
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case 4
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case 5
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
                Standard5 = Application.InputBox("What is the sample name of the fifth standard as it is shown in the output CSV?", "LADR_LuHf_Arranger", Type:=2)
            Case Else
                Ans = MsgBox("Please enter the number of standards, this has to be a value of 1 to 5." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_LuHf_Arranger")
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
StdCheckError:      Ans = MsgBox("Standards variables are not set correctly." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_LuHf_Arranger")
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
    'Check Element number and symbol order
        If Cells(FirstMassRow, 1).Value Like "#*" Then
            ElementSymNumOrder = "NumSym"
        ElseIf Cells(FirstMassRow, 1).Value Like "[A-Z]*" Then
            ElementSymNumOrder = "SymNum"
        End If
        Select Case ElementSymNumOrder
            Case "NumSym"
            'Check and define 238U/206Pb ratio (238U/206Pb as proxy for Lu176/Hf177 or Lu176/Hf176 to determine error correlation in LADR)
                Set Ratio176Lu177HfCol = HeaderRange.Find(what:="238U/206Pb", MatchCase:=False)
                If Not Ratio176Lu177HfCol Is Nothing Then
                    Ratio176Lu177HfCol = HeaderRange.Find(what:="238U/206Pb", MatchCase:=False).Column
                    CheckRatio176Lu177Hf = True
                Else
                    CheckRatio176Lu177Hf = False
                End If
            'Check and define 207Pb/206Pb ratio (207Pb/206Pb as proxy for Hf176/Hf177 or Hf177/Hf176 to determine error correlation in LADR)
                Set Ratio176Hf177HfCol = HeaderRange.Find(what:="207Pb/206Pb", MatchCase:=False)
                If Not Ratio176Hf177HfCol Is Nothing Then
                    Ratio176Hf177HfCol = HeaderRange.Find(what:="207Pb/206Pb", MatchCase:=False).Column
                    CheckRatio176Hf177Hf = True
                Else
                    CheckRatio176Hf177Hf = False
                End If
            'Check that at least one ratio pair (normal or inverse isochron) is present
                If CheckRatio176Lu177Hf = True And CheckRatio176Hf177Hf = True Then
                    RatioPairNormalPresent = True
                Else
                    RatioPairNormalPresent = False
                End If
                If RatioPairNormalPresent = False Then
                    MsgBox ("The 238U/206Pb and 207Pb/206Pb ratio pair is required for this procedure to continue." & vbCrLf & vbCrLf & "Please check that both are present." & vbCrLf & vbCrLf & "Procedure ended.")
                    Exit Sub
                Else
                End If
            Case "SymNum"
            'Check and define 238U/206Pb ratio (238U/206Pb as proxy for Lu176/Hf177 or Lu176/Hf176 to determine error correlation in LADR)
                Set Ratio176Lu177HfCol = HeaderRange.Find(what:="U238/Pb206", MatchCase:=False)
                If Not Ratio176Lu177HfCol Is Nothing Then
                    Ratio176Lu177HfCol = HeaderRange.Find(what:="U238/Pb206", MatchCase:=False).Column
                    CheckRatio176Lu177Hf = True
                Else
                    CheckRatio176Lu177Hf = False
                End If
            'Check and define 207Pb/206Pb ratio (207Pb/206Pb as proxy for Hf176/Hf177 or Hf177/Hf176 to determine error correlation in LADR)
                Set Ratio176Hf177HfCol = HeaderRange.Find(what:="Pb207/Pb206", MatchCase:=False)
                If Not Ratio176Hf177HfCol Is Nothing Then
                    Ratio176Hf177HfCol = HeaderRange.Find(what:="Pb207/Pb206", MatchCase:=False).Column
                    CheckRatio176Hf177Hf = True
                Else
                    CheckRatio176Hf177Hf = False
                End If
            'Check that at least one ratio pair (normal or inverse isochron) is present
                If CheckRatio176Lu177Hf = True And CheckRatio176Hf177Hf = True Then
                    RatioPairNormalPresent = True
                Else
                    RatioPairNormalPresent = False
                End If
                If RatioPairNormalPresent = False Then
                    MsgBox ("The 238U/206Pb and 207Pb/206Pb ratio pair is required for this procedure to continue." & vbCrLf & vbCrLf & "Please check that both are present." & vbCrLf & vbCrLf & "Procedure ended.")
                    Exit Sub
                Else
                End If
            End Select
            'Check and define error correlations (U/Pb systems used to force LADR to calculate error correlation, otherwise two system approach is used but is likely to be unstable)
                Set RhoLuHfCol = HeaderRange.Find(what:="Rho: 207/206 vs 238/206", MatchCase:=False, lookat:=xlPart)
                If Not RhoLuHfCol Is Nothing Then
                    RhoCalc = False
                    RhoLuHfCol = HeaderRange.Find(what:="Rho: 207/206 vs 238/206", MatchCase:=False, lookat:=xlPart).Column
                ElseIf SignalPrecision = True Then
                    RhoCalc = True
                Else
                    RhoCalc = False
                End If
        End With
    'Add new sheets
        Sheets.Add.Name = "Geochronology Data"
        Sheets("Original Data").Activate
    'Copy AL#, sample name and analysis number
        Range(Cells(ODStartRowA, ALNumCol), Cells(ODLastRowA, ALNumCol)).Copy Destination:=Sheets("Geochronology Data").Range("A1")
        Range(Cells(ODStartRowA, SampleCol), Cells(ODLastRowA, AnalysisCol)).Copy Destination:=Sheets("Geochronology Data").Range("B1")
    'Copy geochronology data, suitable for isoplotR input
        'Define GDNextCol, GDLastRow
            GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
            GDNextCol = GDLastCol + 1
            GDLastRow = Sheets("Geochronology Data").Cells(Rows.Count, 1).End(xlUp).Row
        'Copy ratios and uncertainties, label uncertainty columns
        'Normal isochron ratios
            Select Case RatioPairNormalPresent
            Case True
                'Lu176/Hf177
                    Range(Cells(ODStartRowA, Ratio176Lu177HfCol), Cells(ODLastRowA, Ratio176Lu177HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[" & Sheets("Geochronology Data").Cells(1, GDNextCol - 1).Value & "] " & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio176Lu177HfCol), Cells(ODLastRowB, Ratio176Lu177HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
                'Hf176/Hf177
                    Range(Cells(ODStartRowA, Ratio176Hf177HfCol), Cells(ODLastRowA, Ratio176Hf177HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[" & Sheets("Geochronology Data").Cells(1, GDNextCol - 1).Value & "] " & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio176Hf177HfCol), Cells(ODLastRowB, Ratio176Hf177HfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
            'Copy/calculate error correlation (rho)
                Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rho[Lu/Hf][Hf/Hf]"
                Select Case RhoCalc
                    Case False
                        If IsNumeric(RhoLuHfCol) Then
                            Range(Cells(ODStartRowA + 1, RhoLuHfCol), Cells(ODLastRowA, RhoLuHfCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                            GDNextCol = GDNextCol + 1
                        Else
                            GDNextCol = GDNextCol + 1
                        End If
                    Case True
                        'Two system approximation, but may break down as equation is "unstable" - values will be <-1,>1 if this is the case
                        Sheets("Geochronology Data").Activate
                        With ActiveSheet
                            Sheets("Geochronology Data").Range(Cells(2, GDNextCol), Cells(GDLastRow, GDNextCol)).FormulaR1C1 = "=(RC[-1]/RC[-2])/(RC[-3]/RC[-4])"
                            GDNextCol = GDNextCol + 1
                        End With
                        Sheets("Original Data").Activate
                End Select
            Case False
            End Select
        'Copy Comments
            Range(Cells(ODStartRowA, CommentsCol), Cells(ODLastRowA, CommentsCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
            GDNextCol = GDNextCol + 1
        'Copy Source Filename
            Range(Cells(ODStartRowA, SourceFileCol), Cells(ODLastRowA, SourceFileCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
            SFColDel = GDNextCol
    Sheets("Geochronology Data").Activate 'Correct sample label and trailing number for correct sorting in Excel
            GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
            GDLastRow = Sheets("Geochronology Data").Cells(Rows.Count, 1).End(xlUp).Row
            For n = 2 To GDLastRow
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
    'Sort prior to cut - performance optimisation
        'Delete sourcefile column
            Columns(SFColDel).Delete
        'Set GD Last Column
            GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
        'Geochronology Data
            Set GDRange = Sheets("Geochronology Data").Range("A1", Cells(GDLastRow, GDLastCol))
                With GDRange
                .Sort Key1:=Range("C1"), order1:=xlAscending, Header:=xlYes
            End With
    'Separate Standards and Unknowns
        'Add standards worksheets
            Sheets.Add.Name = "Geochronology Data - Standards"
        'Define GD and ED variables
            GDLastRow = Sheets("Geochronology Data").Range("A" & Rows.Count).End(xlUp).Row
            GDLastCol = Sheets("Geochronology Data").Cells(1, Columns.Count).End(xlToLeft).Column
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
            'Rename unknowns sheet
                Sheets("Geochronology Data").Name = "Geochronology Data - Unknowns"
                Sheets("Geochronology Data - Unknowns").Move after:=Sheets("Geochronology Data - Standards")
    'Format and reset performance optimisations
        'Enable screen updating for pane freeze
            Application.ScreenUpdating = True
            Application.EnableEvents = True
        'Format
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
SaveError:                     MsgBox "There was an error saving this file during execution of the add-in." & vbCrLf & "This is likley due to a file of the same name already existing" & vbCrLf & vbCrLf & "Please remember to save your file if you haven't already", , "LADR_LuHf_Arranger"
End Sub

