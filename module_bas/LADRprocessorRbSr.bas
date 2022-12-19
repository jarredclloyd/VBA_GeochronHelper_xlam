Attribute VB_Name = "LADRprocessorRbSr"
'Module for handling LADR Rb-Sr geochronology data outputs
'Created By Jarred Lloyd on 2020-03-08
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
        Dim RhoRbSrCol As Variant
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
        Dim Ratio87Rb86SrCol As Variant
        Dim Ratio87Sr86SrCol As Variant
        Dim Ratio87Rb87SrCol As Variant
        Dim Ratio86Sr87SrCol As Variant
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
        Dim CheckRatio87Rb86Sr As Boolean
        Dim CheckRatio87Sr86Sr As Boolean
        Dim CheckRatio87Rb87Sr As Boolean
        Dim CheckRatio86Sr87Sr As Boolean
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
Sub LADRRbSrArranger()
    'Define number and name of standards
DefineStandards:            NumStandards = Application.InputBox("How many different standards were used?", "LADR_RbSr_Arranger", 4, Type:=1)
        Select Case NumStandards
            Case 1
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case 2
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case 3
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case 4
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case 5
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard5 = Application.InputBox("What is the sample name of the fifth standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case Else
                Ans = MsgBox("Please enter the number of standards, this has to be a value of 1 to 5." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_RbSr_Arranger")
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
StdCheckError:      Ans = MsgBox("Standards variables are not set correctly." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_RbSr_Arranger")
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
                            CPScol1 = HeaderRange.Find(what:="85Rb", MatchCase:=False, lookat:=xlPart).Column
                            Set Ratio87Rb86SrCol = HeaderRange.Find(what:="85Rb/86Sr->102", MatchCase:=False)
                            If Not Ratio87Rb86SrCol Is Nothing Then
                                CPScol2 = HeaderRange.Find(what:="86Sr", MatchCase:=False, lookat:=xlPart).Column
                                CommonIsotope = "Sr86"
                            Else
                                CPScol2 = HeaderRange.Find(what:="88Sr", MatchCase:=False, lookat:=xlPart).Column
                                CommonIsotope = "Sr88"
                            End If
                            CPScol3 = HeaderRange.Find(what:="87Sr", MatchCase:=False, lookat:=xlPart).Column
                        Case False
                    End Select
                'Check and define 85Rb/86Sr ratio (Rb85 used as proxy to measure Rb87)
                    Set Ratio87Rb86SrCol = HeaderRange.Find(what:="85Rb/86Sr->102", MatchCase:=False)
                    If Not Ratio87Rb86SrCol Is Nothing Then
                        Ratio87Rb86SrCol = HeaderRange.Find(what:="85Rb/86Sr->102", MatchCase:=False).Column
                        CheckRatio87Rb86Sr = True
                    Else
                        Set Ratio87Rb86SrCol = HeaderRange.Find(what:="85Rb/86Sr", MatchCase:=False)
                        If Not Ratio87Rb86SrCol Is Nothing Then
                            Ratio87Rb86SrCol = HeaderRange.Find(what:="85Rb/86Sr", MatchCase:=False).Column
                            CheckRatio87Rb86Sr = True
                        Else
                            Set Ratio87Rb86SrCol = HeaderRange.Find(what:="85Rb/88Sr->104", MatchCase:=False)
                            If Not Ratio87Rb86SrCol Is Nothing Then
                                Ratio87Rb86SrCol = HeaderRange.Find(what:="85Rb/88Sr->104", MatchCase:=False).Column
                                CheckRatio87Rb86Sr = True
                            Else
                                CheckRatio87Rb86Sr = False
                            End If
                        End If
                    End If
                'Check and define 87Sr/86Sr ratio
                    Set Ratio87Sr86SrCol = HeaderRange.Find(what:="87Sr->103/86Sr->102", MatchCase:=False)
                    If Not Ratio87Sr86SrCol Is Nothing Then
                        Ratio87Sr86SrCol = HeaderRange.Find(what:="87Sr->103/86Sr->102", MatchCase:=False).Column
                        CheckRatio87Sr86Sr = True
                    Else
                        Set Ratio87Sr86SrCol = HeaderRange.Find(what:="87Sr/86Sr", MatchCase:=False)
                        If Not Ratio87Sr86SrCol Is Nothing Then
                            Ratio87Sr86SrCol = HeaderRange.Find(what:="87Sr/86Sr", MatchCase:=False).Column
                            CheckRatio87Sr86Sr = True
                        Else
                            Set Ratio87Sr86SrCol = HeaderRange.Find(what:="87Sr->103/88Sr->104", MatchCase:=False)
                            If Not Ratio87Sr86SrCol Is Nothing Then
                                Ratio87Sr86SrCol = HeaderRange.Find(what:="87Sr->103/88Sr->104", MatchCase:=False).Column
                                CheckRatio87Sr86Sr = True
                            Else
                                CheckRatio87Sr86Sr = False
                            End If
                        End If
                    End If
                'Check and define 85Rb/87Sr ratio (Rb85 used as proxy to measure Rb87)
                    Set Ratio87Rb87SrCol = HeaderRange.Find(what:="85Rb/87Sr->103", MatchCase:=False)
                    If Not Ratio87Rb87SrCol Is Nothing Then
                        Ratio87Rb87SrCol = HeaderRange.Find(what:="85Rb/87Sr->103", MatchCase:=False).Column
                        CheckRatio87Rb87Sr = True
                    Else
                        Set Ratio87Rb87SrCol = HeaderRange.Find(what:="85Rb/87Sr", MatchCase:=False)
                        If Not Ratio87Rb87SrCol Is Nothing Then
                            Ratio87Rb87SrCol = HeaderRange.Find(what:="85Rb/87Sr", MatchCase:=False).Column
                            CheckRatio87Rb87Sr = True
                        Else
                            CheckRatio87Rb87Sr = False
                        End If
                    End If
                'Check and define 86Sr/867r ratio
                    Set Ratio86Sr87SrCol = HeaderRange.Find(what:="86Sr->102/87Sr->103", MatchCase:=False)
                    If Not Ratio86Sr87SrCol Is Nothing Then
                        Ratio86Sr87SrCol = HeaderRange.Find(what:="86Sr->102/87Sr->103", MatchCase:=False).Column
                        CheckRatio86Sr87Sr = True
                    Else
                        Set Ratio86Sr87SrCol = HeaderRange.Find(what:="86Sr/87Sr", MatchCase:=False)
                        If Not Ratio86Sr87SrCol Is Nothing Then
                            Ratio86Sr87SrCol = HeaderRange.Find(what:="86Sr/87Sr", MatchCase:=False).Column
                            CheckRatio86Sr87Sr = True
                        Else
                            Set Ratio86Sr87SrCol = HeaderRange.Find(what:="88Sr->104/87Sr->103", MatchCase:=False)
                            If Not Ratio86Sr87SrCol Is Nothing Then
                            Ratio86Sr87SrCol = HeaderRange.Find(what:="88Sr->104/87Sr->103", MatchCase:=False).Column
                            CheckRatio86Sr87Sr = True
                            Else
                                CheckRatio86Sr87Sr = False
                            End If
                        End If
                    End If
            Case "SymNum"
                'CPS Columns
                    Select Case TraceElementDataPresent
                        Case True
                            CPScol1 = HeaderRange.Find(what:="Rb85", MatchCase:=False, lookat:=xlPart).Column
                            Set Ratio87Rb86SrCol = HeaderRange.Find(what:="Rb85/Sr86->102", MatchCase:=False)
                            If Not Ratio87Rb86SrCol Is Nothing Then
                                CPScol2 = HeaderRange.Find(what:="Sr86", MatchCase:=False, lookat:=xlPart).Column
                                CommonIsotope = "Sr86"
                            Else
                                CPScol2 = HeaderRange.Find(what:="Sr88", MatchCase:=False, lookat:=xlPart).Column
                                CommonIsotope = "Sr88"
                            End If
                            CPScol3 = HeaderRange.Find(what:="Sr87", MatchCase:=False, lookat:=xlPart).Column
                        Case False
                    End Select
                'Check and define 85Rb/86Sr ratio (Rb85 used as proxy to measure Rb87)
                    Set Ratio87Rb86SrCol = HeaderRange.Find(what:="Rb85/Sr86->102", MatchCase:=False)
                    If Not Ratio87Rb86SrCol Is Nothing Then
                        Ratio87Rb86SrCol = HeaderRange.Find(what:="Rb85/Sr86->102", MatchCase:=False).Column
                        CheckRatio87Rb86Sr = True
                    Else
                        Set Ratio87Rb86SrCol = HeaderRange.Find(what:="Rb85/Sr86", MatchCase:=False)
                        If Not Ratio87Rb86SrCol Is Nothing Then
                            Ratio87Rb86SrCol = HeaderRange.Find(what:="Rb85/Sr86", MatchCase:=False).Column
                            CheckRatio87Rb86Sr = True
                        Else
                            Set Ratio87Rb86SrCol = HeaderRange.Find(what:="Rb85/Sr88->104", MatchCase:=False)
                            If Not Ratio87Rb86SrCol Is Nothing Then
                                Ratio87Rb86SrCol = HeaderRange.Find(what:="Rb85/Sr88->104", MatchCase:=False).Column
                                CheckRatio87Rb86Sr = True
                            Else
                                CheckRatio87Rb86Sr = False
                            End If
                        End If
                    End If
                'Check and define 87Sr/86Sr ratio
                    Set Ratio87Sr86SrCol = HeaderRange.Find(what:="Sr87->103/Sr86->102", MatchCase:=False)
                    If Not Ratio87Sr86SrCol Is Nothing Then
                        Ratio87Sr86SrCol = HeaderRange.Find(what:="Sr87->103/Sr86->102", MatchCase:=False).Column
                        CheckRatio87Sr86Sr = True
                    Else
                        Set Ratio87Sr86SrCol = HeaderRange.Find(what:="Sr87/Sr86", MatchCase:=False)
                        If Not Ratio87Sr86SrCol Is Nothing Then
                            Ratio87Sr86SrCol = HeaderRange.Find(what:="Sr87/Sr86", MatchCase:=False).Column
                            CheckRatio87Sr86Sr = True
                        Else
                            Set Ratio87Sr86SrCol = HeaderRange.Find(what:="Sr87->103/Sr88->104", MatchCase:=False)
                            If Not Ratio87Sr86SrCol Is Nothing Then
                                Ratio87Sr86SrCol = HeaderRange.Find(what:="Sr87->103/Sr88->104", MatchCase:=False).Column
                                CheckRatio87Sr86Sr = True
                            Else
                                CheckRatio87Sr86Sr = False
                            End If
                        End If
                    End If
                'Check and define 85Rb/87Sr ratio (Rb85 used as proxy to measure Rb87)
                    Set Ratio87Rb87SrCol = HeaderRange.Find(what:="Rb85/Sr87->103", MatchCase:=False)
                    If Not Ratio87Rb87SrCol Is Nothing Then
                        Ratio87Rb87SrCol = HeaderRange.Find(what:="Rb85/Sr87->103", MatchCase:=False).Column
                        CheckRatio87Rb87Sr = True
                    Else
                        Set Ratio87Rb87SrCol = HeaderRange.Find(what:="Rb85/Sr87", MatchCase:=False)
                        If Not Ratio87Rb87SrCol Is Nothing Then
                            Ratio87Rb87SrCol = HeaderRange.Find(what:="Rb85/Sr87", MatchCase:=False).Column
                            CheckRatio87Rb87Sr = True
                        Else
                            CheckRatio87Rb87Sr = False
                        End If
                    End If
                'Check and define 86Sr/867r ratio
                    Set Ratio86Sr87SrCol = HeaderRange.Find(what:="Sr86->102/Sr87->103", MatchCase:=False)
                    If Not Ratio86Sr87SrCol Is Nothing Then
                        Ratio86Sr87SrCol = HeaderRange.Find(what:="Sr86->102/Sr87->103", MatchCase:=False).Column
                        CheckRatio86Sr87Sr = True
                    Else
                        Set Ratio86Sr87SrCol = HeaderRange.Find(what:="Sr86/Sr87", MatchCase:=False)
                        If Not Ratio86Sr87SrCol Is Nothing Then
                            Ratio86Sr87SrCol = HeaderRange.Find(what:="Sr86/Sr87", MatchCase:=False).Column
                            CheckRatio86Sr87Sr = True
                        Else
                            Set Ratio86Sr87SrCol = HeaderRange.Find(what:="Sr88->104/Sr87->103", MatchCase:=False)
                            If Not Ratio86Sr87SrCol Is Nothing Then
                            Ratio86Sr87SrCol = HeaderRange.Find(what:="Sr88->104/Sr87->103", MatchCase:=False).Column
                            CheckRatio86Sr87Sr = True
                            Else
                                CheckRatio86Sr87Sr = False
                            End If
                        End If
                    End If
            End Select
            'Check that at least one ratio pair (normal or inverse isochron) is present
                If CheckRatio87Rb86Sr = True And CheckRatio87Sr86Sr = True Then
                    RatioPairNormalPresent = True
                Else
                    RatioPairNormalPresent = False
                End If
                If CheckRatio87Rb87Sr = True And CheckRatio86Sr87Sr = True Then
                    RatioPairInversePresent = True
                Else
                    RatioPairInversePresent = False
                End If
                If RatioPairNormalPresent = False And RatioPairInversePresent = False Then
                    MsgBox ("At least one pair of ratios is required for this procedure to continue." & vbCrLf & vbCrLf & "Please check that either 85Rb/86Sr AND 87Sr/86Sr OR 85Rb/87Sr AND 86Sr/87Sr are present in the data. These can be in alias forms like 85Rb/87Sr->103." & vbCrLf & vbCrLf & "Procedure ended.")
                    Exit Sub
                Else
                End If
            'Check and define error correlations (U/Pb systems used to force LADR to calculate error correlation, otherwise two system approach is used but is likely to be unstable)
                Set RhoRbSrCol = HeaderRange.Find(what:="Rho: 87Rb/86Sr vs 87Sr/86Sr", MatchCase:=False, lookat:=xlPart)
                If Not RhoRbSrCol Is Nothing Then
                    RhoCalc = False
                    RhoRbSrCol = HeaderRange.Find(what:="Rho: 87Rb/87Sr vs 86Sr/87Sr", MatchCase:=False, lookat:=xlPart).Column
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
    'Copy elemental data
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
                'Rb85
                    Range(Cells(ODStartRowC, CPScol1), Cells(ODLastRowC, CPScol1)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rb85_CPS"
                    GDNextCol = GDNextCol + 1
                'Sr86 or Sr88 (depending on what was used for ratio)
                    Range(Cells(ODStartRowC, CPScol2), Cells(ODLastRowC, CPScol2)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = CommonIsotope & "_CPS"
                    GDNextCol = GDNextCol + 1
                'Sr87
                    Range(Cells(ODStartRowC, CPScol3), Cells(ODLastRowC, CPScol3)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Sr87_CPS"
                    GDNextCol = GDNextCol + 1
                Case False
            End Select
        'Copy ratios and uncertainties, label uncertainty columns
        'Normal isochron ratios
            Select Case RatioPairNormalPresent
            Case True
                '87Rb/86Sr
                    Range(Cells(ODStartRowA, Ratio87Rb86SrCol), Cells(ODLastRowA, Ratio87Rb86SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rb87Sr86"
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rb87Sr86_" & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio87Rb86SrCol), Cells(ODLastRowB, Ratio87Rb86SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
                '87Sr/86Sr
                    Range(Cells(ODStartRowA, Ratio87Sr86SrCol), Cells(ODLastRowA, Ratio87Sr86SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Sr87Sr86"
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Sr87Sr86_" & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio87Sr86SrCol), Cells(ODLastRowB, Ratio87Sr86SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
            'Copy/calculate error correlation (rho)
                Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rho_Rb87Sr86_Sr87Sr86]"
                Select Case RhoCalc
                    Case False
                        If IsNumeric(RhoRbSrCol) Then
                            Range(Cells(ODStartRowA + 1, RhoRbSrCol), Cells(ODLastRowA, RhoRbSrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
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
                '85Rb(87Rb)/87Sr
                    Range(Cells(ODStartRowA, Ratio87Rb87SrCol), Cells(ODLastRowA, Ratio87Rb87SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rb87Sr87"
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rb87Sr87_" & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio87Rb87SrCol), Cells(ODLastRowB, Ratio87Rb87SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
                '86Sr/87Sr
                    Range(Cells(ODStartRowA, Ratio86Sr87SrCol), Cells(ODLastRowA, Ratio86Sr87SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Sr86Sr87"
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Sr86Sr87_" & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio86Sr87SrCol), Cells(ODLastRowB, Ratio86Sr87SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
            'Copy/calculate error correlation (rho)
                Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rho_Rb87Sr87_Sr86Sr87"
                Select Case RhoCalc
                    Case False
                        If IsNumeric(RhoRbSrCol) Then
                            Range(Cells(ODStartRowA + 1, RhoRbSrCol), Cells(ODLastRowA, RhoRbSrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
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
                            CPScol1 = Range("A1", Cells(1, LastCol)).Find(what:="Rb85_CPS", MatchCase:=True).Column
                            CPScol2 = Range("A1", Cells(1, LastCol)).Find(what:="Sr87_CPS", MatchCase:=True).Column
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
                            CPScol1 = Range("A1", Cells(1, LastCol)).Find(what:="Rb85_CPS", MatchCase:=True).Column
                            CPScol2 = Range("A1", Cells(1, LastCol)).Find(what:="Sr87_CPS", MatchCase:=True).Column
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
SaveError:                     MsgBox "There was an error saving this file during execution of the add-in." & vbCrLf & "This is likley due to a file of the same name already existing" & vbCrLf & vbCrLf & "Please remember to save your file if you haven't already", , "LADR_RbSr_Arranger"
End Sub

Sub LADRRbSrArrangerBatch()
'This will batch process a folder's CSV files, based on the host file opened
    FolderName = Application.ActiveWorkbook.Path
    If Right(FolderName, 1) <> Application.PathSeparator Then FolderName = FolderName & Application.PathSeparator
    Fname = Dir(FolderName & "*.csv")
    'loop through the files
    Do While Len(Fname)
        With Workbooks.Open(FolderName & Fname)
           Call LADRRbSrArranger
           ActiveWorkbook.Close
        End With
        ' Go to the next file in the folder
        Fname = Dir
    Loop
End Sub

Sub LADRRbSrArrangerBatchConfirm(control As IRibbonControl)
''This is to prevent accidentally ruining an excel spreadsheet
    Msg = "Do you wish to batch process all CSV output files (LADR Rb-Sr with elemental data) in the host folder?" & vbCrLf & vbCrLf & "It is a good idea to move the CSV files you want to process into a separate folder as this will process all CSV files in the host folder and result in execution errors" & vbCrLf & vbCrLf & "Continue?"
    Ans = MsgBox(Msg, vbYesNo)
    Select Case Ans
        Case vbYes
        Call LADRRbSrArrangerBatch
        Case vbNo
        GoTo Quit:
    End Select
Quit:
End Sub
Sub LADRRbSrArrangerConfirm(control As IRibbonControl)
'This is to prevent accidentally running this procedure
    Msg = "This will rearrange and format a LADR CSV output file for QQQ Rb-Sr geochronology and elemental data." & vbCrLf & vbCrLf & "It will save the result as an XLSX with the same name as the input CSV." & vbCrLf & vbCrLf & "It will not overwrite the CSV or any existing XLSX file of the same name in the directory of the CSV." & vbCrLf & vbCrLf & "Do you want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR_RbSr_Arranger")
    Select Case Ans
        Case vbYes
        Call LADRRbSrArranger
        Case vbNo
        GoTo Quit:
    End Select
Quit:
End Sub
Sub LADRRbSrArrangerRhoConfirm(control As IRibbonControl)
'This is to prevent accidentally running this procedure
    Msg = "This will rearrange and format a LADR CSV output file for QQQ Rb-Sr (proxied as UPb for rho) geochronology." & vbCrLf & vbCrLf & "It will save the result as an XLSX with the same name as the input CSV." & vbCrLf & vbCrLf & "It will not overwrite the CSV or any existing XLSX file of the same name in the directory of the CSV." & vbCrLf & vbCrLf & "Do you want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR_RbSr_Arranger_Rho")
    Select Case Ans
        Case vbYes
        Call LADRRbSrArrangerRho
        Case vbNo
        GoTo Quit:
    End Select
Quit:
End Sub
Sub LADRRbSrArrangerRho()
    'Define number and name of standards
DefineStandards:            NumStandards = Application.InputBox("How many different standards were used?", "LADR_RbSr_Arranger", 4, Type:=1)
        Select Case NumStandards
            Case 1
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case 2
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case 3
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case 4
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case 5
                Standard1 = Application.InputBox("What is the sample name of the first standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard2 = Application.InputBox("What is the sample name of the second standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard3 = Application.InputBox("What is the sample name of the third standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard4 = Application.InputBox("What is the sample name of the fourth standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
                Standard5 = Application.InputBox("What is the sample name of the fifth standard as it is shown in the output CSV?", "LADR_RbSr_Arranger", Type:=2)
            Case Else
                Ans = MsgBox("Please enter the number of standards, this has to be a value of 1 to 5." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_RbSr_Arranger")
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
StdCheckError:      Ans = MsgBox("Standards variables are not set correctly." & vbCrLf & "Do you want to continue?", vbYesNo, "LADR_RbSr_Arranger")
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
                'Check and define 238U/206Pb ratio (238U/206Pb as proxy for 85Rb/86Sr or 85Rb/87Sr to determine error correlation in LADR)
                    Set Ratio87Rb86SrCol = HeaderRange.Find(what:="238U/206Pb", MatchCase:=False)
                    If Not Ratio87Rb86SrCol Is Nothing Then
                        Ratio87Rb86SrCol = HeaderRange.Find(what:="238U/206Pb", MatchCase:=False).Column
                        CheckRatio87Rb86Sr = True
                    Else
                        CheckRatio87Rb86Sr = False
                    End If
                'Check and define 207Pb/206Pb ratio (207Pb/206Pb as proxy for 87Sr/86Sr or 86Sr/87Sr to determine error correlation in LADR)
                    Set Ratio87Sr86SrCol = HeaderRange.Find(what:="207Pb/206Pb", MatchCase:=False)
                    If Not Ratio87Sr86SrCol Is Nothing Then
                        Ratio87Sr86SrCol = HeaderRange.Find(what:="207Pb/206Pb", MatchCase:=False).Column
                        CheckRatio87Sr86Sr = True
                    Else
                        CheckRatio87Sr86Sr = False
                    End If
            Case "SymNum"
                'Check and define 238U/206Pb ratio (238U/206Pb as proxy for 85Rb/86Sr or 85Rb/87Sr to determine error correlation in LADR)
                    Set Ratio87Rb86SrCol = HeaderRange.Find(what:="U238/Pb206", MatchCase:=False)
                    If Not Ratio87Rb86SrCol Is Nothing Then
                        Ratio87Rb86SrCol = HeaderRange.Find(what:="U238/Pb206", MatchCase:=False).Column
                        CheckRatio87Rb86Sr = True
                    Else
                        CheckRatio87Rb86Sr = False
                    End If
                'Check and define 207Pb/206Pb ratio (207Pb/206Pb as proxy for 87Sr/86Sr or 86Sr/87Sr to determine error correlation in LADR)
                    Set Ratio87Sr86SrCol = HeaderRange.Find(what:="Pb207/Pb206", MatchCase:=False)
                    If Not Ratio87Sr86SrCol Is Nothing Then
                        Ratio87Sr86SrCol = HeaderRange.Find(what:="Pb207/Pb206", MatchCase:=False).Column
                        CheckRatio87Sr86Sr = True
                    Else
                        CheckRatio87Sr86Sr = False
                    End If
            End Select
            'Check that at least one ratio pair (normal or inverse isochron) is present
                If CheckRatio87Rb86Sr = True And CheckRatio87Sr86Sr = True Then
                    RatioPairNormalPresent = True
                Else
                    RatioPairNormalPresent = False
                End If
                If RatioPairNormalPresent = False Then
                    MsgBox ("The 238U/206Pb and 207Pb/206Pb ratio pair is required for this procedure to continue." & vbCrLf & vbCrLf & "Please check that both are present." & vbCrLf & vbCrLf & "Procedure ended.")
                    Exit Sub
                Else
                End If
            'Check and define error correlations (U/Pb systems used to force LADR to calculate error correlation, otherwise two system approach is used but is likely to be unstable)
                Set RhoRbSrCol = HeaderRange.Find(what:="Rho: 207/206 vs 238/206", MatchCase:=False, lookat:=xlPart)
                If Not RhoRbSrCol Is Nothing Then
                    RhoCalc = False
                    RhoRbSrCol = HeaderRange.Find(what:="Rho: 207/206 vs 238/206", MatchCase:=False, lookat:=xlPart).Column
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
                '85Rb/86Sr
                    Range(Cells(ODStartRowA, Ratio87Rb86SrCol), Cells(ODLastRowA, Ratio87Rb86SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[" & Sheets("Geochronology Data").Cells(1, GDNextCol - 1).Value & "] " & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio87Rb86SrCol), Cells(ODLastRowB, Ratio87Rb86SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
                '87Sr/86Sr
                    Range(Cells(ODStartRowA, Ratio87Sr86SrCol), Cells(ODLastRowA, Ratio87Sr86SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(1, GDNextCol)
                    GDNextCol = GDNextCol + 1
                    Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Uncertainty[" & Sheets("Geochronology Data").Cells(1, GDNextCol - 1).Value & "] " & StandardErrorLevel & "SE"
                    Range(Cells(ODStartRowB, Ratio87Sr86SrCol), Cells(ODLastRowB, Ratio87Sr86SrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
                    GDNextCol = GDNextCol + 1
            'Copy/calculate error correlation (rho)
                Sheets("Geochronology Data").Cells(1, GDNextCol).Value = "Rho[Rb/Sr][Sr/Sr]"
                Select Case RhoCalc
                    Case False
                        If IsNumeric(RhoRbSrCol) Then
                            Range(Cells(ODStartRowA + 1, RhoRbSrCol), Cells(ODLastRowA, RhoRbSrCol)).Copy Destination:=Sheets("Geochronology Data").Cells(2, GDNextCol)
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
SaveError:                     MsgBox "There was an error saving this file during execution of the add-in." & vbCrLf & "This is likley due to a file of the same name already existing" & vbCrLf & vbCrLf & "Please remember to save your file if you haven't already", , "LADR_RbSr_Arranger"
End Sub
