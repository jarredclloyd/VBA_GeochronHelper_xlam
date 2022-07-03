Attribute VB_Name = "ReOs"
Option Explicit
Option Compare Text

'Variable declaration
    Dim FolderName As String
    Dim Fname As String
    Dim DestName As String
    Dim FSO As Object
    Dim FSOFileName As Object
    Dim FSOFolderName As Object
    Dim Mrow As Long
    Dim LCol As Long
    Dim n As Long
    Dim Mass As String
    Dim Msg As String
    Dim Ans As Variant
    Dim LRow As Long
    Dim FRow As Long
    Dim Re185 As Long
    Dim Re199 As Long
    Dim Os201 As Long
    Dim Os187Corr As Long
    Dim Re187Os187 As Long
    Dim Os187Re187 As Long
    Dim SearchRng As Range
    Dim Age As Long

Private Sub ReOsQQQHeaderCorrEquations()

With ActiveSheet
        'Determine mass header row
            Mrow = .Range("A:A").Find(what:="Time [Sec]", MatchCase:=True, LookAt:=xlWhole).Row
        'Determine last column
            LCol = Cells(Mrow, Columns.Count).End(xlToLeft).Column
            
    'Change shifted headers
    For n = 2 To LCol
        If Cells(Mrow, n).Value = "Lu175 -> 257" Then
            Cells(Mrow, n).Value = "Lu179"
        ElseIf Cells(Mrow, n).Value = "Hf178 -> 260" Then
            Cells(Mrow, n).Value = "Hf177"
        ElseIf Cells(Mrow, n).Value = "Re185 -> 199" Then
            Cells(Mrow, n).Value = "Re199"
        ElseIf Cells(Mrow, n).Value = "Os187 -> 201" Then
            Cells(Mrow, n).Value = "Os201"
        ElseIf Cells(Mrow, n).Value = "Os189 -> 203" Then
            Cells(Mrow, n).Value = "Os203"
        ElseIf Cells(Mrow, n).Value Like "*-*" Then
            Mass = Cells(Mrow, n).Value
            Mass = Left(Mass, InStrRev(Mass, "-") - 2)
            Cells(Mrow, n).Value = Mass
        Else
            GoTo NextN
        End If
NextN:    Next n

    'add calculated columns
        LCol = LCol + 1
        Cells(Mrow, LCol).Value = "Os187(corr)"
        LCol = LCol + 1
        Cells(Mrow, LCol).Value = "Re187_187Os"
        LCol = LCol + 1
        Cells(Mrow, LCol).Value = "Os187_Re187"
        LCol = LCol + 1
        'the next column is for age calculations, but the header is put as a dummy value (Cn277) as iolite won't accept Age(Ma) in the import.
        Cells(Mrow, LCol).Value = "Cn277"
        
    'determine columns and rows
        Set SearchRng = Range(Cells(Mrow, 2), Cells(Mrow, LCol))
        Re185 = SearchRng.Find(what:="Re185", MatchCase:=False).Column
        Re199 = SearchRng.Find(what:="Re199", MatchCase:=False).Column
        Os201 = SearchRng.Find(what:="Os201", MatchCase:=False).Column
        Os187Corr = SearchRng.Find(what:="Os187(corr)", MatchCase:=False).Column
        Re187Os187 = SearchRng.Find(what:="Re187_187Os", MatchCase:=False).Column
        Os187Re187 = SearchRng.Find(what:="Os187_Re187", MatchCase:=False).Column
        Age = SearchRng.Find(what:="Cn277", MatchCase:=False).Column
        FRow = Mrow + 1
        LRow = Range("A" & Mrow).End(xlDown).Row
        
    'calculate values
        Range(Cells(FRow, Os187Corr), Cells(LRow, Os187Corr)).FormulaR1C1 = "=RC" & Os201 & "-(RC" & Re199 & "*1.6737)"
        Range(Cells(FRow, Re187Os187), Cells(LRow, Re187Os187)).FormulaR1C1 = "=IF(OR(RC" & Re185 & "=0,RC" & Os187Corr & "=0),0,(RC" & Re185 & "*1.6737)/RC" & Os187Corr & ")"
        Range(Cells(FRow, Os187Re187), Cells(LRow, Os187Re187)).FormulaR1C1 = "=IF(OR(RC" & Re185 & "=0,RC" & Os187Corr & "=0),0,RC" & Os187Corr & "/(RC" & Re185 & "*1.6737))"
        Range(Cells(FRow, Age), Cells(LRow, Age)).FormulaR1C1 = "=IF(RC" & Os187Re187 & "=0,0,((LN(RC" & Os187Re187 & "+1))/0.000016668))"
      
End With

End Sub

Private Sub BatchMassHdrCrReOs()
'This will batch process a folder's CSV files, based on the host file opened
    
    'Create FileSystemObject as FSO
        Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'Determine Path
        FolderName = Application.ActiveWorkbook.Path
            If Right(FolderName, 1) <> Application.PathSeparator Then FolderName = FolderName & Application.PathSeparator
        Set FSOFolderName = FSO.GetFolder(FolderName)
        Fname = Dir(FolderName & "*.csv")
        DestName = (FolderName & "Originals\")
    
    'Set sceen optimisations
        Application.ScreenUpdating = False
        Application.StatusBar = False
        Application.EnableEvents = False
    
    'Check for originals, if they don't exist, copy them.
        If Not FSO.FolderExists(DestName) Then
            MkDir (DestName)
            For Each FSOFileName In FSOFolderName.Files
            FSOFileName.Copy (DestName)
            Next
        End If
          
        Set FSOFolderName = Nothing
        Set FSOFileName = Nothing
    
    'loop through the files
        Do While Len(Fname)
            With Workbooks.Open(FolderName & Fname)
                Call ReOsQQQHeaderCorrEquations
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, FileFormat:=6, ConflictResolution:=xlLocalSessionChanges
                ActiveWorkbook.Close
                Application.DisplayAlerts = True
            End With
        'Adds minor delay to stabilise operation, and allows for excel to register events
            Sleep 1
            DoEvents
        ' Go to the next file in the folder
            Fname = Dir
        Loop
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = True
    MsgBox "Task completed, thanks for waiting"
    
End Sub
Sub MassShiftHdrReOsConfirm(control As IRibbonControl)
'This is to prevent accidentally running this procedure

    Msg = "This will batch correct mass shifted headers (e.g. QQQ data/CSVs only) and add equations for Re-Os geochronology." & vbCrLf & vbCrLf & "It will save a copy of the original CSV (into a new folder, Originals) if they do not already exist." & vbCrLf & vbCrLf & "It may take some time depending on the number of files and if copying of original files is needed. Please be patient." & vbCrLf & vbCrLf & "Do you want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "Mass Shift Header Correction Re-Os")
    Select Case Ans
        Case vbYes
        Call BatchMassHdrCrReOs
        Case vbNo
        GoTo Quit:
    End Select

Quit:
End Sub

