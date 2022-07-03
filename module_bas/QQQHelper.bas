Attribute VB_Name = "QQQHelper"
'Module for removing "shifted" values from elemental mass headers
Option Explicit
Option Compare Text
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

'Variable declaration
    Dim FolderName As String
    Dim FileName As String
    Dim DestName As String
    Dim FSO As Object
    Dim FSOFileName As Object
    Dim FSOFolderName As Object
    Dim MassRow As Long
    Dim LastColumn As Long
    Dim n As Long
    Dim Mass As String
    Dim Msg As String
    Dim Ans As Variant
    Dim DecaySystem As String
    Dim FileCount As Long
Private Sub LADRRbSrNormErrorCorrelationWorkaroundConfirm(control As IRibbonControl)
'This is to prevent accidentally ruining an excel spreadsheet
    Msg = "WARNING: There are inherent limitations on the stability of this program due to VBA/Excel memory management. I STRONGLY recommend you use the cross-platform PowerShell code available on GitHub: https://github.com/jarredclloyd?tab=repositories that is written for this operation." & vbCrLf & vbCrLf & "This procedure will create copy of LuHf CSV files and adjust the paret/daugther headers to U238, Pb207, and Pb206 so that LADR can directly calculate error correlations?" & vbCrLf & vbCrLf & "Have you saved the workbook, and do you really want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR Rb/Sr, Sr/Sr error correlation workaround")
    Select Case Ans
        Case vbYes
            DecaySystem = "RbSrNorm"
            Call LADRErrorCorrelationWorkaroundBatch
        Case vbNo
            Exit Sub
    End Select
End Sub
Private Sub LADRRbSrInvErrorCorrelationWorkaroundConfirm(control As IRibbonControl)
'This is to prevent accidentally ruining an excel spreadsheet
    Msg = "WARNING: There are inherent limitations on the stability of this program due to VBA/Excel memory management. I STRONGLY recommend you use the cross-platform PowerShell code available on GitHub: https://github.com/jarredclloyd?tab=repositories that is written for this operation." & vbCrLf & vbCrLf & "This procedure will create copy of LuHf CSV files and adjust the paret/daugther headers to U238, Pb207, and Pb206 so that LADR can directly calculate error correlations?" & vbCrLf & vbCrLf & "Have you saved the workbook, and do you really want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR Rb/Sr, Sr/Sr error correlation workaround")
    Select Case Ans
        Case vbYes
            DecaySystem = "RbSrInv"
            Call LADRErrorCorrelationWorkaroundBatch
        Case vbNo
            Exit Sub
    End Select
End Sub
Private Sub LADRLuHfNormErrorCorrelationWorkaroundConfirm(control As IRibbonControl)
'This is to prevent accidentally ruining an excel spreadsheet
    Msg = "WARNING: There are inherent limitations on the stability of this program due to VBA/Excel memory management. I STRONGLY recommend you use the cross-platform PowerShell code available on GitHub: https://github.com/jarredclloyd?tab=repositories that is written for this operation." & vbCrLf & vbCrLf & "This procedure will create copy of LuHf CSV files and adjust the paret/daugther headers to U238, Pb207, and Pb206 so that LADR can directly calculate error correlations?" & vbCrLf & vbCrLf & "Have you saved the workbook, and do you really want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR Lu/Hf, Hf/Hf error correlation workaround")
    Select Case Ans
        Case vbYes
            DecaySystem = "LuHfNorm"
            Call LADRErrorCorrelationWorkaroundBatch
        Case vbNo
            Exit Sub
    End Select
End Sub
Private Sub LADRLuHfInvErrorCorrelationWorkaroundConfirm(control As IRibbonControl)
'This is to prevent accidentally ruining an excel spreadsheet
    Msg = "WARNING: There are inherent limitations on the stability of this program due to VBA/Excel memory management. I STRONGLY recommend you use the cross-platform PowerShell code available on GitHub: https://github.com/jarredclloyd?tab=repositories that is written for this operation." & vbCrLf & vbCrLf & "This procedure will create copy of LuHf CSV files and adjust the paret/daugther headers to U238, Pb207, and Pb206 so that LADR can directly calculate error correlations?" & vbCrLf & vbCrLf & "Have you saved the workbook, and do you really want to continue?"
    Ans = MsgBox(Msg, vbYesNo, "LADR Lu/Hf, Hf/Hf error correlation workaround")
    Select Case Ans
        Case vbYes
            DecaySystem = "LuHfInv"
            Call LADRErrorCorrelationWorkaroundBatch
        Case vbNo
            Exit Sub
    End Select
End Sub
Private Sub LADRErrorCorrelationWorkaroundBatch()
'This will batch process a folder's CSV files, based on the host file opened
    'Create FileSystemObject as FSO
        Set FSO = CreateObject("Scripting.FileSystemObject")
    'Determine Path
        FolderName = Application.ActiveWorkbook.Path
            If Right(FolderName, 1) <> Application.PathSeparator Then FolderName = FolderName & Application.PathSeparator
        Set FSOFolderName = FSO.GetFolder(FolderName)
        DestName = (FolderName & "Originals\")
    'Set sceen optimisations
        Application.ScreenUpdating = False
        Application.StatusBar = False
        Application.EnableEvents = False
    'Count number of CSV files in folder and prevent code run on folders with >500 CSV files (limit of VBA/Excel stability)
        For Each FSOFileName In FSOFolderName.Files
            If FSOFileName Like "*.csv" Then
                FileCount = FileCount + 1
            End If
        Next FSOFileName
        If FileCount > 500 Then
            Application.EnableEvents = True
            Application.StatusBar = True
            Application.ScreenUpdating = True
            MsgBox "There are more than 500 CSV files in this folder. Please split your files into batches of ~500." & vbCrLf & vbCrLf & "I STRONGLY recommend you use the PowerShell script I have developed for this operation."
            Exit Sub
        Else
        End If
    'Check for original files folder, if it doesn't exist create folder and copy files.
        If Not FSO.FolderExists(DestName) Then
            MkDir (DestName)
            For Each FSOFileName In FSOFolderName.Files
                FSOFileName.Copy (DestName)
            Next
        End If
        With ActiveSheet
            'Determine mass header row
                MassRow = .Range("A:A").Find(what:="Time [Sec]", MatchCase:=True, lookat:=xlWhole).Row
            'Determine last column
                LastColumn = Cells(MassRow, Columns.Count).End(xlToLeft).Column
        End With
    'loop through the files
        For Each FSOFileName In FSOFolderName.Files
            Workbooks.Open (FSOFileName)
            Call LADRErrorCorrelationWorkaround
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs FileName:=Application.ActiveWorkbook.FullName, FileFormat:=6, ConflictResolution:=xlLocalSessionChanges
            Application.DisplayAlerts = True
            ActiveWorkbook.Close
        'Allows desktop to register events for stability of the procedure
            DoEvents
            Sleep (100)
        Next
    Set FSOFolderName = Nothing
    Set FSOFileName = Nothing
    Set FSO = Nothing
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    Close
    MsgBox "Task completed, thanks for waiting." & vbCrLf & "I recommend you create a folder named 'DecaySystem'_to_UPb (e.g., RbSr_to_UPb) and place the converted files in it."
End Sub
Private Sub LADRErrorCorrelationWorkaround()
    With ActiveSheet
        'Change shifted headers
            Select Case DecaySystem
                Case "RbSrNorm"
                    For n = 2 To LastColumn
                        If Cells(MassRow, n).Value = "Rb85 -> 85" Then
                            Cells(MassRow, n).Value = "U238"
                        ElseIf Cells(MassRow, n).Value = "Sr87 -> 103" Then
                            Cells(MassRow, n).Value = "Pb207"
                        ElseIf Cells(MassRow, n).Value = "Sr86 -> 102" Then
                            Cells(MassRow, n).Value = "Pb206"
                        ElseIf Cells(MassRow, n).Value = "U238 -> 270" Then
                            Cells(MassRow, n).Value = "U235"
                        Else
                        End If
                    Next n
                Case "RbSrInv"
                    For n = 2 To LastColumn
                        If Cells(MassRow, n).Value = "Rb85 -> 85" Then
                            Cells(MassRow, n).Value = "U238"
                        ElseIf Cells(MassRow, n).Value = "Sr87 -> 103" Then
                            Cells(MassRow, n).Value = "Pb206"
                        ElseIf Cells(MassRow, n).Value = "Sr86 -> 102" Then
                            Cells(MassRow, n).Value = "Pb207"
                        ElseIf Cells(MassRow, n).Value = "U238 -> 270" Then
                            Cells(MassRow, n).Value = "U235"
                        Else
                        End If
                    Next n
                Case "LuHfNorm"
                    For n = 2 To LastColumn
                        If Cells(MassRow, n).Value = "Lu175 -> 175" Then
                            Cells(MassRow, n).Value = "U238"
                        ElseIf Cells(MassRow, n).Value = "Hf176 -> 258" Then
                            Cells(MassRow, n).Value = "Pb207"
                        ElseIf Cells(MassRow, n).Value = "Hf178 -> 260" Then
                            Cells(MassRow, n).Value = "Pb206"
                        ElseIf Cells(MassRow, n).Value = "U238 -> 270" Then
                            Cells(MassRow, n).Value = "U235"
                        Else
                        End If
                    Next n
                Case "LuHfInv"
                    For n = 2 To LastColumn
                        If Cells(MassRow, n).Value = "Lu175 -> 175" Then
                            Cells(MassRow, n).Value = "U238"
                        ElseIf Cells(MassRow, n).Value = "Hf176 -> 258" Then
                            Cells(MassRow, n).Value = "Pb206"
                        ElseIf Cells(MassRow, n).Value = "Hf178 -> 260" Then
                            Cells(MassRow, n).Value = "Pb207"
                        ElseIf Cells(MassRow, n).Value = "U238 -> 270" Then
                            Cells(MassRow, n).Value = "U235"
                        Else
                        End If
                    Next n
                'Case "SmNd"
                    'For n = 2 To LastColumn
                        'If Cells(MassRow, n).Value = "Rb85 -> 85" Then
                            'Cells(MassRow, n).Value = "U238"
                        'ElseIf Cells(MassRow, n).Value = "Sr87 -> 103" Then
                            'Cells(MassRow, n).Value = "Pb207"
                        'ElseIf Cells(MassRow, n).Value = "Sr86 -> 102" Then
                            'Cells(MassRow, n).Value = "Pb206"
                        'ElseIf Cells(MassRow, n).Value = "U238 -> 270" Then
                            'Cells(MassRow, n).Value = "U235"
                        'Else
                        'End If
                    'Next n
                'Case "ReOs"
                    'For n = 2 To LastColumn
                        'If Cells(MassRow, n).Value = "Rb85 -> 85" Then
                            'Cells(MassRow, n).Value = "U238"
                        'ElseIf Cells(MassRow, n).Value = "Sr87 -> 103" Then
                            'Cells(MassRow, n).Value = "Pb207"
                        'ElseIf Cells(MassRow, n).Value = "Sr86 -> 102" Then
                            'Cells(MassRow, n).Value = "Pb206"
                        'ElseIf Cells(MassRow, n).Value = "U238 -> 270" Then
                            'Cells(MassRow, n).Value = "U235"
                        'Else
                        'End If
                    'Next n
            End Select
    End With
End Sub
