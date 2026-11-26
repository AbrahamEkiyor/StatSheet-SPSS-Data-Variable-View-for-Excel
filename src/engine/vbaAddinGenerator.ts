export function generateStatSheetVbaModule(): string {
  return `Attribute VB_Name = "StatSheet_SPSS"
Option Explicit

' ============================================================================
' StatSheet SPSS Data & Variable View — Offline Excel Add-in (.xlam)
' Zero-server, 100% offline VBA implementation for Windows & macOS Excel
' ============================================================================

Private Const VAR_SHEET As String = "Variable_View"
Private Const BACKUP_SHEET As String = "_SPSS_RawBackup"
Private Const STATE_SHEET As String = "_SPSS_State"
Private Const OUTPUT_SHEET As String = "_SPSS_Output"

' Keyboard shortcuts registered on Add-in load
Public Sub Auto_Open()
    On Error Resume Next
    Application.OnKey "^+L", "ToggleValueLabels"
    Application.OnKey "^+V", "SwitchDataAndVariableView"
    CreateLegacyMenu
    CreateCellContextMenu
    On Error GoTo 0
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Application.OnKey "^+L"
    Application.OnKey "^+V"
    RemoveLegacyMenu
    RemoveCellContextMenu
    On Error GoTo 0
End Sub

' Optional RibbonX callbacks (used if customUI14.xml is attached)
Public Sub Ribbon_ToggleValueLabels(ByVal control As Object)
    ToggleValueLabels
End Sub

Public Sub Ribbon_SwitchView(ByVal control As Object)
    SwitchDataAndVariableView
End Sub

Public Sub Ribbon_SyncVars(ByVal control As Object)
    InitOrSyncVariableView
End Sub

Public Sub Ribbon_EditLabels(ByVal control As Object)
    EditValueLabelsForSelection
End Sub

Public Sub Ribbon_Validate(ByVal control As Object)
    ValidateAndFlagMissing
End Sub

Public Sub Ribbon_Descriptives(ByVal control As Object)
    GenerateDescriptivesForSelection
End Sub

Public Sub Ribbon_ExportJson(ByVal control As Object)
    ExportCodebookJson
End Sub

Private Sub CreateLegacyMenu()
    Dim cb As CommandBar
    Dim menuPopup As CommandBarPopup
    Dim btn As CommandBarButton

    On Error Resume Next
    Set cb = Application.CommandBars("Worksheet Menu Bar")
    If cb Is Nothing Then Exit Sub

    RemoveLegacyMenu
    Set menuPopup = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    menuPopup.Caption = "StatSheet &SPSS"
    menuPopup.Tag = "StatSheetSPSSMenu"

    Set btn = menuPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Toggle Value Labels (1 <-> A)  [Ctrl+Shift+L]"
    btn.OnAction = "ToggleValueLabels"

    Set btn = menuPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Switch Data / Variable View  [Ctrl+Shift+V]"
    btn.OnAction = "SwitchDataAndVariableView"

    Set btn = menuPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Sync / Initialize Variable View"
    btn.OnAction = "InitOrSyncVariableView"
    btn.BeginGroup = True

    Set btn = menuPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Edit Value Labels for Selection..."
    btn.OnAction = "EditValueLabelsForSelection"

    Set btn = menuPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Validate Rules & Flag Missing Values"
    btn.OnAction = "ValidateAndFlagMissing"
    btn.BeginGroup = True

    Set btn = menuPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Clear Validation Highlights"
    btn.OnAction = "ClearValidationHighlights"

    Set btn = menuPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Descriptives & Frequencies Report"
    btn.OnAction = "GenerateDescriptivesForSelection"
    btn.BeginGroup = True

    Set btn = menuPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Export Codebook (.json)"
    btn.OnAction = "ExportCodebookJson"
    On Error GoTo 0
End Sub

Private Sub RemoveLegacyMenu()
    Dim cb As CommandBar
    Dim ctrl As CommandBarControl
    On Error Resume Next
    Set cb = Application.CommandBars("Worksheet Menu Bar")
    If cb Is Nothing Then Exit Sub
    For Each ctrl In cb.Controls
        If ctrl.Tag = "StatSheetSPSSMenu" Then ctrl.Delete
    Next ctrl
    On Error GoTo 0
End Sub

Private Sub CreateCellContextMenu()
    Dim cb As CommandBar
    Dim subMenu As CommandBarPopup
    Dim btn As CommandBarButton

    On Error Resume Next
    Set cb = Application.CommandBars("Cell")
    If cb Is Nothing Then Exit Sub

    RemoveCellContextMenu
    Set subMenu = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    subMenu.Caption = "StatSheet SPSS"
    subMenu.Tag = "StatSheetCellMenu"
    subMenu.BeginGroup = True

    Set btn = subMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Toggle Value Labels (1 <-> A)"
    btn.OnAction = "ToggleValueLabels"

    Set btn = subMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Switch Data / Variable View"
    btn.OnAction = "SwitchDataAndVariableView"

    Set btn = subMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Sync / Initialize Variable View"
    btn.OnAction = "InitOrSyncVariableView"

    Set btn = subMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Edit Value Labels..."
    btn.OnAction = "EditValueLabelsForSelection"

    Set btn = subMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Validate Rules & Flag Missing"
    btn.OnAction = "ValidateAndFlagMissing"

    Set btn = subMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "Descriptives & Frequencies"
    btn.OnAction = "GenerateDescriptivesForSelection"
    On Error GoTo 0
End Sub

Private Sub RemoveCellContextMenu()
    Dim cb As CommandBar
    Dim ctrl As CommandBarControl
    On Error Resume Next
    Set cb = Application.CommandBars("Cell")
    If cb Is Nothing Then Exit Sub
    For Each ctrl In cb.Controls
        If ctrl.Tag = "StatSheetCellMenu" Then ctrl.Delete
    Next ctrl
    On Error GoTo 0
End Sub

' ============================================================================
' 1. Initialize or Sync 11-Column SPSS Variable View
' ============================================================================
Public Sub InitOrSyncVariableView()
    Dim wb As Workbook
    Dim dataWs As Worksheet
    Dim varWs As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim varName As String
    Dim nextRow As Long
    Dim existingRow As Long

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, "StatSheet SPSS"
        Exit Sub
    End If

    Set dataWs = ResolveDataWorksheet(wb)
    If dataWs Is Nothing Then
        MsgBox "Select a worksheet containing variable headers in Row 1.", vbExclamation, "StatSheet SPSS"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    Set varWs = EnsureVariableViewSheet(wb)
    SaveState wb, GetDisplayMode(wb), dataWs.Name

    lastCol = dataWs.Cells(1, dataWs.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then GoTo Cleanup

    For col = 1 To lastCol
        varName = Trim$(CStr(dataWs.Cells(1, col).Value))
        If varName <> "" Then
            existingRow = FindVariableRow(varWs, varName)
            If existingRow = 0 Then
                nextRow = varWs.Cells(varWs.Rows.Count, 1).End(xlUp).Row + 1
                If nextRow < 2 Then nextRow = 2

                Dim inferredType As String
                Dim inferredMeasure As String
                Dim inferredAlign As String
                InferColumnMetadata dataWs, col, inferredType, inferredMeasure, inferredAlign

                varWs.Cells(nextRow, 1).Value = varName
                varWs.Cells(nextRow, 2).Value = inferredType
                varWs.Cells(nextRow, 3).Value = IIf(inferredType = "String", 24, 8)
                varWs.Cells(nextRow, 4).Value = IIf(inferredType = "String", 0, 2)
                varWs.Cells(nextRow, 5).Value = ""
                varWs.Cells(nextRow, 6).Value = ""
                varWs.Cells(nextRow, 7).Value = ""
                varWs.Cells(nextRow, 8).Value = 14
                varWs.Cells(nextRow, 9).Value = inferredAlign
                varWs.Cells(nextRow, 10).Value = inferredMeasure
                varWs.Cells(nextRow, 11).Value = "Input"
                varWs.Cells(nextRow, 12).Value = ""
                varWs.Cells(nextRow, 13).Value = ""
                varWs.Cells(nextRow, 14).Value = "No"
            End If
        End If
    Next col

    ApplyVariableViewDropdowns varWs
    varWs.Columns("A:N").AutoFit
    varWs.Columns(5).ColumnWidth = 28
    varWs.Columns(6).ColumnWidth = 34

    MsgBox "Variable_View synchronized with '" & dataWs.Name & "' (" & lastCol & " columns).", vbInformation, "StatSheet SPSS"

Cleanup:
    Application.ScreenUpdating = True
End Sub

' ============================================================================
' 2. Switch Between Data View & Variable View (Ctrl+Shift+V)
' ============================================================================
Public Sub SwitchDataAndVariableView()
    Dim wb As Workbook
    Dim dataSheetName As String
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    If ActiveSheet.Name = VAR_SHEET Then
        dataSheetName = GetStoredDataSheetName(wb)
        If dataSheetName <> "" And SheetExists(wb, dataSheetName) Then
            wb.Worksheets(dataSheetName).Activate
        Else
            Dim fallbackWs As Worksheet
            Set fallbackWs = ResolveDataWorksheet(wb)
            If Not fallbackWs Is Nothing Then fallbackWs.Activate
        End If
    Else
        If ActiveSheet.Name <> BACKUP_SHEET And ActiveSheet.Name <> STATE_SHEET And ActiveSheet.Name <> OUTPUT_SHEET Then
            SaveState wb, GetDisplayMode(wb), ActiveSheet.Name
        End If
        If Not SheetExists(wb, VAR_SHEET) Then
            InitOrSyncVariableView
        End If
        wb.Worksheets(VAR_SHEET).Activate
    End If
End Sub

' ============================================================================
' 3. Reversible Value Labels Toggle (1 <-> A) [Ctrl+Shift+L]
' ============================================================================
Public Sub ToggleValueLabels()
    Dim wb As Workbook
    Dim dataWs As Worksheet
    Dim varWs As Worksheet
    Dim backupWs As Worksheet
    Dim currentMode As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim r As Long
    Dim varName As String
    Dim varRow As Long
    Dim labelsStr As String
    Dim codes() As String
    Dim lbls() As String
    Dim pairCount As Long
    Dim swappedCount As Long

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    If Not SheetExists(wb, VAR_SHEET) Then
        InitOrSyncVariableView
    End If
    Set varWs = wb.Worksheets(VAR_SHEET)
    Set dataWs = ResolveDataWorksheet(wb)
    If dataWs Is Nothing Then
        MsgBox "No active Data worksheet found.", vbExclamation, "StatSheet SPSS"
        Exit Sub
    End If

    currentMode = GetDisplayMode(wb)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row
    lastCol = dataWs.Cells(1, dataWs.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 1 Then
        MsgBox "Data worksheet needs headers in Row 1 and at least one data row.", vbInformation, "StatSheet SPSS"
        GoTo Cleanup
    End If

    Set backupWs = EnsureHiddenSheet(wb, BACKUP_SHEET)

    If currentMode <> "LABELS" Then
        ' ---- RAW -> LABELS ----
        backupWs.Cells.Clear
        dataWs.Range(dataWs.Cells(1, 1), dataWs.Cells(lastRow, lastCol)).Copy
        backupWs.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False

        swappedCount = 0
        For col = 1 To lastCol
            varName = Trim$(CStr(dataWs.Cells(1, col).Value))
            varRow = FindVariableRow(varWs, varName)
            If varRow > 0 Then
                labelsStr = Trim$(CStr(varWs.Cells(varRow, 6).Value))
                If ParseValueLabels(labelsStr, codes, lbls, pairCount) Then
                    For r = 2 To lastRow
                        Dim rawVal As String
                        rawVal = Trim$(CStr(dataWs.Cells(r, col).Value))
                        If rawVal <> "" Then
                            Dim mappedLbl As String
                            mappedLbl = LookupLabelByCode(rawVal, codes, lbls, pairCount)
                            If mappedLbl <> "" Then
                                dataWs.Cells(r, col).Value = mappedLbl
                                swappedCount = swappedCount + 1
                            End If
                        End If
                    Next r

                    If UCase$(Trim$(CStr(varWs.Cells(varRow, 14).Value))) = "YES" Then
                        ApplyColumnValidationList dataWs.Range(dataWs.Cells(2, col), dataWs.Cells(lastRow, col)), lbls, pairCount
                    End If
                End If
            End If
        Next col

        SaveState wb, "LABELS", dataWs.Name
        Application.StatusBar = "StatSheet: Value Labels ON (" & swappedCount & " cells displayed as labels; raw codes backed up)"
        MsgBox "Value Labels: ON" & vbCrLf & swappedCount & " cells switched to labels. Original raw codes preserved in hidden backup.", vbInformation, "StatSheet SPSS"
    Else
        ' ---- LABELS -> RAW (Bidirectional safe restore) ----
        swappedCount = 0
        For col = 1 To lastCol
            varName = Trim$(CStr(dataWs.Cells(1, col).Value))
            varRow = FindVariableRow(varWs, varName)
            If varRow > 0 Then
                labelsStr = Trim$(CStr(varWs.Cells(varRow, 6).Value))
                If ParseValueLabels(labelsStr, codes, lbls, pairCount) Then
                    For r = 2 To lastRow
                        Dim dispVal As String
                        dispVal = Trim$(CStr(dataWs.Cells(r, col).Value))
                        If dispVal <> "" Then
                            Dim mappedCode As String
                            mappedCode = LookupCodeByLabel(dispVal, codes, lbls, pairCount)
                            If mappedCode <> "" Then
                                If IsNumeric(mappedCode) Then
                                    dataWs.Cells(r, col).Value = CDbl(mappedCode)
                                Else
                                    dataWs.Cells(r, col).Value = mappedCode
                                End If
                                swappedCount = swappedCount + 1
                            ElseIf Not IsEmpty(backupWs.Cells(r, col).Value) Then
                                dataWs.Cells(r, col).Value = backupWs.Cells(r, col).Value
                            End If
                        End If
                    Next r

                    If UCase$(Trim$(CStr(varWs.Cells(varRow, 14).Value))) = "YES" Then
                        ApplyColumnValidationList dataWs.Range(dataWs.Cells(2, col), dataWs.Cells(lastRow, col)), codes, pairCount
                    End If
                End If
            End If
        Next col

        SaveState wb, "RAW", dataWs.Name
        Application.StatusBar = "StatSheet: Value Labels OFF (Raw codes restored)"
        MsgBox "Value Labels: OFF" & vbCrLf & swappedCount & " cells restored to raw numeric/text codes.", vbInformation, "StatSheet SPSS"
    End If

    GoTo Cleanup

ErrHandler:
    MsgBox "Error toggling Value Labels: " & Err.Description, vbCritical, "StatSheet SPSS"

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ============================================================================
' 4. Quick Value Labels Editor for Active Column / Variable
' ============================================================================
Public Sub EditValueLabelsForSelection()
    Dim wb As Workbook
    Dim varWs As Worksheet
    Dim dataWs As Worksheet
    Dim varName As String
    Dim varRow As Long
    Dim currentVal As String
    Dim nextVal As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Not SheetExists(wb, VAR_SHEET) Then InitOrSyncVariableView
    Set varWs = wb.Worksheets(VAR_SHEET)

    If ActiveSheet.Name = VAR_SHEET Then
        varRow = ActiveCell.Row
        If varRow < 2 Then
            MsgBox "Select a variable row in Variable_View.", vbInformation, "StatSheet SPSS"
            Exit Sub
        End If
        varName = Trim$(CStr(varWs.Cells(varRow, 1).Value))
    Else
        Set dataWs = ResolveDataWorksheet(wb)
        If dataWs Is Nothing Then Exit Sub
        varName = Trim$(CStr(dataWs.Cells(1, ActiveCell.Column).Value))
        varRow = FindVariableRow(varWs, varName)
    End If

    If varRow = 0 Or varName = "" Then
        MsgBox "Select a valid variable column or Variable_View row.", vbExclamation, "StatSheet SPSS"
        Exit Sub
    End If

    currentVal = CStr(varWs.Cells(varRow, 6).Value)
    nextVal = InputBox( _
        "Define Value Labels for '" & varName & "'" & vbCrLf & vbCrLf & _
        "Format: 1=Low; 2=Medium; 3=High" & vbCrLf & _
        "Or type LIKERT5 or YESNO for quick presets:", _
        "StatSheet Value Labels Editor", _
        currentVal)

    If StrPtr(nextVal) = 0 Then Exit Sub

    nextVal = Trim$(nextVal)
    If UCase$(nextVal) = "LIKERT5" Then
        nextVal = "1=Strongly Disagree; 2=Disagree; 3=Neutral; 4=Agree; 5=Strongly Agree; -9=Skipped"
    ElseIf UCase$(nextVal) = "YESNO" Then
        nextVal = "0=No; 1=Yes; 8=Unsure"
    End If

    varWs.Cells(varRow, 6).Value = nextVal
    MsgBox "Updated Value Labels for '" & varName & "'.", vbInformation, "StatSheet SPSS"
End Sub

' ============================================================================
' 5. Missing Values & Validation Scanner
' ============================================================================
Public Sub ValidateAndFlagMissing()
    Dim wb As Workbook
    Dim dataWs As Worksheet
    Dim varWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim col As Long, r As Long
    Dim varName As String, varRow As Long
    Dim missingSpec As String, minSpec As String, maxSpec As String, restrictSpec As String
    Dim labelsStr As String
    Dim codes() As String, lbls() As String, pairCount As Long
    Dim missingCount As Long, invalidCount As Long
    Dim isLabelsMode As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Not SheetExists(wb, VAR_SHEET) Then InitOrSyncVariableView
    Set varWs = wb.Worksheets(VAR_SHEET)
    Set dataWs = ResolveDataWorksheet(wb)
    If dataWs Is Nothing Then Exit Sub

    isLabelsMode = (GetDisplayMode(wb) = "LABELS")
    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row
    lastCol = dataWs.Cells(1, dataWs.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub

    Application.ScreenUpdating = False
    ClearHighlightsOnSheet dataWs, lastRow, lastCol

    For col = 1 To lastCol
        varName = Trim$(CStr(dataWs.Cells(1, col).Value))
        varRow = FindVariableRow(varWs, varName)
        If varRow > 0 Then
            labelsStr = Trim$(CStr(varWs.Cells(varRow, 6).Value))
            ParseValueLabels labelsStr, codes, lbls, pairCount
            missingSpec = Trim$(CStr(varWs.Cells(varRow, 7).Value))
            minSpec = Trim$(CStr(varWs.Cells(varRow, 12).Value))
            maxSpec = Trim$(CStr(varWs.Cells(varRow, 13).Value))
            restrictSpec = UCase$(Trim$(CStr(varWs.Cells(varRow, 14).Value)))

            For r = 2 To lastRow
                Dim cell As Range
                Set cell = dataWs.Cells(r, col)
                Dim rawVal As String
                rawVal = Trim$(CStr(cell.Value))

                If rawVal <> "" Then
                    If isLabelsMode And pairCount > 0 Then
                        Dim revCode As String
                        revCode = LookupCodeByLabel(rawVal, codes, lbls, pairCount)
                        If revCode <> "" Then rawVal = revCode
                    End If

                    If MatchesMissingSpec(rawVal, missingSpec) Then
                        cell.Interior.Color = RGB(254, 243, 199)
                        AddOrUpdateNote cell, "StatSheet Missing Value: Code '" & rawVal & "' excluded from valid stats."
                        missingCount = missingCount + 1
                    Else
                        Dim errReason As String
                        errReason = CheckCellValidation(rawVal, minSpec, maxSpec, restrictSpec = "YES", codes, pairCount)
                        If errReason <> "" Then
                            cell.Interior.Color = RGB(254, 226, 226)
                            AddOrUpdateNote cell, "StatSheet Validation Issue: " & errReason
                            invalidCount = invalidCount + 1
                        End If
                    End If
                End If
            Next r
        End If
    Next col

    Application.ScreenUpdating = True
    MsgBox "Validation scan complete on '" & dataWs.Name & "':" & vbCrLf & _
           "- User-missing cells flagged (amber): " & missingCount & vbCrLf & _
           "- Validation errors flagged (red): " & invalidCount, vbInformation, "StatSheet SPSS"
End Sub

Public Sub ClearValidationHighlights()
    Dim wb As Workbook
    Dim dataWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set dataWs = ResolveDataWorksheet(wb)
    If dataWs Is Nothing Then Exit Sub
    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row
    lastCol = dataWs.Cells(1, dataWs.Columns.Count).End(xlToLeft).Column
    ClearHighlightsOnSheet dataWs, lastRow, lastCol
End Sub

' ============================================================================
' 6. Quick Descriptives & Frequencies Output Sheet
' ============================================================================
Public Sub GenerateDescriptivesForSelection()
    Dim wb As Workbook
    Dim dataWs As Worksheet
    Dim varWs As Worksheet
    Dim outWs As Worksheet
    Dim col As Long, lastRow As Long, r As Long
    Dim varName As String, varLabel As String, missingSpec As String, labelsStr As String
    Dim varRow As Long
    Dim codes() As String, lbls() As String, pairCount As Long
    Dim isLabelsMode As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Not SheetExists(wb, VAR_SHEET) Then InitOrSyncVariableView
    Set varWs = wb.Worksheets(VAR_SHEET)
    Set dataWs = ResolveDataWorksheet(wb)
    If dataWs Is Nothing Then Exit Sub

    If ActiveSheet.Name = VAR_SHEET Then
        If ActiveCell.Row < 2 Then Exit Sub
        varName = Trim$(CStr(varWs.Cells(ActiveCell.Row, 1).Value))
        col = FindDataColumn(dataWs, varName)
    Else
        col = ActiveCell.Column
        varName = Trim$(CStr(dataWs.Cells(1, col).Value))
    End If

    If col = 0 Or varName = "" Then
        MsgBox "Select a variable column to analyze.", vbExclamation, "StatSheet SPSS"
        Exit Sub
    End If

    varRow = FindVariableRow(varWs, varName)
    If varRow > 0 Then
        varLabel = CStr(varWs.Cells(varRow, 5).Value)
        labelsStr = CStr(varWs.Cells(varRow, 6).Value)
        missingSpec = CStr(varWs.Cells(varRow, 7).Value)
    End If
    ParseValueLabels labelsStr, codes, lbls, pairCount
    isLabelsMode = (GetDisplayMode(wb) = "LABELS")

    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim totalN As Long, validN As Long, userMissingN As Long, sysMissingN As Long
    Dim numVals() As Double, numCount As Long
    Dim freqKeys() As String, freqCounts() As Long, freqMissing() As Boolean, freqTotal As Long
    ReDim numVals(1 To lastRow - 1)
    ReDim freqKeys(1 To lastRow - 1)
    ReDim freqCounts(1 To lastRow - 1)
    ReDim freqMissing(1 To lastRow - 1)

    For r = 2 To lastRow
        totalN = totalN + 1
        Dim rawVal As String
        rawVal = Trim$(CStr(dataWs.Cells(r, col).Value))
        If rawVal = "" Then
            sysMissingN = sysMissingN + 1
        Else
            If isLabelsMode And pairCount > 0 Then
                Dim rev As String
                rev = LookupCodeByLabel(rawVal, codes, lbls, pairCount)
                If rev <> "" Then rawVal = rev
            End If

            Dim isMiss As Boolean
            isMiss = MatchesMissingSpec(rawVal, missingSpec)
            If isMiss Then
                userMissingN = userMissingN + 1
            Else
                validN = validN + 1
                If IsNumeric(rawVal) Then
                    numCount = numCount + 1
                    numVals(numCount) = CDbl(rawVal)
                End If
            End If

            Dim idx As Long, foundIdx As Long
            foundIdx = 0
            For idx = 1 To freqTotal
                If freqKeys(idx) = rawVal Then
                    foundIdx = idx
                    Exit For
                End If
            Next idx
            If foundIdx > 0 Then
                freqCounts(foundIdx) = freqCounts(foundIdx) + 1
            Else
                freqTotal = freqTotal + 1
                freqKeys(freqTotal) = rawVal
                freqCounts(freqTotal) = 1
                freqMissing(freqTotal) = isMiss
            End If
        End If
    Next r

    Set outWs = EnsureOutputSheet(wb)
    outWs.Cells.Clear

    outWs.Range("A1").Value = "StatSheet Descriptive & Frequency Report"
    outWs.Range("A1").Font.Bold = True
    outWs.Range("A1").Font.Size = 13
    outWs.Range("A2").Value = "Variable: " & varName & IIf(varLabel <> "", " — " & varLabel, "")

    outWs.Range("A4:D4").Value = Array("Total N", "Valid N", "User Missing", "System Missing")
    outWs.Range("A4:D4").Font.Bold = True
    outWs.Range("A5:D5").Value = Array(totalN, validN, userMissingN, sysMissingN)

    If numCount > 0 Then
        ReDim Preserve numVals(1 To numCount)
        Dim meanVal As Double, sdVal As Double, minVal As Double, maxVal As Double, medVal As Double
        meanVal = Application.WorksheetFunction.Average(numVals)
        minVal = Application.WorksheetFunction.Min(numVals)
        maxVal = Application.WorksheetFunction.Max(numVals)
        medVal = Application.WorksheetFunction.Median(numVals)
        If numCount > 1 Then
            sdVal = Application.WorksheetFunction.StDev_S(numVals)
        Else
            sdVal = 0
        End If

        outWs.Range("A7:E7").Value = Array("Mean", "Median", "Std. Dev.", "Min", "Max")
        outWs.Range("A7:E7").Font.Bold = True
        outWs.Range("A8:E8").Value = Array(Round(meanVal, 3), Round(medVal, 3), Round(sdVal, 3), minVal, maxVal)
    End If

    outWs.Range("A10:F10").Value = Array("Code", "Value Label", "Status", "Frequency", "Percent", "Valid %")
    outWs.Range("A10:F10").Font.Bold = True

    Dim outRow As Long
    outRow = 11
    Dim i As Long
    For i = 1 To freqTotal
        Dim lbl As String
        lbl = LookupLabelByCode(freqKeys(i), codes, lbls, pairCount)
        If lbl = "" Then lbl = freqKeys(i)
        outWs.Cells(outRow, 1).Value = freqKeys(i)
        outWs.Cells(outRow, 2).Value = lbl
        outWs.Cells(outRow, 3).Value = IIf(freqMissing(i), "Missing", "Valid")
        outWs.Cells(outRow, 4).Value = freqCounts(i)
        outWs.Cells(outRow, 5).Value = IIf(totalN > 0, Round((freqCounts(i) / totalN) * 100, 1), 0)
        If Not freqMissing(i) And validN > 0 Then
            outWs.Cells(outRow, 6).Value = Round((freqCounts(i) / validN) * 100, 1)
        Else
            outWs.Cells(outRow, 6).Value = "—"
        End If
        outRow = outRow + 1
    Next i

    outWs.Columns("A:F").AutoFit
    outWs.Activate
End Sub

' ============================================================================
' 7. Portable JSON Codebook Export (Compatible with StatSheet Web Simulator)
' ============================================================================
Public Sub ExportCodebookJson()
    Dim wb As Workbook
    Dim varWs As Worksheet
    Dim lastRow As Long, r As Long
    Dim jsonPath As String
    Dim fNum As Integer
    Dim sb As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Not SheetExists(wb, VAR_SHEET) Then InitOrSyncVariableView
    Set varWs = wb.Worksheets(VAR_SHEET)

    If wb.Path = "" Then
        MsgBox "Save your Excel workbook first so the Codebook JSON can be exported alongside it.", vbExclamation, "StatSheet SPSS"
        Exit Sub
    End If

    jsonPath = wb.Path & Application.PathSeparator & "statsheet_codebook.json"
    lastRow = varWs.Cells(varWs.Rows.Count, 1).End(xlUp).Row

    sb = "{" & vbCrLf & _
         "  ""format"": ""statsheet-codebook-v1""," & vbCrLf & _
         "  ""datasetName"": """ & JsonEscape(wb.Name) & """," & vbCrLf & _
         "  ""variables"": [" & vbCrLf

    For r = 2 To lastRow
        Dim vName As String
        vName = Trim$(CStr(varWs.Cells(r, 1).Value))
        If vName <> "" Then
            If r > 2 Then sb = sb & "," & vbCrLf
            sb = sb & "    {" & vbCrLf & _
                 "      ""id"": ""v_" & r & """," & vbCrLf & _
                 "      ""name"": """ & JsonEscape(vName) & """," & vbCrLf & _
                 "      ""type"": """ & LCase$(JsonEscape(CStr(varWs.Cells(r, 2).Value))) & """," & vbCrLf & _
                 "      ""width"": " & Val(CStr(varWs.Cells(r, 3).Value)) & "," & vbCrLf & _
                 "      ""decimals"": " & Val(CStr(varWs.Cells(r, 4).Value)) & "," & vbCrLf & _
                 "      ""label"": """ & JsonEscape(CStr(varWs.Cells(r, 5).Value)) & """," & vbCrLf & _
                 "      ""valueLabels"": " & ValueLabelsToJson(CStr(varWs.Cells(r, 6).Value)) & "," & vbCrLf & _
                 "      ""missingValues"": { ""mode"": ""none"", ""discrete"": [] }," & vbCrLf & _
                 "      ""columns"": 128," & vbCrLf & _
                 "      ""align"": """ & LCase$(JsonEscape(CStr(varWs.Cells(r, 9).Value))) & """," & vbCrLf & _
                 "      ""measure"": """ & LCase$(JsonEscape(CStr(varWs.Cells(r, 10).Value))) & """," & vbCrLf & _
                 "      ""role"": """ & LCase$(JsonEscape(CStr(varWs.Cells(r, 11).Value))) & """" & vbCrLf & _
                 "    }"
        End If
    Next r

    sb = sb & vbCrLf & "  ]" & vbCrLf & "}"

    fNum = FreeFile
    Open jsonPath For Output As #fNum
    Print #fNum, sb
    Close #fNum

    MsgBox "Exported JSON Codebook to:" & vbCrLf & jsonPath, vbInformation, "StatSheet SPSS"
End Sub

' ============================================================================
' Helper Routines (100% Pure VBA — Compatible with Windows & macOS Excel)
' ============================================================================
Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Private Function EnsureVariableViewSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    If SheetExists(wb, VAR_SHEET) Then
        Set ws = wb.Worksheets(VAR_SHEET)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = VAR_SHEET
        ws.Range("A1:N1").Value = Array( _
            "Name", "Type", "Width", "Decimals", "Label", "Values", _
            "Missing", "Columns", "Align", "Measure", "Role", _
            "Val_Min", "Val_Max", "Restrict_Labels")
        ws.Range("A1:N1").Font.Bold = True
        ws.Range("A1:N1").Interior.Color = RGB(241, 245, 249)
    End If
    Set EnsureVariableViewSheet = ws
End Function

Private Function EnsureHiddenSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    If SheetExists(wb, sheetName) Then
        Set ws = wb.Worksheets(sheetName)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Visible = xlSheetVeryHidden
    Set EnsureHiddenSheet = ws
End Function

Private Function EnsureOutputSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    If SheetExists(wb, OUTPUT_SHEET) Then
        Set ws = wb.Worksheets(OUTPUT_SHEET)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = OUTPUT_SHEET
    End If
    Set EnsureOutputSheet = ws
End Function

Private Function ResolveDataWorksheet(wb As Workbook) As Worksheet
    Dim storedName As String
    Dim ws As Worksheet
    If ActiveSheet.Name <> VAR_SHEET And ActiveSheet.Name <> BACKUP_SHEET And ActiveSheet.Name <> STATE_SHEET And ActiveSheet.Name <> OUTPUT_SHEET Then
        Set ResolveDataWorksheet = ActiveSheet
        Exit Function
    End If
    storedName = GetStoredDataSheetName(wb)
    If storedName <> "" And SheetExists(wb, storedName) Then
        Set ResolveDataWorksheet = wb.Worksheets(storedName)
        Exit Function
    End If
    For Each ws In wb.Worksheets
        If ws.Name <> VAR_SHEET And ws.Name <> BACKUP_SHEET And ws.Name <> STATE_SHEET And ws.Name <> OUTPUT_SHEET Then
            Set ResolveDataWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Sub SaveState(wb As Workbook, modeStr As String, dataSheetName As String)
    Dim st As Worksheet
    Set st = EnsureHiddenSheet(wb, STATE_SHEET)
    st.Range("A1").Value = modeStr
    st.Range("A2").Value = dataSheetName
    st.Range("A3").Value = Now
End Sub

Private Function GetDisplayMode(wb As Workbook) As String
    If Not SheetExists(wb, STATE_SHEET) Then
        GetDisplayMode = "RAW"
    Else
        GetDisplayMode = UCase$(Trim$(CStr(wb.Worksheets(STATE_SHEET).Range("A1").Value)))
        If GetDisplayMode = "" Then GetDisplayMode = "RAW"
    End If
End Function

Private Function GetStoredDataSheetName(wb As Workbook) As String
    If SheetExists(wb, STATE_SHEET) Then
        GetStoredDataSheetName = Trim$(CStr(wb.Worksheets(STATE_SHEET).Range("A2").Value))
    End If
End Function

Private Function FindVariableRow(varWs As Worksheet, varName As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = varWs.Cells(varWs.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(varWs.Cells(r, 1).Value)), varName, vbTextCompare) = 0 Then
            FindVariableRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindDataColumn(dataWs As Worksheet, varName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = dataWs.Cells(1, dataWs.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(dataWs.Cells(1, c).Value)), varName, vbTextCompare) = 0 Then
            FindDataColumn = c
            Exit Function
        End If
    Next c
End Function

Private Sub InferColumnMetadata(ws As Worksheet, col As Long, ByRef vType As String, ByRef vMeasure As String, ByRef vAlign As String)
    Dim lastRow As Long, r As Long, hasText As Boolean
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow > 50 Then lastRow = 50
    For r = 2 To lastRow
        Dim v As Variant
        v = ws.Cells(r, col).Value
        If Not IsEmpty(v) And Trim$(CStr(v)) <> "" Then
            If Not IsNumeric(v) Then
                hasText = True
                Exit For
            End If
        End If
    Next r
    If hasText Then
        vType = "String": vMeasure = "Nominal": vAlign = "Left"
    Else
        vType = "Numeric": vMeasure = "Scale": vAlign = "Right"
    End If
End Sub

Private Sub ApplyVariableViewDropdowns(varWs As Worksheet)
    Dim maxRow As Long
    maxRow = 200
    On Error Resume Next
    With varWs.Range("B2:B" & maxRow).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Numeric,Comma,Dot,Scientific,Date,Dollar,String"
    End With
    With varWs.Range("I2:I" & maxRow).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Left,Center,Right"
    End With
    With varWs.Range("J2:J" & maxRow).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Scale,Ordinal,Nominal"
    End With
    With varWs.Range("K2:K" & maxRow).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Input,Target,Both,None,Partition,Split"
    End With
    With varWs.Range("N2:N" & maxRow).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="No,Yes"
    End With
    On Error GoTo 0
End Sub

Private Function ParseValueLabels(ByVal spec As String, ByRef codes() As String, ByRef lbls() As String, ByRef count As Long) As Boolean
    Dim parts() As String, i As Long, eqPos As Long
    Dim c As String, l As String
    count = 0
    If Trim$(spec) = "" Then Exit Function
    parts = Split(spec, ";")
    ReDim codes(1 To UBound(parts) + 1)
    ReDim lbls(1 To UBound(parts) + 1)
    For i = LBound(parts) To UBound(parts)
        eqPos = InStr(1, parts(i), "=")
        If eqPos > 1 Then
            c = Trim$(Left$(parts(i), eqPos - 1))
            l = Trim$(Mid$(parts(i), eqPos + 1))
            If c <> "" And l <> "" Then
                count = count + 1
                codes(count) = c
                lbls(count) = l
            End If
        End If
    Next i
    ParseValueLabels = (count > 0)
End Function

Private Function LookupLabelByCode(ByVal codeVal As String, ByRef codes() As String, ByRef lbls() As String, ByVal count As Long) As String
    Dim i As Long
    For i = 1 To count
        If StrComp(codes(i), codeVal, vbTextCompare) = 0 Then
            LookupLabelByCode = lbls(i)
            Exit Function
        End If
    Next i
End Function

Private Function LookupCodeByLabel(ByVal labelVal As String, ByRef codes() As String, ByRef lbls() As String, ByVal count As Long) As String
    Dim i As Long
    For i = 1 To count
        If StrComp(lbls(i), labelVal, vbTextCompare) = 0 Then
            LookupCodeByLabel = codes(i)
            Exit Function
        End If
    Next i
End Function

Private Sub ApplyColumnValidationList(rng As Range, ByRef items() As String, ByVal count As Long)
    Dim listStr As String, i As Long
    If count = 0 Then Exit Sub
    For i = 1 To count
        If i > 1 Then listStr = listStr & ","
        listStr = listStr & Replace(items(i), ",", " ")
    Next i
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:=listStr
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

Private Function MatchesMissingSpec(ByVal rawVal As String, ByVal spec As String) As Boolean
    Dim tokens() As String, i As Long, tok As String
    If Trim$(spec) = "" Then Exit Function
    tokens = Split(spec, ",")
    For i = LBound(tokens) To UBound(tokens)
        tok = Trim$(tokens(i))
        If InStr(1, LCase$(tok), " thru ") > 0 Then
            Dim bounds() As String
            bounds = Split(LCase$(tok), " thru ")
            If UBound(bounds) = 1 And IsNumeric(rawVal) And IsNumeric(Trim$(bounds(0))) And IsNumeric(Trim$(bounds(1))) Then
                If CDbl(rawVal) >= CDbl(Trim$(bounds(0))) And CDbl(rawVal) <= CDbl(Trim$(bounds(1))) Then
                    MatchesMissingSpec = True
                    Exit Function
                End If
            End If
        ElseIf StrComp(rawVal, tok, vbTextCompare) = 0 Then
            MatchesMissingSpec = True
            Exit Function
        End If
    Next i
End Function

Private Function CheckCellValidation(ByVal rawVal As String, ByVal minSpec As String, ByVal maxSpec As String, ByVal restrictToLabels As Boolean, ByRef codes() As String, ByVal pairCount As Long) As String
    If IsNumeric(rawVal) Then
        If minSpec <> "" And IsNumeric(minSpec) Then
            If CDbl(rawVal) < CDbl(minSpec) Then
                CheckCellValidation = "Value " & rawVal & " is below minimum (" & minSpec & ")."
                Exit Function
            End If
        End If
        If maxSpec <> "" And IsNumeric(maxSpec) Then
            If CDbl(rawVal) > CDbl(maxSpec) Then
                CheckCellValidation = "Value " & rawVal & " exceeds maximum (" & maxSpec & ")."
                Exit Function
            End If
        End If
    End If
    If restrictToLabels And pairCount > 0 Then
        If LookupLabelByCode(rawVal, codes, codes, pairCount) = "" Then
            CheckCellValidation = "Code '" & rawVal & "' is not in defined Value Labels."
        End If
    End If
End Function

Private Sub ClearHighlightsOnSheet(ws As Worksheet, lastRow As Long, lastCol As Long)
    If lastRow < 2 Or lastCol < 1 Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
    rng.Interior.ColorIndex = xlColorIndexNone
    On Error Resume Next
    Dim c As Range
    For Each c In rng.Cells
        If Not c.Comment Is Nothing Then
            If Left$(c.Comment.Text, 9) = "StatSheet" Then c.Comment.Delete
        End If
    Next c
    On Error GoTo 0
End Sub

Private Sub AddOrUpdateNote(cell As Range, noteText As String)
    On Error Resume Next
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
    cell.AddComment noteText
    On Error GoTo 0
End Sub

Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\\", "\\\\")
    s = Replace(s, """", "\\""")
    JsonEscape = s
End Function

Private Function ValueLabelsToJson(ByVal spec As String) As String
    Dim codes() As String, lbls() As String, count As Long, i As Long
    Dim res As String
    If Not ParseValueLabels(spec, codes, lbls, count) Then
        ValueLabelsToJson = "{}"
        Exit Function
    End If
    res = "{"
    For i = 1 To count
        If i > 1 Then res = res & ", "
        res = res & """" & JsonEscape(codes(i)) & """: """ & JsonEscape(lbls(i)) & """"
    Next i
    ValueLabelsToJson = res & "}"
End Function
`;
}

export function generateWindowsXlamBuilderScript(): string {
  const ribbonXml = generateCustomUiRibbonXml().replace(/'/g, "''");
  return `# ============================================================================
# StatSheet SPSS Excel Add-in (.xlam) Automated Builder for Windows
# Builds StatSheet_SPSS.xlam WITH a native "SPSS View" Ribbon Tab directly
# into your Excel AddIns folder.
# Usage: Close Excel, put StatSheet_SPSS.bas in the same folder,
#        right-click this .ps1 file -> Run with PowerShell
# ============================================================================

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$basPath = Join-Path $scriptDir "StatSheet_SPSS.bas"

if (-not (Test-Path $basPath)) {
    Write-Host "ERROR: StatSheet_SPSS.bas not found in $scriptDir" -ForegroundColor Red
    exit 1
}

$addInsDir = Join-Path $env:APPDATA "Microsoft\\AddIns"
if (-not (Test-Path $addInsDir)) {
    New-Item -ItemType Directory -Path $addInsDir | Out-Null
}
$xlamPath = Join-Path $addInsDir "StatSheet_SPSS.xlam"

# 1. Build .xlam with VBA module via Excel COM
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    try {
        $null = $wb.VBProject.VBComponents.Import($basPath)
    } catch {
        Write-Host ""
        Write-Host "Excel blocked programmatic VBA access." -ForegroundColor Yellow
        Write-Host "To allow one-click build, open Excel -> File -> Options -> Trust Center ->" -ForegroundColor Yellow
        Write-Host "Trust Center Settings -> Macro Settings -> check 'Trust access to the VBA project object model'," -ForegroundColor Yellow
        Write-Host "then re-run this script (or import StatSheet_SPSS.bas manually via Alt+F11)." -ForegroundColor Yellow
        $wb.Close($false)
        exit 1
    }

    $wb.IsAddin = $true
    # 55 = xlOpenXMLAddIn (.xlam)
    $wb.SaveAs($xlamPath, 55)
    $wb.Close($false)
}
finally {
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        $excel = $null
    }
}

# 2. Embed native "SPSS View" Ribbon Tab (customUI14.xml) into the .xlam zip
Add-Type -AssemblyName System.IO.Compression.FileSystem
$ribbonXml = '${ribbonXml}'

$zip = [System.IO.Compression.ZipFile]::Open($xlamPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    $existingCustomUi = $zip.GetEntry("customUI/customUI14.xml")
    if ($existingCustomUi -ne $null) { $existingCustomUi.Delete() }
    $customUiEntry = $zip.CreateEntry("customUI/customUI14.xml")
    $writer = New-Object System.IO.StreamWriter($customUiEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $writer.Write($ribbonXml)
    $writer.Dispose()

    $relsEntry = $zip.GetEntry("_rels/.rels")
    if ($relsEntry -ne $null) {
        $reader = New-Object System.IO.StreamReader($relsEntry.Open())
        $relsXml = $reader.ReadToEnd()
        $reader.Dispose()

        $relTag = '<Relationship Id="StatSheetRibbon" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
        if ($relsXml -notmatch "customUI/customUI14.xml") {
            $relsXml = $relsXml -replace "</Relationships>", "$relTag</Relationships>"
            $relsEntry.Delete()
            $newRels = $zip.CreateEntry("_rels/.rels")
            $relsWriter = New-Object System.IO.StreamWriter($newRels.Open(), (New-Object System.Text.UTF8Encoding($false)))
            $relsWriter.Write($relsXml)
            $relsWriter.Dispose()
        }
    }
}
finally {
    $zip.Dispose()
}

# 3. Register and enable the add-in in Excel
$excel2 = $null
try {
    $excel2 = New-Object -ComObject Excel.Application
    $excel2.Visible = $false
    $excel2.DisplayAlerts = $false
    $addin = $excel2.AddIns.Add($xlamPath, $true)
    $addin.Installed = $true
}
finally {
    if ($excel2 -ne $null) {
        $excel2.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel2) | Out-Null
    }
}

Write-Host ""
Write-Host "SUCCESS! Installed offline add-in with native 'SPSS View' Ribbon tab at:" -ForegroundColor Green
Write-Host "  $xlamPath" -ForegroundColor Cyan
Write-Host "Open Desktop Excel and look for the 'SPSS View' tab on the top Ribbon (or right-click any cell -> StatSheet SPSS)." -ForegroundColor Green
`;
}

export function generateCustomUiRibbonXml(): string {
  return `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="StatSheetTab" label="SPSS View">
        <group id="StatSheetDataGroup" label="Views &amp; Labels">
          <button id="btnToggleLabels"
                  label="Value Labels (1↔A)"
                  size="large"
                  imageMso="OutlineShowDetail"
                  onAction="Ribbon_ToggleValueLabels" />
          <button id="btnSwitchView"
                  label="Data / Variable View"
                  size="large"
                  imageMso="SwitchWindows"
                  onAction="Ribbon_SwitchView" />
          <button id="btnSyncVars"
                  label="Sync Variable View"
                  size="normal"
                  imageMso="RefreshAll"
                  onAction="Ribbon_SyncVars" />
        </group>
        <group id="StatSheetToolsGroup" label="Codebook &amp; Stats">
          <button id="btnEditLabels"
                  label="Edit Value Labels"
                  size="normal"
                  imageMso="Properties"
                  onAction="Ribbon_EditLabels" />
          <button id="btnValidate"
                  label="Validate &amp; Missing"
                  size="normal"
                  imageMso="DataValidation"
                  onAction="Ribbon_Validate" />
          <button id="btnDescriptives"
                  label="Descriptives &amp; Freq"
                  size="normal"
                  imageMso="ChartInsert"
                  onAction="Ribbon_Descriptives" />
          <button id="btnExportJson"
                  label="Export Codebook"
                  size="normal"
                  imageMso="FileSaveAs"
                  onAction="Ribbon_ExportJson" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>`;
}
