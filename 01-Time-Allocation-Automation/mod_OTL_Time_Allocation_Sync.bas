Option Explicit

'====================================================================
' Module Name:  mod_OTL_Time_Allocation_Sync
' Purpose:      Automates pulling raw time-tracking data, calculates
'               proportional allocation hours, formats output, and 
'               performs a two-way write-back with deduplication & zero-hour cleanup.
' Author:       Ting Huang
'====================================================================

Sub Macro1_ExecuteTimeAllocation()
    Dim StartRow As Long, EndRow As Long
    Dim TC As Long, SR As Long, LR As Long, SA As Long
    Dim Count As Long, CheckCount As Long
    Dim i As Long
    Dim InputCode As String
    Dim wsInput As Worksheet, wsOutput As Worksheet, wsAlloc As Worksheet
    Dim LastRowOutput As Long
    
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsOutput = ThisWorkbook.Sheets("Output")
    Set wsAlloc = ThisWorkbook.Sheets("Allocation")
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Loading data from Master OTL Loader..."
    
    ' Pull data from source file
    If Not LoadInputFromCALoader(wsInput) Then
        GoTo CleanExit
    End If
    
    Application.StatusBar = "Processing allocation calculations..."
    
    wsOutput.Cells.Clear
    wsInput.Range("A1:E1").Copy wsOutput.Range("A1")
    wsOutput.Range("F1").Value = "Allocation %"
    wsOutput.Range("G1").Value = "Calculated Hours"
    wsOutput.Range("H1").Value = "Variance"
    
    CheckCount = Application.CountA(wsInput.Range("A:A")) - 1
    
    If CheckCount <= 0 Then
        MsgBox "No matching data found in Master OTL Loader for the specified tab.", vbInformation
        GoTo CleanExit
    End If
    
    SR = 2
    SA = 2
    Count = 1
    
    Do Until Count > CheckCount
        
        InputCode = CStr(wsInput.Range("C" & SA).Value)
        
        ' Map code rules to allocation range bounds
        Select Case InputCode
            Case "4000002"
                StartRow = 2
                EndRow = 14
            Case "4031995"
                StartRow = 17
                EndRow = 32
            Case "4048807"
                StartRow = 35
                EndRow = 46
            Case Else
                SA = SA + 1
                Count = Count + 1
                GoTo NextIter
        End Select
        
        TC = EndRow - StartRow + 1
        LR = SR + TC - 1
        
        ' Copy Columns A/B (Dates & Metadata)
        wsInput.Range("A" & SA & ":B" & SA).Copy
        wsOutput.Range("A" & SR & ":B" & LR).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        wsOutput.Range("B" & SR & ":B" & LR).NumberFormat = "mm/dd/yyyy"
        
        ' Copy Column C - client codes & formatting
        wsAlloc.Range("J" & StartRow & ":J" & EndRow).Copy
        wsOutput.Range("C" & SR).PasteSpecial xlPasteValues
        For i = 0 To TC - 1
            wsOutput.Range("C" & (SR + i)).Interior.Color = _
                wsAlloc.Range("J" & (StartRow + i)).Interior.Color
        Next i
        
        ' Copy Column D - description & formatting
        wsAlloc.Range("L" & StartRow & ":L" & EndRow).Copy
        wsOutput.Range("D" & SR).PasteSpecial xlPasteValues
        For i = 0 To TC - 1
            wsOutput.Range("D" & (SR + i)).Interior.Color = _
                wsAlloc.Range("L" & (StartRow + i)).Interior.Color
        Next i
        
        ' Copy Column E - raw hours
        wsInput.Range("E" & SA).Copy
        wsOutput.Range("E" & SR & ":E" & LR).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        
        ' Copy Column F - allocation % & formatting
        wsAlloc.Range("K" & StartRow & ":K" & EndRow).Copy
        wsOutput.Range("F" & SR).PasteSpecial xlPasteValues
        For i = 0 To TC - 1
            wsOutput.Range("F" & (SR + i)).Interior.Color = _
                wsAlloc.Range("K" & (StartRow + i)).Interior.Color
        Next i
        
        ' Populate Columns G/H formulas (Calculated Hours & Variance)
        wsOutput.Range("G" & SR & ":G" & LR - 1).FormulaR1C1 = "=ROUND(RC[-2]*RC[-1],1)"
        wsOutput.Range("H" & SR & ":H" & LR - 1).FormulaR1C1 = "=RC[-1]"
        wsOutput.Range("G" & LR).FormulaR1C1 = "=SUM(R" & SR & "C:R" & LR - 1 & "C)"
        wsOutput.Range("H" & LR).FormulaR1C1 = "=RC[-3]-RC[-1]"
        
        ' Apply color formatting for G/H
        For i = 0 To TC - 1
            wsOutput.Range("G" & (SR + i)).Interior.Color = _
                wsAlloc.Range("M" & (StartRow + i)).Interior.Color
            wsOutput.Range("H" & (SR + i)).Interior.Color = _
                wsAlloc.Range("N" & (StartRow + i)).Interior.Color
        Next i
        
        SR = LR + 1
        SA = SA + 1
        Count = Count + 1

NextIter:
    Loop
    
    ' Freeze values while preserving formatting
    wsOutput.UsedRange.Copy
    wsOutput.UsedRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    LastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row
    If LastRowOutput >= 2 Then
        wsOutput.Range("B2:B" & LastRowOutput).NumberFormat = "mm/dd/yyyy"
    End If
    
    ' Confirm before writing back to source master file
    If MsgBox("Output generated successfully." & vbCrLf & vbCrLf & _
               "Write this data back to Master OTL Loader now?", _
               vbYesNo + vbQuestion, "Confirm Write-Back") = vbYes Then
        Call PushOutputBackToSource
    End If

CleanExit:
    If Not wsOutput Is Nothing Then
        wsOutput.Activate
        wsOutput.Range("A1").Select
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
End Sub


' Helper Function: Pulls matching source data from Master OTL Loader
Private Function LoadInputFromCALoader(ByVal wsInput As Worksheet) As Boolean
    Dim srcPath As String
    Dim srcTabName As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim lastSrcRow As Long
    Dim r As Long, destRow As Long
    Dim codeVal As String, descVal As String
    Dim matched As Boolean
    
    LoadInputFromCALoader = False
    
    ' Path anonymized for portfolio/public safety
    srcPath = ThisWorkbook.Path & "\CA_Loader_OTL.xlsx"
    
    On Error Resume Next
    srcTabName = CStr(ThisWorkbook.Sheets("Button").Range("B1").Value)
    On Error GoTo 0
    
    If Len(Trim(srcTabName)) = 0 Then
        MsgBox "Button!B1 is empty. Please enter the source tab name.", vbExclamation
        Exit Function
    End If
    
    If Dir(srcPath) = "" Then
        MsgBox "Source file not found at path:" & vbCrLf & srcPath, vbCritical
        Exit Function
    End If
    
    On Error Resume Next
    Set wbSrc = Workbooks.Open(Filename:=srcPath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0
    
    If wbSrc Is Nothing Then
        MsgBox "Unable to open source file:" & vbCrLf & srcPath, vbCritical
        Exit Function
    End If
    
    On Error Resume Next
    Set wsSrc = wbSrc.Sheets(srcTabName)
    On Error GoTo 0
    
    If wsSrc Is Nothing Then
        MsgBox "Tab '" & srcTabName & "' not found in source file.", vbCritical
        wbSrc.Close SaveChanges:=False
        Exit Function
    End If
    
    Dim lastInputRow As Long
    lastInputRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row
    If lastInputRow >= 2 Then
        wsInput.Range("A2:E" & lastInputRow).Clear
    End If
    
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row
    destRow = 2
    
    For r = 2 To lastSrcRow
        codeVal = CStr(wsSrc.Range("C" & r).Value)
        descVal = CStr(wsSrc.Range("D" & r).Value)
        matched = False
        
        Select Case codeVal
            Case "4000002"
                If Trim(descVal) = "ADMIN-005-General Admin" Then matched = True ' Anonymized description
            Case "4031995", "4048807"
                matched = True
        End Select
        
        If matched Then
            wsInput.Range("A" & destRow).Value = wsSrc.Range("A" & r).Value
            wsInput.Range("B" & destRow).Value = wsSrc.Range("B" & r).Value
            wsInput.Range("C" & destRow).Value = wsSrc.Range("C" & r).Value
            wsInput.Range("D" & destRow).Value = wsSrc.Range("D" & r).Value
            wsInput.Range("E" & destRow).Value = wsSrc.Range("E" & r).Value
            wsInput.Range("B" & destRow).NumberFormat = "mm/dd/yyyy"
            destRow = destRow + 1
        End If
    Next r
    
    wbSrc.Close SaveChanges:=False
    LoadInputFromCALoader = True
End Function


' Writes allocated output back to Master OTL and deletes staging rows
Sub PushOutputBackToSource()
    Dim wsOutput As Worksheet
    Dim srcPath As String, srcTabName As String
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim lastSrcRow As Long, lastOutRow As Long
    Dim r As Long, destRow As Long, i As Long
    Dim codeVal As String, descVal As String
    Dim matched As Boolean
    
    Set wsOutput = ThisWorkbook.Sheets("Output")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Writing Output back to Master OTL..."
    
    srcPath = ThisWorkbook.Path & "\CA_Loader_OTL.xlsx"
    
    On Error Resume Next
    srcTabName = CStr(ThisWorkbook.Sheets("Button").Range("B1").Value)
    On Error GoTo 0
    
    If Len(Trim(srcTabName)) = 0 Then
        MsgBox "Button!B1 is empty. Please enter the source tab name.", vbExclamation
        GoTo CleanExit
    End If
    
    If Dir(srcPath) = "" Then
        MsgBox "Source file not found at path:" & vbCrLf & srcPath, vbCritical
        GoTo CleanExit
    End If
    
    lastOutRow = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row
    If lastOutRow < 2 Then
        MsgBox "Output sheet is empty. Nothing to write back.", vbExclamation
        GoTo CleanExit
    End If
    
    On Error Resume Next
    Set wbSrc = Workbooks.Open(Filename:=srcPath, ReadOnly:=False, UpdateLinks:=0)
    On Error GoTo 0
    
    If wbSrc Is Nothing Then
        MsgBox "Unable to open source file (file may be locked):" & vbCrLf & srcPath, vbCritical
        GoTo CleanExit
    End If
    
    On Error Resume Next
    Set wsSrc = wbSrc.Sheets(srcTabName)
    On Error GoTo 0
    
    If wsSrc Is Nothing Then
        MsgBox "Tab '" & srcTabName & "' not found in source file.", vbCritical
        wbSrc.Close SaveChanges:=False
        GoTo CleanExit
    End If
    
    ' Step 1: Remove original staging/input records from source
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row
    
    For r = lastSrcRow To 2 Step -1
        codeVal = CStr(wsSrc.Range("C" & r).Value)
        descVal = CStr(wsSrc.Range("D" & r).Value)
        matched = False
        
        Select Case codeVal
            Case "4000002"
                If Trim(descVal) = "ADMIN-005-General Admin" Then matched = True
            Case "4031995", "4048807"
                matched = True
        End Select
        
        If matched Then
            wsSrc.Rows(r).Delete
        End If
    Next r
    
    ' Step 2: Append newly calculated allocated entries
    destRow = wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row + 1
    If destRow < 2 Then destRow = 2
    
    For i = 2 To lastOutRow
        If Len(Trim(CStr(wsOutput.Range("C" & i).Value))) = 0 Then GoTo NextOutRow
        
        wsSrc.Range("A" & destRow).Value = wsOutput.Range("A" & i).Value
        wsSrc.Range("B" & destRow).Value = wsOutput.Range("B" & i).Value
        wsSrc.Range("C" & destRow).Value = wsOutput.Range("C" & i).Value
        wsSrc.Range("D" & destRow).Value = wsOutput.Range("D" & i).Value
        wsSrc.Range("E" & destRow).Value = wsOutput.Range("H" & i).Value
        wsSrc.Range("B" & destRow).NumberFormat = "mm/dd/yyyy"
        
        ' Copy formatting attributes
        wsOutput.Range("A" & i).Copy
        wsSrc.Range("A" & destRow).PasteSpecial Paste:=xlPasteFormats
        wsOutput.Range("B" & i).Copy
        wsSrc.Range("B" & destRow).PasteSpecial Paste:=xlPasteFormats
        wsOutput.Range("C" & i).Copy
        wsSrc.Range("C" & destRow).PasteSpecial Paste:=xlPasteFormats
        wsOutput.Range("D" & i).Copy
        wsSrc.Range("D" & destRow).PasteSpecial Paste:=xlPasteFormats
        wsOutput.Range("H" & i).Copy
        wsSrc.Range("E" & destRow).PasteSpecial Paste:=xlPasteFormats
        
        wsSrc.Range("B" & destRow).NumberFormat = "mm/dd/yyyy"
        destRow = destRow + 1
NextOutRow:
    Next i
    
    Application.CutCopyMode = False
    
    ' Step 3: Deduplicate, aggregate, and purge zero-hour rows
    Application.StatusBar = "Deduplicating and aggregating records..."
    Call DedupAndSumSource(wsSrc)
    
    ' Step 4: Save master source
    wbSrc.Save
    wbSrc.Close SaveChanges:=False
    
    MsgBox "Output successfully written back to Master OTL Loader (Deduplicated & Aggregated).", vbInformation

CleanExit:
    Application.CutCopyMode = False
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


' Deduplicates source rows using a Composite Key (A+B+C+D), aggregates E, and purges E=0
Private Sub DedupAndSumSource(ByVal wsSrc As Worksheet)
    Dim dict As Object
    Dim delDict As Object
    Dim lastRow As Long, r As Long
    Dim keyStr As String
    Dim eVal As Double
    Dim firstRow As Long
    Dim delUnion As Range
    Dim k As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set delDict = CreateObject("Scripting.Dictionary")
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' Pass 1: Build composite keys and aggregate values into initial instance
    For r = 2 To lastRow
        keyStr = CStr(wsSrc.Range("A" & r).Value) & "|" & _
                 CStr(wsSrc.Range("B" & r).Value) & "|" & _
                 CStr(wsSrc.Range("C" & r).Value) & "|" & _
                 CStr(wsSrc.Range("D" & r).Value)
        
        If IsNumeric(wsSrc.Range("E" & r).Value) Then
            eVal = CDbl(wsSrc.Range("E" & r).Value)
        Else
            eVal = 0
        End If
        
        If dict.Exists(keyStr) Then
            firstRow = dict(keyStr)
            wsSrc.Range("E" & firstRow).Value = _
                CDbl(wsSrc.Range("E" & firstRow).Value) + eVal
            If Not delDict.Exists(r) Then delDict.Add r, True
        Else
            dict.Add keyStr, r
        End If
    Next r
    
    ' Pass 2: Flag zero-sum rows for deletion
    For Each k In dict.Keys
        firstRow = dict(k)
        If IsNumeric(wsSrc.Range("E" & firstRow).Value) Then
            If CDbl(wsSrc.Range("E" & firstRow).Value) = 0 Then
                If Not delDict.Exists(firstRow) Then delDict.Add firstRow, True
            End If
        End If
    Next k
    
    If delDict.Count = 0 Then Exit Sub
    
    ' Execute batch range deletion
    For Each k In delDict.Keys
        If delUnion Is Nothing Then
            Set delUnion = wsSrc.Rows(CLng(k))
        Else
            Set delUnion = Union(delUnion, wsSrc.Rows(CLng(k)))
        End If
    Next k
    
    If Not delUnion Is Nothing Then delUnion.Delete
End Sub
