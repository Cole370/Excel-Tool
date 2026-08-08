Attribute VB_Name = "UndoStack"
Option Explicit

'==============================================================================
' UndoStack.bas  -  Build step 2: undo snapshot + restore.
'
' A VBA macro that writes to cells WIPES Excel's built-in undo stack, so Ctrl+Z
' cannot rescue the user. Before every write batch the tool must snapshot the
' target cells' current .Formula values; RestoreLastBatch() puts them back.
'
' Storage: a very-hidden sheet "_ClaudeUndo" inside the workbook that owns the
' cells. One row per cell: BatchID | Timestamp | Sheet | Address | Formula.
' The last MAX_BATCHES batches are kept; older ones are pruned.  (Spec 2.2)
'==============================================================================

Private Const UNDO_SHEET As String = "_ClaudeUndo"
Private Const MAX_BATCHES As Long = 10

' Test tally (used only by Test_Undo)
Private uPass As Long
Private uTotal As Long

'------------------------------------------------------------------------------
' SnapshotRange - save a range's current formulas as a new undo batch.
' Returns the assigned BatchID. Call this BEFORE writing to the range.
'------------------------------------------------------------------------------
Public Function SnapshotRange(ByVal rng As Range) As Long
    Dim wb As Workbook, ws As Worksheet, c As Range
    Dim batchId As Long, r As Long, stamp As String

    If rng Is Nothing Then Err.Raise vbObjectError + 100, , "SnapshotRange: range is Nothing"
    Set wb = rng.Worksheet.Parent
    Set ws = EnsureUndoSheet(wb)

    batchId = NextBatchId(ws)
    stamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2

    For Each c In rng.Cells
        ' Text-format the row so a stored "=..." is kept as text, not evaluated.
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).NumberFormat = "@"
        ws.Cells(r, 1).Value = batchId
        ws.Cells(r, 2).Value = stamp
        ws.Cells(r, 3).Value = c.Worksheet.Name
        ws.Cells(r, 4).Value = c.Address
        ws.Cells(r, 5).Value = c.Formula
        r = r + 1
    Next c

    PruneBatches ws
    SnapshotRange = batchId
End Function

'------------------------------------------------------------------------------
' RestoreLastBatch - put the most recent snapshotted batch back into the cells,
' then remove it from the stack. This is the user's "undo" for the tool.
'------------------------------------------------------------------------------
Public Sub RestoreLastBatch()
    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, i As Long, targetBatch As Long
    Dim restored As Long, failed As Long, stamp As String
    Dim prevEvents As Boolean, prevScreen As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook to restore into.", vbExclamation, "Claude Restore"
        Exit Sub
    End If

    On Error Resume Next
    Set ws = wb.Worksheets(UNDO_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Nothing to restore - this workbook has no Claude undo history.", _
               vbInformation, "Claude Restore"
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Nothing to restore - the undo history is empty.", _
               vbInformation, "Claude Restore"
        Exit Sub
    End If

    targetBatch = CLng(Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow)))

    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    On Error GoTo CleanFail
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For i = 2 To lastRow
        If CLng(ws.Cells(i, 1).Value) = targetBatch Then
            stamp = CStr(ws.Cells(i, 2).Value)
            On Error Resume Next
            Err.Clear
            wb.Worksheets(CStr(ws.Cells(i, 3).Value)) _
              .Range(CStr(ws.Cells(i, 4).Value)).Formula = ws.Cells(i, 5).Value
            If Err.Number <> 0 Then failed = failed + 1 Else restored = restored + 1
            On Error GoTo CleanFail
        End If
    Next i

    ' Consume the batch: delete its rows (bottom-up so indices stay valid).
    For i = lastRow To 2 Step -1
        If CLng(ws.Cells(i, 1).Value) = targetBatch Then ws.Rows(i).Delete
    Next i

    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    Dim msg As String
    msg = "Restored " & restored & " cell(s) from batch " & targetBatch & _
          " (saved " & stamp & ")."
    If failed > 0 Then msg = msg & vbCrLf & failed & " cell(s) could not be restored " & _
                                "(their sheet may have been deleted)."
    MsgBox msg, vbInformation, "Claude Restore"
    Exit Sub

CleanFail:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "Restore hit an error: " & Err.Description, vbExclamation, "Claude Restore"
End Sub

'------------------------------------------------------------------------------
' Internal helpers
'------------------------------------------------------------------------------
Private Function EnsureUndoSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(UNDO_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = UNDO_SHEET
        ws.Range("A1:E1").Value = Array("BatchID", "Timestamp", "Sheet", "Address", "Formula")
    End If
    ws.Visible = xlSheetVeryHidden
    Set EnsureUndoSheet = ws
End Function

Private Function NextBatchId(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        NextBatchId = 1
    Else
        NextBatchId = CLng(Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow))) + 1
    End If
End Function

Private Sub PruneBatches(ByVal ws As Worksheet)
    Dim lastRow As Long, i As Long, maxB As Long, cutoff As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    maxB = CLng(Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow)))
    cutoff = maxB - MAX_BATCHES     ' keep only batches with ID > cutoff
    If cutoff < 1 Then Exit Sub
    For i = lastRow To 2 Step -1
        If CLng(ws.Cells(i, 1).Value) <= cutoff Then ws.Rows(i).Delete
    Next i
End Sub

'==============================================================================
' Demo_WriteToSelectedCell - Gate 5 helper. Snapshots the selected cell, then
' overwrites it, so the user can watch RestoreLastBatch put it back.
'==============================================================================
Public Sub Demo_WriteToSelectedCell()
    Dim c As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Click a cell first, then run this again.", vbExclamation, "Claude Demo"
        Exit Sub
    End If
    Set c = ActiveCell
    SnapshotRange c
    c.Formula = "CLAUDE_WAS_HERE"
    MsgBox "I saved cell " & c.Address(False, False) & ", then overwrote it with " & _
           "CLAUDE_WAS_HERE." & vbCrLf & vbCrLf & _
           "Ctrl+Z will NOT undo this. Now run RestoreLastBatch to put it back.", _
           vbInformation, "Claude Demo"
End Sub

'==============================================================================
' Test_Undo - automated self-test on a throwaway sheet. Run with F5.
' Expected final line: === RESULT: ALL 6 TESTS PASSED ===
'==============================================================================
Public Sub Test_Undo()
    Dim wb As Workbook, ts As Worksheet
    Dim prevAlerts As Boolean
    uPass = 0
    uTotal = 0
    Debug.Print "=== UNDO STACK TESTS ==="

    Set wb = ActiveWorkbook
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ' Fresh throwaway sheet
    On Error Resume Next
    wb.Worksheets("_ClaudeUndoTest").Delete
    On Error GoTo 0
    Set ts = wb.Worksheets.Add
    ts.Name = "_ClaudeUndoTest"

    ' Seed known content: a formula, a number, text, and an empty cell.
    ts.Range("A1").Formula = "=1+1"
    ts.Range("A2").Value = 42
    ts.Range("A3").Value = "hello"
    ' A4 intentionally left empty

    ' Snapshot, then clobber everything.
    SnapshotRange ts.Range("A1:A4")
    ts.Range("A1:A4").Value = "XXX"

    AssertU "clobber changed the formula cell", (ts.Range("A1").Value = "XXX"), True

    ' Restore and verify each cell came back.
    RestoreLastBatch

    AssertU "formula restored (=1+1)", (ts.Range("A1").Formula = "=1+1"), True
    AssertU "number restored (42)", (ts.Range("A2").Value = 42), True
    AssertU "text restored (hello)", (ts.Range("A3").Value = "hello"), True
    AssertU "empty cell stayed empty", (Len(ts.Range("A4").Formula) = 0), True

    ' The batch should have been consumed (no rows left for it).
    Dim uh As Worksheet, leftover As Long
    Set uh = wb.Worksheets(UNDO_SHEET)
    leftover = uh.Cells(uh.Rows.Count, 1).End(xlUp).Row
    AssertU "batch consumed after restore", (leftover < 2), True

    ' Clean up the throwaway sheet.
    ts.Delete
    Application.DisplayAlerts = prevAlerts

    Debug.Print "-------------------------------------"
    If uPass = uTotal Then
        Debug.Print "=== RESULT: ALL " & uTotal & " TESTS PASSED ==="
    Else
        Debug.Print "=== RESULT: " & uPass & "/" & uTotal & " passed - " & (uTotal - uPass) & " FAILED ==="
    End If
End Sub

Private Sub AssertU(ByVal desc As String, ByVal got As Boolean, ByVal expected As Boolean)
    uTotal = uTotal + 1
    If got = expected Then
        uPass = uPass + 1
        Debug.Print "[PASS] " & desc
    Else
        Debug.Print "[FAIL] " & desc & "  (expected " & expected & ", got " & got & ")"
    End If
End Sub
