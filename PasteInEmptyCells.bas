'==============================================================================
' PasteInEmptyCells.bas
'
' Pastes the formula currently on the clipboard into every EMPTY cell within
' the user's current selection, mimicking a manual Ctrl+V on each empty cell.
'
' HOW TO USE:
'   1. Copy any cell that contains the formula you want to spread (Ctrl+C).
'   2. Select the range where you want to paste (can include non-empty cells –
'      those will be left untouched).
'   3. Run PasteFormulaInEmptyCells  (assign it to a button or a keyboard
'      shortcut for convenience).
'
' BEHAVIOUR:
'   - Only cells that are truly empty (IsEmpty) receive the paste.
'   - Relative references in the copied formula are adjusted for each target
'     cell exactly as Excel does during a normal Ctrl+V, because the macro
'     relies on PasteSpecial xlPasteAll on a one-cell target range.
'   - If nothing is on the Excel clipboard the macro warns the user and exits.
'==============================================================================
Option Explicit

Public Sub PasteFormulaInEmptyCells()

    Dim rngSelection    As Range
    Dim rngTarget       As Range
    Dim cell            As Range
    Dim bPastedAtLeast  As Boolean

    ' ── 1. Capture the current selection ──────────────────────────────────────
    On Error Resume Next
    Set rngSelection = Selection
    On Error GoTo 0

    If rngSelection Is Nothing Then
        MsgBox "No range is selected. Please select a range first.", _
               vbExclamation, "Paste In Empty Cells"
        Exit Sub
    End If

    If Not TypeOf rngSelection Is Range Then
        MsgBox "Please select a worksheet range before running this macro.", _
               vbExclamation, "Paste In Empty Cells"
        Exit Sub
    End If

    ' ── 2. Verify something is on the Excel clipboard ─────────────────────────
    If Not ClipboardHasCopiedCell() Then
        MsgBox "Nothing is on the Excel clipboard." & vbCrLf & _
               "Please copy a cell (Ctrl+C) before running this macro.", _
               vbExclamation, "Paste In Empty Cells"
        Exit Sub
    End If

    ' ── 3. Iterate and paste only into empty cells ────────────────────────────
    Application.ScreenUpdating = False
    Application.EnableEvents   = False

    On Error GoTo ErrHandler

    bPastedAtLeast = False

    For Each cell In rngSelection.Cells
        If IsEmpty(cell) Then
            ' PasteSpecial on a single-cell range adjusts relative references
            ' the same way a manual Ctrl+V would.
            cell.PasteSpecial Paste:=xlPasteAll
            bPastedAtLeast = True
        End If
    Next cell

    ' ── 4. Clear the marching-ants border so it looks like a normal paste ──────
    Application.CutCopyMode = False

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents   = True

    If Not bPastedAtLeast Then
        MsgBox "No empty cells were found in the selection. Nothing was pasted.", _
               vbInformation, "Paste In Empty Cells"
    End If

    Exit Sub

ErrHandler:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    MsgBox "An error occurred during paste:" & vbCrLf & Err.Description, _
           vbCritical, "Paste In Empty Cells"

End Sub


' ─────────────────────────────────────────────────────────────────────────────
' Helper: returns True when Excel's own clipboard holds a copied cell/range.
' Excel clears CutCopyMode when the clipboard is empty or after it was used,
' so checking Application.CutCopyMode is the most reliable lightweight test.
' ─────────────────────────────────────────────────────────────────────────────
Private Function ClipboardHasCopiedCell() As Boolean
    ClipboardHasCopiedCell = (Application.CutCopyMode = xlCopy)
End Function
