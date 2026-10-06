' VBA Module: PO_Backfiller
' Backfills purchase order data from CSV files into the tracker sheet
'
' Usage: Call BackfillTracker with paths to CSVs and tracker file

Option Explicit

Function LoadCSVMap(filePath As String) As Object
    ' Reads a CSV file and returns a Dictionary mapping SeqNo to PO value
    ' filePath: path to CSV file (e.g., "C:\path\Fulfilled POs.csv")

    Dim dict As Object
    Dim fso As Object
    Dim file As Object
    Dim headerLine As String
    Dim line As String
    Dim seqno As String
    Dim poValue As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        Err.Raise vbObjectError + 1, "LoadCSVMap", "File not found: " & filePath
    End If

    Set file = fso.OpenTextFile(filePath, 1) ' 1 = ForReading

    headerLine = file.ReadLine() ' Skip header row

    While Not file.AtEndOfStream
        line = file.ReadLine()
        If Len(line) > 0 Then
            ' Parse CSV line (handle quoted fields)
            Call ParseCSVLine(line, seqno, poValue)

            If seqno <> "" And poValue <> "" Then
                ' If key exists, check for conflict; otherwise add it
                If dict.Exists(seqno) Then
                    If dict(seqno) <> poValue Then
                        file.Close
                        Err.Raise vbObjectError + 2, "LoadCSVMap", "Duplicate SeqNo error: " & seqno & _
                                  " has conflicting values: " & dict(seqno) & " vs " & poValue
                    End If
                    ' Otherwise same value -- skip silently
                Else
                    dict.Add seqno, poValue
                End If
            End If
        End If
    Wend

    file.Close
    Set file = Nothing
    Set fso = Nothing

    Set LoadCSVMap = dict
End Function

Sub ParseCSVLine(csvLine As String, ByRef field1 As String, ByRef field2 As String)
    ' Two-column CSV parser, handling quoted values.
    ' field1 = text before the FIRST comma that is not inside quotes
    ' field2 = everything after it (so commas inside field2 are safe)
    ' A doubled quote ("") inside a quoted field = one literal quote mark.
    ' Handles: SeqNo123,123456
    '          SeqNo124,"Mt total is $7,746"    -> field2 = Mt total is $7,746
    '          SeqNo125,"He said ""hi"" loudly" -> field2 = He said "hi" loudly

    Dim inQuotes As Boolean
    Dim currentField As String
    Dim fieldCount As Long
    Dim i As Long
    Dim char As String

    inQuotes = False
    currentField = ""
    fieldCount = 1
    field1 = ""
    field2 = ""

    For i = 1 To Len(csvLine)
        char = Mid(csvLine, i, 1)

        If char = """" Then
            If inQuotes And Mid(csvLine, i + 1, 1) = """" Then
                ' Doubled quote = one literal quote mark; skip the pair's 2nd quote
                currentField = currentField & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf char = "," And Not inQuotes And fieldCount = 1 Then
            ' End of field1; everything from here on belongs to field2
            field1 = Trim(currentField)
            currentField = ""
            fieldCount = 2
        Else
            currentField = currentField & char
        End If
    Next i

    ' Last field
    If fieldCount = 1 Then
        field1 = Trim(currentField)
    Else
        field2 = Trim(currentField)
    End If
End Sub

Function NormalizePO(poValue As Variant) As String
    ' Converts PO to canonical form: "123456"
    ' Handles: "123456", 123456, 123456.0 all normalize to same value

    Dim strValue As String
    Dim numValue As Double

    If IsEmpty(poValue) Or poValue = "" Then
        NormalizePO = ""
        Exit Function
    End If

    strValue = CStr(poValue)
    strValue = Trim(strValue)

    ' Try to convert to number and back to normalize
    On Error Resume Next
    numValue = CDbl(strValue)
    If Err.Number = 0 And numValue = Int(numValue) Then
        NormalizePO = CStr(Int(numValue))
    Else
        NormalizePO = strValue
    End If
    On Error GoTo 0
End Function

Sub FillRowColor(ws As Worksheet, row As Long, fillColor As Long)
    ' Fills entire row with a color.
    ' Pass the color in with VBA's RGB() function at the call site:
    '   FillRowColor ws, row, RGB(255, 255, 0)   ' yellow
    '   FillRowColor ws, row, RGB(255, 0, 0)     ' red

    Dim col As Long

    For col = 1 To ws.UsedRange.Columns.Count
        ws.Cells(row, col).Interior.Color = fillColor
    Next col
End Sub

Sub AddCellComment(cell As Range, commentText As String)
    ' Adds or updates a comment on a cell.
    ' Excel stamps every comment with whatever Application.UserName is set to
    ' and won't let you pass an author directly -- so borrow the setting,
    ' stamp the comment as "PO Pipeline", then put the user's name back.

    Dim oldName As String

    On Error Resume Next
    cell.ClearComments ' Remove existing comment if any
    On Error GoTo 0

    oldName = Application.UserName
    Application.UserName = "PO Pipeline"
    cell.AddComment commentText
    Application.UserName = oldName
End Sub

Function BackfillTracker(trackerFilePath As String, sheetName As String, seqnoCol As Long, poCol As Long, _
                    fulfilledCSVPath As String, unfulfilledCSVPath As String) As String
    '
    ' Main backfill macro
    '
    ' Parameters:
    '   trackerFilePath: path to tracker Excel file (e.g., "C:\Users\...\MAXISMED COMMISSION TRACKER.xlsm")
    '   sheetName: name of sheet to backfill (e.g., "Invoice Line Items")
    '   seqnoCol: column number for SeqNo (e.g., 7 for column G)
    '   poCol: column number for PO (e.g., 8 for column H)
    '   fulfilledCSVPath: path to Fulfilled POs.csv
    '   unfulfilledCSVPath: path to Unfulfilled POs.csv
    '
    ' Example call:
    '   BackfillTracker "C:\tracker.xlsm", "Invoice Line Items", 7, 8, _
    '                   "C:\output\Fulfilled POs.csv", "C:\output\Unfulfilled POs.csv"
    '

    Dim fulfilledMap As Object
    Dim unfulfilledMap As Object
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim ws As Worksheet
    Dim existingPOs As Object
    Dim row As Long
    Dim maxRow As Long
    Dim seqnoCell As Range
    Dim poCell As Range
    Dim seqno As String
    Dim normPO As String
    Dim filledCount As Long
    Dim flaggedCount As Long
    Dim duplicateCount As Long
    Dim poValue As Variant
    Dim normValue As String
    Dim unfulfilledNote As String

    On Error GoTo ErrorHandler

    ' Load CSV data into dictionaries
    Debug.Print "[Backfill] Loading fulfilled data from " & fulfilledCSVPath
    Set fulfilledMap = LoadCSVMap(fulfilledCSVPath)
    Debug.Print "[Backfill] Loaded " & fulfilledMap.Count & " fulfilled records"

    Debug.Print "[Backfill] Loading unfulfilled data from " & unfulfilledCSVPath
    Set unfulfilledMap = LoadCSVMap(unfulfilledCSVPath)
    Debug.Print "[Backfill] Loaded " & unfulfilledMap.Count & " unfulfilled records"

    ' Empty path = use the workbook this macro lives in (already open, so never closed here)
    If trackerFilePath = "" Then
        Set wb = ThisWorkbook
    Else
        Set wb = Workbooks.Open(trackerFilePath)
        openedHere = True
    End If
    Set ws = wb.Sheets(sheetName)

    ' Track existing POs to detect duplicates
    Set existingPOs = CreateObject("Scripting.Dictionary")

    ' Last data row: start at the very bottom of the SeqNo column and jump up
    ' (like Ctrl+Up) -- safer than UsedRange.Rows.Count, which only equals the
    ' last row if the sheet's used area starts at row 1
    maxRow = ws.Cells(ws.Rows.Count, seqnoCol).End(xlUp).Row

    ' Baseline scan: collect POs already on IF- rows before this run
    For row = 2 To maxRow
        Set seqnoCell = ws.Cells(row, seqnoCol)
        Set poCell = ws.Cells(row, poCol)

        seqno = CStr(seqnoCell.Value)

        If seqno <> "" And Left(seqno, 2) = "IF" Then
            normPO = NormalizePO(poCell.Value)
            If normPO <> "" Then
                If Not existingPOs.Exists(normPO) Then
                    existingPOs.Add normPO, row
                End If
            End If
        End If
    Next row

    ' Counters for summary
    filledCount = 0
    flaggedCount = 0
    duplicateCount = 0

    ' Main backfill loop
    For row = 2 To maxRow
        Set seqnoCell = ws.Cells(row, seqnoCol)
        Set poCell = ws.Cells(row, poCol)

        seqno = CStr(seqnoCell.Value)

        ' Only process IF- rows with empty PO cells
        If seqno <> "" And Left(seqno, 2) = "IF" Then
            If poCell.Value = "" Or IsEmpty(poCell.Value) Then

                ' Check if SeqNo is in fulfilled map
                If fulfilledMap.Exists(seqno) Then
                    poValue = fulfilledMap(seqno)
                    poCell.Value = poValue

                    ' Check for duplicate
                    normValue = NormalizePO(poValue)

                    If existingPOs.Exists(normValue) And existingPOs(normValue) <> row Then
                        ' This PO is already used elsewhere -- flag as duplicate
                        AddCellComment poCell, "Duplicate PO#"
                        FillRowColor ws, row, RGB(255, 255, 0) ' Yellow
                        Debug.Print "[Backfill] Row " & row & " (" & seqno & "): Duplicate PO " & poValue & " (also on row " & existingPOs(normValue) & ")"
                        duplicateCount = duplicateCount + 1
                    Else
                        existingPOs(normValue) = row
                    End If

                    filledCount = filledCount + 1
                    Debug.Print "[Backfill] Row " & row & " (" & seqno & "): PO = " & poValue

                ' Check if SeqNo is in unfulfilled map
                ElseIf unfulfilledMap.Exists(seqno) Then
                    unfulfilledNote = unfulfilledMap(seqno)

                    poCell.Value = "SCRUBSHEET DISCREPANCY"
                    AddCellComment poCell, unfulfilledNote
                    FillRowColor ws, row, RGB(255, 0, 0) ' Red

                    Debug.Print "[Backfill] Row " & row & " (" & seqno & "): Unfulfilled (note: " & unfulfilledNote & ")"
                    flaggedCount = flaggedCount + 1
                End If
            End If
        End If
    Next row

    ' Save workbook (with VBA preserved)
    wb.Save
    If openedHere Then wb.Close

    ' Print summary to Immediate window
    Debug.Print ""
    Debug.Print "[Backfill] ========== BACKFILL COMPLETE =========="
    Debug.Print "[Backfill] Filled " & filledCount & " PO(s)"
    Debug.Print "[Backfill] Flagged " & flaggedCount & " unfulfilled row(s)"
    Debug.Print "[Backfill] Flagged " & duplicateCount & " duplicate(s)"
    Debug.Print "[Backfill] ==========================================="

    BackfillTracker = "Filled: " & filledCount & "; Flagged (unfulfilled): " & flaggedCount & _
                      "; Flagged (duplicates): " & duplicateCount

    Exit Function
ErrorHandler:
    ' Don't leave a file we opened sitting locked inside a hidden Excel
    If openedHere Then wb.Close SaveChanges:=False
    BackfillTracker = "ERROR: " & Err.Description
    Debug.Print "[Backfill] ERROR: " & Err.Description
End Function
