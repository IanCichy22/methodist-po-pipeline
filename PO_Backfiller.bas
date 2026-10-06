' VBA Module: PO_Backfiller
' Backfills purchase order data from CSV files into the tracker sheet
'
' Usage: Call BackfillTracker with paths to CSVs and tracker file

Function LoadCSVMap(filePath As String) As Object
    ' Reads a CSV file and returns a Dictionary mapping SeqNo to PO value
    ' filePath: path to CSV file (e.g., "C:\path\Fulfilled POs.csv")

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        Err.Raise vbObjectError + 1, "LoadCSVMap", "File not found: " & filePath
    End If

    Dim file As Object
    Set file = fso.OpenTextFile(filePath, 1) ' 1 = ForReading

    Dim headerLine As String
    headerLine = file.ReadLine() ' Skip header row

    Dim line As String
    While Not file.AtEndOfStream
        line = file.ReadLine()
        If Len(line) > 0 Then
            Dim seqno As String
            Dim poValue As String

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
    ' Parses a CSV line into two fields, handling quoted values
    ' Handles: SeqNo123,123456
    '          SeqNo124,"Mt total is $7,746"

    Dim inQuotes As Boolean
    Dim currentField As String
    Dim fieldCount As Integer
    Dim i As Long

    inQuotes = False
    currentField = ""
    fieldCount = 1
    field1 = ""
    field2 = ""

    For i = 1 To Len(csvLine)
        Dim char As String
        char = Mid(csvLine, i, 1)

        If char = """" Then
            inQuotes = Not inQuotes
        ElseIf char = "," And Not inQuotes Then
            ' End of field
            If fieldCount = 1 Then
                field1 = Trim(currentField)
            Else
                field2 = Trim(currentField)
            End If
            currentField = ""
            fieldCount = fieldCount + 1
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

    If IsEmpty(poValue) Or poValue = "" Then
        NormalizePO = ""
        Exit Function
    End If

    Dim strValue As String
    strValue = CStr(poValue)
    strValue = Trim(strValue)

    ' Try to convert to number and back to normalize
    On Error Resume Next
    Dim numValue As Double
    numValue = CDbl(strValue)
    If Err.Number = 0 And numValue = Int(numValue) Then
        NormalizePO = CStr(Int(numValue))
    Else
        NormalizePO = strValue
    End If
    On Error GoTo 0
End Function

Sub FillRowColor(ws As Worksheet, row As Long, colorHex As String)
    ' Fills entire row with a color
    ' colorHex: RGB color like "FFFF00" (yellow) or "FF0000" (red)

    Dim col As Long
    Dim maxCol As Long
    maxCol = ws.UsedRange.Columns.Count

    For col = 1 To maxCol
        Dim r As Long, g As Long, b As Long
        r = CLngFromHex(Left(colorHex, 2))
        g = CLngFromHex(Mid(colorHex, 3, 2))
        b = CLngFromHex(Right(colorHex, 2))

        ws.Cells(row, col).Interior.Color = RGB(r, g, b)
    Next col
End Sub

Function CLngFromHex(hexStr As String) As Long
    ' Converts hex string like "FF" to decimal 255
    CLngFromHex = Val("&H" & hexStr)
End Function

Sub AddCellComment(cell As Range, commentText As String)
    ' Adds or updates a comment on a cell

    On Error Resume Next
    cell.ClearComments ' Remove existing comment if any
    On Error GoTo 0

    cell.AddComment commentText
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

    On Error GoTo ErrorHandler

    ' Load CSV data into dictionaries
    Debug.Print "[Backfill] Loading fulfilled data from " & fulfilledCSVPath
    Dim fulfilledMap As Object
    Set fulfilledMap = LoadCSVMap(fulfilledCSVPath)
    Debug.Print "[Backfill] Loaded " & fulfilledMap.Count & " fulfilled records"

    Debug.Print "[Backfill] Loading unfulfilled data from " & unfulfilledCSVPath
    Dim unfulfilledMap As Object
    Set unfulfilledMap = LoadCSVMap(unfulfilledCSVPath)
    Debug.Print "[Backfill] Loaded " & unfulfilledMap.Count & " unfulfilled records"

    ' Empty path = use the workbook this macro lives in (already open, so never closed here)
    Dim wb As Workbook
    Dim openedHere As Boolean
    If trackerFilePath = "" Then
        Set wb = ThisWorkbook
    Else
        Set wb = Workbooks.Open(trackerFilePath)
        openedHere = True
    End If
    Dim ws As Worksheet
    Set ws = wb.Sheets(sheetName)

    ' Track existing POs to detect duplicates
    Dim existingPOs As Object
    Set existingPOs = CreateObject("Scripting.Dictionary")

    ' Baseline scan: collect POs already on IF- rows before this run
    Dim row As Long
    Dim maxRow As Long
    maxRow = ws.UsedRange.Rows.Count

    For row = 2 To maxRow
        Dim seqnoCell As Range
        Dim poCell As Range

        Set seqnoCell = ws.Cells(row, seqnoCol)
        Set poCell = ws.Cells(row, poCol)

        Dim seqno As String
        seqno = CStr(seqnoCell.Value)

        If seqno <> "" And Left(seqno, 2) = "IF" Then
            Dim normPO As String
            normPO = NormalizePO(poCell.Value)
            If normPO <> "" Then
                If Not existingPOs.Exists(normPO) Then
                    existingPOs.Add normPO, row
                End If
            End If
        End If
    Next row

    ' Counters for summary
    Dim filledCount As Long
    Dim flaggedCount As Long
    Dim duplicateCount As Long
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
                    Dim poValue As Variant
                    poValue = fulfilledMap(seqno)
                    poCell.Value = poValue

                    ' Check for duplicate
                    Dim normValue As String
                    normValue = NormalizePO(poValue)

                    If existingPOs.Exists(normValue) And existingPOs(normValue) <> row Then
                        ' This PO is already used elsewhere -- flag as duplicate
                        AddCellComment poCell, "Duplicate PO#"
                        FillRowColor ws, row, "FFFF00" ' Yellow
                        Debug.Print "[Backfill] Row " & row & " (" & seqno & "): Duplicate PO " & poValue & " (also on row " & existingPOs(normValue) & ")"
                        duplicateCount = duplicateCount + 1
                    Else
                        existingPOs(normValue) = row
                    End If

                    filledCount = filledCount + 1
                    Debug.Print "[Backfill] Row " & row & " (" & seqno & "): PO = " & poValue

                ' Check if SeqNo is in unfulfilled map
                ElseIf unfulfilledMap.Exists(seqno) Then
                    Dim unfulfilledNote As String
                    unfulfilledNote = unfulfilledMap(seqno)

                    poCell.Value = "SCRUBSHEET DISCREPANCY"
                    AddCellComment poCell, unfulfilledNote
                    FillRowColor ws, row, "FF0000" ' Red

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
    BackfillTracker = "ERROR: " & Err.Description
    Debug.Print "[Backfill] ERROR: " & Err.Description
End Function
