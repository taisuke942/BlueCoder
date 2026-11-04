Option Explicit

' HULFT8 for Windows / Excel VBA
' Source: utligen cards (snd/rcv), utllist -s/-r -v84 histories.
' Paste ALL lines into one standard module (Alt+F11 > Insert > Module).
' Execute: HULFT_4Files_Import
' Windows Excel only. No references required (late binding).

Private Const DEFAULT_TEXT_CHARSET As String = "shift_jis"
Private Const MASK_SECRET_VALUES As Boolean = True
Private Const HISTORY_BATCH_SIZE As Long = 2000
Private Const EXCEL_MAX_ROW As Long = 1048576

Public Sub HULFT_4Files_Import()
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim report As String

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    On Error GoTo FatalError
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    report = "HULFT import results" & vbCrLf & String$(24, "-") & vbCrLf
    report = report & ImportWithDialog("配信管理", "SNDFILE", False) & vbCrLf
    report = report & ImportWithDialog("集信管理", "RCVFILE", False) & vbCrLf
    report = report & ImportWithDialog("配信履歴", "", True) & vbCrLf
    report = report & ImportWithDialog("集信履歴", "", True)

    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents

    MsgBox report & vbCrLf & vbCrLf & _
           "Save the workbook as .xlsm to keep this macro.", _
           vbInformation, "HULFT text import"
    Exit Sub

FatalError:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    MsgBox "Import stopped: " & Err.Description, vbExclamation
End Sub

Private Function ImportWithDialog(ByVal sheetName As String, _
                                  ByVal cardTag As String, _
                                  ByVal isHistory As Boolean) As String
    Dim picked As Variant
    Dim count As Long, invalidLines As Long
    Dim suffix As String

    On Error GoTo ImportError

    picked = Application.GetOpenFilename( _
        "Text files (*.txt;*.prm;*.dat),*.txt;*.prm;*.dat,All files (*.*),*.*", _
        1, "Choose HULFT file for: " & sheetName, , False)

    If VarType(picked) = vbBoolean Then
        ImportWithDialog = sheetName & ": skipped"
        Exit Function
    End If

    If isHistory Then
        count = ImportHistoryV84(CStr(picked), sheetName, invalidLines)
        If invalidLines > 0 Then
            suffix = " (ignored " & CStr(invalidLines) & " non-V84 lines)"
        End If
    Else
        count = ImportDefinitionCards(CStr(picked), sheetName, cardTag)
    End If

    ImportWithDialog = sheetName & ": " & CStr(count) & " rows" & suffix
    Exit Function

ImportError:
    ImportWithDialog = sheetName & ": ERROR - " & Err.Description
End Function

Private Function ImportDefinitionCards(ByVal filePath As String, _
                                       ByVal sheetName As String, _
                                       ByVal expectedTag As String) As Long
    Dim stm As Object, colIndex As Object, card As Object
    Dim records As Collection
    Dim line As String, key As String, value As String
    Dim eq As Long, rowNum As Long, colNum As Long
    Dim keys As Variant, matrix() As Variant
    Dim ws As Worksheet

    Set records = New Collection
    Set colIndex = CreateObject("Scripting.Dictionary")
    colIndex.CompareMode = vbTextCompare
    colIndex.Add "FILEID", True

    Set stm = OpenHulftText(filePath)

    Do Until stm.EOS
        line = Trim$(stm.ReadText(-2))
        line = StripBom(line)
        If Len(line) = 0 Then GoTo NextCardLine
        If Left$(line, 1) = "#" Or Left$(line, 1) = ";" Then GoTo NextCardLine

        If UCase$(line) = "END" Then
            If Not card Is Nothing Then
                If Len(CStr(card("FILEID"))) > 0 Then records.Add card
                Set card = Nothing
            End If
            GoTo NextCardLine
        End If

        eq = InStr(1, line, "=", vbBinaryCompare)
        If eq <= 1 Then GoTo NextCardLine

        key = UCase$(Trim$(Left$(line, eq - 1)))
        value = Trim$(Mid$(line, eq + 1))

        If key = "SNDFILE" Or key = "RCVFILE" Then
            ' Accept a new card even if a previous card had no END.
            If Not card Is Nothing Then
                If Len(CStr(card("FILEID"))) > 0 Then records.Add card
            End If
            Set card = Nothing

            If key = expectedTag And Len(value) > 0 Then
                Set card = CreateObject("Scripting.Dictionary")
                card.CompareMode = vbTextCompare
                card.Add "FILEID", value
            End If
        ElseIf Not card Is Nothing Then
            If MASK_SECRET_VALUES And IsSecretParameter(key) Then
                value = "[MASKED]"
            End If

            If card.Exists(key) Then
                card(key) = value
            Else
                card.Add key, value
            End If

            If Not colIndex.Exists(key) Then colIndex.Add key, True
        End If

NextCardLine:
    Loop

    If Not card Is Nothing Then
        If Len(CStr(card("FILEID"))) > 0 Then records.Add card
    End If
    stm.Close

    If records.Count = 0 Then
        Err.Raise vbObjectError + 8101, , _
            "No valid " & expectedTag & " cards found. Select utligen output."
    End If
    If records.Count >= EXCEL_MAX_ROW Then
        Err.Raise vbObjectError + 8102, , "Too many rows for one Excel sheet."
    End If

    keys = colIndex.Keys
    ReDim matrix(1 To records.Count + 1, 1 To colIndex.Count)

    For colNum = 0 To UBound(keys)
        matrix(1, colNum + 1) = CStr(keys(colNum))
    Next colNum

    For rowNum = 1 To records.Count
        Set card = records(rowNum)
        For colNum = 0 To UBound(keys)
            key = CStr(keys(colNum))
            If card.Exists(key) Then
                matrix(rowNum + 1, colNum + 1) = _
                    SafeCellString(CStr(card(key)))
            Else
                matrix(rowNum + 1, colNum + 1) = ""
            End If
        Next colNum
    Next rowNum

    Set ws = PrepareOutputSheet(sheetName)
    With ws.Cells(1, 1).Resize(records.Count + 1, colIndex.Count)
        .NumberFormat = "@"
        .Value2 = matrix
    End With
    StyleOutputSheet ws, records.Count + 1, colIndex.Count

    ImportDefinitionCards = records.Count
End Function

Private Function ImportHistoryV84(ByVal filePath As String, _
                                  ByVal sheetName As String, _
                                  ByRef invalidLines As Long) As Long
    Dim stm As Object, ws As Worksheet
    Dim line As String
    Dim batch() As Variant
    Dim found As Long, buffered As Long
    Dim firstRow As Long

    ReDim batch(1 To HISTORY_BATCH_SIZE, 1 To 10)
    Set stm = OpenHulftText(filePath)

    Do Until stm.EOS
        line = StripBom(stm.ReadText(-2))

        If IsValidV84Row(line) Then
            If ws Is Nothing Then
                Set ws = PrepareOutputSheet(sheetName)
                ws.Range("A:J").NumberFormat = "@"
                WriteHistoryHeaders ws
            End If

            If found >= EXCEL_MAX_ROW - 1 Then
                Err.Raise vbObjectError + 8103, , _
                    "History exceeds the Excel row limit."
            End If

            found = found + 1
            buffered = buffered + 1

            ' The positions below are 1-based character offsets.
            ' For HULFT V84, all fields before FILENAME are ASCII.
            batch(buffered, 1) = SafeCellString(Trim$(Mid$(line, 1, 50)))
            batch(buffered, 2) = SafeCellString(Trim$(Mid$(line, 53, 68)))
            batch(buffered, 3) = Trim$(Mid$(line, 123, 10))
            batch(buffered, 4) = Trim$(Mid$(line, 135, 8))
            batch(buffered, 5) = Trim$(Mid$(line, 145, 8))
            batch(buffered, 6) = Trim$(Mid$(line, 155, 18))
            batch(buffered, 7) = Trim$(Mid$(line, 175, 18))
            batch(buffered, 8) = Trim$(Mid$(line, 195, 12))
            batch(buffered, 9) = Trim$(Mid$(line, 209, 3))
            batch(buffered, 10) = SafeCellString(Trim$(Mid$(line, 214)))

            If buffered = HISTORY_BATCH_SIZE Then
                firstRow = 2 + found - buffered
                WriteHistoryBatch ws, batch, buffered, firstRow
                buffered = 0
                Erase batch
                ReDim batch(1 To HISTORY_BATCH_SIZE, 1 To 10)
            End If
        ElseIf Len(Trim$(line)) > 0 Then
            invalidLines = invalidLines + 1
        End If
    Loop

    stm.Close

    If found = 0 Then
        Err.Raise vbObjectError + 8104, , _
            "No HULFT -v84 history rows found. Export with utllist -s/-r -v84."
    End If

    If buffered > 0 Then
        firstRow = 2 + found - buffered
        WriteHistoryBatch ws, batch, buffered, firstRow
    End If

    StyleOutputSheet ws, found + 1, 10
    ImportHistoryV84 = found
End Function

Private Function IsValidV84Row(ByVal line As String) As Boolean
    ' HULFT V84 fixed-width: DATE starts at byte 122 (0-based),
    ' START TIME at 134, END TIME at 144, STATUS at 194.
    ' The prefix uses ASCII, so byte and VBA character offsets agree.
    If Len(line) < 213 Then Exit Function
    If Mid$(line, 137, 1) <> ":" Then Exit Function
    If Mid$(line, 140, 1) <> ":" Then Exit Function
    If Mid$(line, 147, 1) <> ":" Then Exit Function
    If Mid$(line, 150, 1) <> ":" Then Exit Function
    If Mid$(line, 201, 1) <> "-" Then Exit Function
    If Trim$(Mid$(line, 1, 50)) = "" Then Exit Function
    IsValidV84Row = True
End Function

Private Sub WriteHistoryHeaders(ByVal ws As Worksheet)
    Dim hdr(1 To 1, 1 To 10) As Variant
    hdr(1, 1) = "FILEID"
    hdr(1, 2) = "HOST"
    hdr(1, 3) = "開始日"
    hdr(1, 4) = "開始時刻"
    hdr(1, 5) = "終了時刻"
    hdr(1, 6) = "レコード件数"
    hdr(1, 7) = "データサイズ"
    hdr(1, 8) = "ステータス"
    hdr(1, 9) = "接続形態"
    hdr(1, 10) = "ファイル名"
    ws.Range("A1").Resize(1, 10).Value2 = hdr
End Sub

Private Sub WriteHistoryBatch(ByVal ws As Worksheet, _
                              ByRef batch() As Variant, _
                              ByVal usedRows As Long, _
                              ByVal firstRow As Long)
    Dim small() As Variant, r As Long, c As Long

    If usedRows = HISTORY_BATCH_SIZE Then
        ws.Cells(firstRow, 1).Resize(usedRows, 10).Value2 = batch
    Else
        ReDim small(1 To usedRows, 1 To 10)
        For r = 1 To usedRows
            For c = 1 To 10
                small(r, c) = batch(r, c)
            Next c
        Next r
        ws.Cells(firstRow, 1).Resize(usedRows, 10).Value2 = small
    End If
End Sub

Private Function OpenHulftText(ByVal filePath As String) As Object
    Dim raw As Object, result As Object
    Dim bytes As Variant, encoding As String
    Dim lo As Long, size As Long

    encoding = DEFAULT_TEXT_CHARSET

    Set raw = CreateObject("ADODB.Stream")
    raw.Type = 1 ' adTypeBinary
    raw.Open
    raw.LoadFromFile filePath
    size = CLng(raw.Size)

    If size >= 2 Then
        bytes = raw.Read(IIf(size >= 3, 3, 2))
        lo = LBound(bytes)

        If UBound(bytes) - lo >= 2 Then
            If bytes(lo) = 239 And bytes(lo + 1) = 187 _
                             And bytes(lo + 2) = 191 Then
                encoding = "utf-8"
            End If
        End If

        If bytes(lo) = 255 And bytes(lo + 1) = 254 Then
            encoding = "unicode"
        ElseIf bytes(lo) = 254 And bytes(lo + 1) = 255 Then
            encoding = "unicodeFFFE"
        End If
    End If
    raw.Close

    Set result = CreateObject("ADODB.Stream")
    result.Type = 2 ' adTypeText
    result.Charset = encoding
    result.Open
    result.LoadFromFile filePath
    Set OpenHulftText = result
End Function

Private Function PrepareOutputSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                     After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ' Only these four target sheets are cleared on successful parsing.
        ws.Cells.Clear
    End If

    Set PrepareOutputSheet = ws
End Function

Private Sub StyleOutputSheet(ByVal ws As Worksheet, _
                             ByVal lastRow As Long, _
                             ByVal lastCol As Long)
    Dim c As Long

    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .AutoFilter
    End With

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(35, 60, 92)
        .RowHeight = 23
    End With

    For c = 1 To lastCol
        ws.Columns(c).AutoFit
        If ws.Columns(c).ColumnWidth > 55 Then
            ws.Columns(c).ColumnWidth = 55
        End If
    Next c
End Sub

Private Function StripBom(ByVal value As String) As String
    If Len(value) > 0 Then
        If AscW(Left$(value, 1)) = -257 Or _
           AscW(Left$(value, 1)) = 65279 Then
            value = Mid$(value, 2)
        End If
    End If
    StripBom = value
End Function

Private Function SafeCellString(ByVal value As String) As String
    value = Replace$(value, Chr$(0), "")
    If Len(value) > 32767 Then value = Left$(value, 32767)
    SafeCellString = value
End Function

Private Function IsSecretParameter(ByVal name As String) As Boolean
    Select Case UCase$(name)
        Case "PASSWORD", "PW", "PASS", "PASSPHRASE", _
             "SECRET", "SECRETKEY", "ENCRYPTKEY", "ENCRYPTION_KEY"
            IsSecretParameter = True
        Case Else
            IsSecretParameter = False
    End Select
End Function
