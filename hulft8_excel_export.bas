Option Explicit

' HULFT8 Windows: Excel 配信管理 / 集信管理 -> utliupdt definition cards.
' Add this as a NEW standard module to the workbook that contains the import VBA.
' Run HULFT_ExportCardsForImport. The export does not execute HULFT.
' Output is Shift-JIS (CP932) / CRLF, one parameter per line.
' Only the 2 management-information sheets are exported. Histories are not.

Private Const SND_COLUMNS As String = _
    "FILENAME INTERVAL BLOCKLEN BLOCKCNT COMP COMPSIZE DEFLATELEVEL TRANSPRTY " & _
    "TRANSTYPE CODESET KJCHNGE SHIFTTRANSACT CLEAR PREJOBID JOBID EJOBID " & _
    "DBID MAILID GRPID FMTID PASSWORD COMMENT"

Private Const RCV_COLUMNS As String = _
    "FILENAME CODESET TRANSMODE ABNORMAL RCVTYPE GENCTL GENMNGNO JOBID " & _
    "EJOBID GRPID PASSWORD MAILID JOBWAIT DATAVERIFY COMMENT"

Private Const SND_REQUIRED As String = _
    "FILENAME COMP TRANSTYPE KJCHNGE GRPID"

Private Const RCV_REQUIRED As String = _
    "FILENAME TRANSMODE ABNORMAL RCVTYPE GENCTL"

Public Sub HULFT_ExportCardsForImport()
    Dim sndSheet As Worksheet, rcvSheet As Worksheet
    Dim sndLines As Collection, rcvLines As Collection
    Dim sndCount As Long, rcvCount As Long
    Dim folder As String, sndPath As String, rcvPath As String
    Dim collisions As String, outText As String

    On Error GoTo ExportError

    Set sndSheet = FindHulftSheet("配信管理")
    Set rcvSheet = FindHulftSheet("集信管理")

    If sndSheet Is Nothing And rcvSheet Is Nothing Then
        MsgBox "配信管理・集信管理シートが見つかりません。" & vbCrLf & _
               "先に取込VBAで管理情報を読み込んでください。", vbExclamation
        Exit Sub
    End If

    ' Validate all data before asking for an output location or writing files.
    If Not sndSheet Is Nothing Then
        Set sndLines = CompileHulftCards(sndSheet, "SNDFILE", _
                                          SND_COLUMNS, SND_REQUIRED, sndCount)
    End If
    If Not rcvSheet Is Nothing Then
        Set rcvLines = CompileHulftCards(rcvSheet, "RCVFILE", _
                                          RCV_COLUMNS, RCV_REQUIRED, rcvCount)
    End If

    folder = ChooseExportFolder()
    If Len(folder) = 0 Then Exit Sub

    sndPath = folder & "\hulft_snd.txt"
    rcvPath = folder & "\hulft_rcv.txt"

    If Not sndLines Is Nothing Then
        If PathAlreadyExists(sndPath) Then collisions = collisions & sndPath & vbCrLf
    End If
    If Not rcvLines Is Nothing Then
        If PathAlreadyExists(rcvPath) Then collisions = collisions & rcvPath & vbCrLf
    End If

    If Len(collisions) > 0 Then
        If MsgBox("次のファイルがすでに存在します。" & vbCrLf & vbCrLf & _
                  collisions & vbCrLf & _
                  "元のutligenエクスポートを上書きする危険があります。" & _
                  vbCrLf & "バックアップを確認のうえ上書きしますか？", _
                  vbYesNo + vbExclamation + vbDefaultButton2, _
                  "既存ファイルの上書き確認") <> vbYes Then Exit Sub
    End If

    If Not sndLines Is Nothing Then WriteShiftJisFile sndPath, sndLines
    If Not rcvLines Is Nothing Then WriteShiftJisFile rcvPath, rcvLines

    outText = "定義カードを出力しました。" & vbCrLf & vbCrLf
    If Not sndLines Is Nothing Then
        outText = outText & "配信管理: " & sndCount & "件" & vbCrLf & sndPath & vbCrLf
    End If
    If Not rcvLines Is Nothing Then
        outText = outText & "集信管理: " & rcvCount & "件" & vbCrLf & rcvPath & vbCrLf
    End If
    outText = outText & vbCrLf & _
              "内容を確認してから、移行先のテスト用HULFTでutliupdtを実行してください。"
    MsgBox outText, vbInformation, "HULFT定義カード出力"
    Exit Sub

ExportError:
    MsgBox "出力を中止しました。" & vbCrLf & Err.Description & vbCrLf & _
           "途中でエラーになった場合、出力先のファイルを確認してください。", _
           vbCritical, "HULFT定義カード出力"
End Sub

Private Function FindHulftSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set FindHulftSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Private Function CompileHulftCards( _
    ByVal ws As Worksheet, ByVal cardType As String, _
    ByVal allowedList As String, ByVal requiredList As String, _
    ByRef cardCount As Long) As Collection

    Dim lastCol As Long, lastRow As Long
    Dim headers() As String, c As Long, r As Long
    Dim allowed As Object, seenHeaders As Object, seenIDs As Object
    Dim rowValues As Object, lines As Collection
    Dim tag As Variant, value As String, id As String, k As String
    Dim hasAnyValue As Boolean
    Dim tagList As Variant

    Set allowed = NewKeywordDictionary(allowedList)
    Set seenHeaders = CreateObject("Scripting.Dictionary")
    seenHeaders.CompareMode = vbTextCompare
    Set seenIDs = CreateObject("Scripting.Dictionary")
    seenIDs.CompareMode = vbTextCompare
    Set lines = New Collection

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        Err.Raise vbObjectError + 9230, , _
            ws.Name & "シートにデータ行がありません。"
    End If
    If lastCol < 2 Then
        Err.Raise vbObjectError + 9231, , _
            ws.Name & "シートの項目列が不足しています。"
    End If

    ReDim headers(1 To lastCol)

    For c = 1 To lastCol
        k = UCase$(Trim$(ReadCellText(ws.Cells(1, c), ws.Name, 1, c)))
        If Len(k) = 0 Then
            Err.Raise vbObjectError + 9232, , _
                ws.Name & "シートの1行目に空欄の見出しがあります。列=" & c
        End If
        If seenHeaders.Exists(k) Then
            Err.Raise vbObjectError + 9233, , _
                ws.Name & "シートで見出しが重複しています: " & k
        End If
        If k <> "FILEID" And Not allowed.Exists(k) Then
            Err.Raise vbObjectError + 9234, , _
                ws.Name & "シートの列「" & k & "」はHULFTの定義カード項目ではありません。" & _
                vbCrLf & "この列は削除するか、対応項目名にしてください。"
        End If
        seenHeaders.Add k, True
        headers(c) = k
    Next c

    If headers(1) <> "FILEID" Then
        Err.Raise vbObjectError + 9235, , _
            ws.Name & "シートのA1はFILEIDである必要があります。"
    End If

    tagList = Split(allowedList, " ")
    cardCount = 0

    For r = 2 To lastRow
        Set rowValues = CreateObject("Scripting.Dictionary")
        rowValues.CompareMode = vbTextCompare
        hasAnyValue = False

        For c = 1 To lastCol
            k = headers(c)
            value = ReadCellText(ws.Cells(r, c), ws.Name, r, c)

            If InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 _
               Or InStr(value, Chr$(0)) > 0 Then
                Err.Raise vbObjectError + 9236, , _
                    ws.Name & "!" & ws.Cells(r, c).Address(False, False) & _
                    " に改行またはNUL文字があります。"
            End If

            If Len(Trim$(value)) > 0 Then hasAnyValue = True

            If k = "PASSWORD" Then
                If UCase$(Trim$(value)) = "[MASKED]" Or _
                   UCase$(Trim$(value)) = "%ERROR" Then
                    Err.Raise vbObjectError + 9237, , _
                        ws.Name & "!" & ws.Cells(r, c).Address(False, False) & _
                        " の暗号キーは復元できません。" & vbCrLf & _
                        "実際の暗号キーを設定するか、元の定義カードを参照してください。"
                End If
            End If

            rowValues.Add k, value
        Next c

        id = CStr(rowValues("FILEID"))
        If Len(Trim$(id)) = 0 Then
            If hasAnyValue Then
                Err.Raise vbObjectError + 9238, , _
                    ws.Name & "シート " & r & "行目のFILEIDが空欄です。"
            End If
            GoTo NextDataRow
        End If

        If Trim$(id) <> id Or InStr(id, "=") > 0 Or _
            UCase$(id) = "_INSTANT_TRANSFER_" Then
            Err.Raise vbObjectError + 9239, , _
                ws.Name & "シート " & r & "行目のFILEIDが不正です: " & id
        End If

        If seenIDs.Exists(id) Then
            Err.Raise vbObjectError + 9240, , _
                ws.Name & "シートに重複FILEIDがあります: " & id
        End If
        seenIDs.Add id, True

        For Each tag In Split(requiredList, " ")
            k = CStr(tag)
            If Not rowValues.Exists(k) Then
                Err.Raise vbObjectError + 9241, , _
                    ws.Name & "シートの必須項目 " & k & " がありません。"
            End If
            If Len(Trim$(CStr(rowValues(k)))) = 0 Then
                Err.Raise vbObjectError + 9242, , _
                    ws.Name & "シート " & r & "行目の必須項目 " & k & " が空欄です。"
            End If
        Next tag

        lines.Add cardType & "=" & id

        ' Use the official order, even if the user reorders Excel columns.
        For Each tag In tagList
            k = CStr(tag)
            If rowValues.Exists(k) Then
                value = CStr(rowValues(k))
                If Len(value) > 0 Then lines.Add k & "=" & value
            End If
        Next tag

        lines.Add "END"
        cardCount = cardCount + 1

NextDataRow:
    Next r

    If cardCount = 0 Then
        Err.Raise vbObjectError + 9243, , _
            ws.Name & "シートに出力可能な管理情報がありません。"
    End If

    Set CompileHulftCards = lines
End Function

Private Function NewKeywordDictionary(ByVal words As String) As Object
    Dim dict As Object, item As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    For Each item In Split(words, " ")
        dict.Add CStr(item), True
    Next item
    Set NewKeywordDictionary = dict
End Function

Private Function ReadCellText( _
    ByVal cell As Range, ByVal sheetName As String, _
    ByVal row As Long, ByVal col As Long) As String

    If cell.HasFormula Then
        Err.Raise vbObjectError + 9250, , _
            sheetName & "!" & cell.Address(False, False) & _
            " は数式です。値に変換してから出力してください。"
    End If

    If IsError(cell.Value2) Then
        Err.Raise vbObjectError + 9251, , _
            sheetName & "!" & cell.Address(False, False) & _
            " がExcelエラー値になっています。"
    End If

    If IsEmpty(cell.Value2) Then
        ReadCellText = ""
    Else
        ReadCellText = CStr(cell.Value2)
    End If
End Function

Private Function ChooseExportFolder() As String
    Dim dlg As Object
    Set dlg = Application.FileDialog(4) ' msoFileDialogFolderPicker
    With dlg
        .Title = "HULFT定義カードTXTの出力先フォルダを選択"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        ChooseExportFolder = .SelectedItems(1)
    End With
End Function

Private Function PathAlreadyExists(ByVal path As String) As Boolean
    PathAlreadyExists = (Len(Dir$(path, vbNormal Or vbReadOnly Or vbHidden)) > 0)
End Function

Private Sub WriteShiftJisFile(ByVal path As String, ByVal lines As Collection)
    Dim stream As Object
    Dim allLines() As String
    Dim n As Long

    ReDim allLines(0 To lines.Count - 1)
    For n = 1 To lines.Count
        allLines(n - 1) = CStr(lines(n))
    Next n

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "shift_jis"
    stream.Open
    stream.WriteText Join(allLines, vbCrLf) & vbCrLf
    stream.SaveToFile path, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub
