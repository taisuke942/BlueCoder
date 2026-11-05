Option Explicit

' HULFT8 for Windows - Excel management list -> edited definition cards.
' Version 2: TEMPLATE-PRESERVING exporter, 2026-11-05
'
' REQUIRED: Keep the original utligen output TXT for each sheet.
' On export, select the ORIGINAL TXT. The program changes only Excel-edited
' fields; other lines, comments, inactive parameters, and ordering stay.
'
' Replace the PREVIOUS export VBA module with this ENTIRE module.
' Keep the separate import VBA module as-is.
' Run: HULFT_ExportCardsForImport
'
' Scope: only SNDFILE / RCVFILE cards. Histories are never imported into HULFT.
' Safety: no new/deleted/renamed FILEIDs, no overwrite of original/output files.
' Source text defaults to Japanese Windows CP932 (Shift-JIS). UTF-8 BOM and
' UTF-16LE BOM are detected. Output uses the same detected encoding.
' This tool has not been run against the user's actual HULFT environment.

Private Const SND_ALLOWED As String = _
    "FILENAME INTERVAL BLOCKLEN BLOCKCNT COMP COMPSIZE DEFLATELEVEL " & _
    "TRANSPRTY TRANSTYPE CODESET KJCHNGE SHIFTTRANSACT CLEAR PREJOBID " & _
    "JOBID EJOBID DBID MAILID GRPID FMTID PASSWORD COMMENT"

Private Const RCV_ALLOWED As String = _
    "FILENAME CODESET TRANSMODE ABNORMAL RCVTYPE GENCTL GENMNGNO " & _
    "JOBID EJOBID GRPID PASSWORD MAILID JOBWAIT DATAVERIFY COMMENT"

Private Const SND_REQUIRED As String = _
    "FILENAME COMP TRANSTYPE KJCHNGE GRPID"

Private Const RCV_REQUIRED As String = _
    "FILENAME TRANSMODE ABNORMAL RCVTYPE GENCTL"

Public Sub HULFT_ExportCardsForImport()
    Dim src(1 To 2) As String
    Dim dst(1 To 2) As String
    Dim encoding(1 To 2) As String
    Dim generated(1 To 2) As String
    Dim cardCount(1 To 2) As Long
    Dim editCount(1 To 2) As Long
    Dim selected(1 To 2) As Boolean
    Dim names(1 To 2) As String
    Dim types(1 To 2) As String
    Dim baseFolder As String, report As String
    Dim ws As Worksheet, rows As Object
    Dim idx As Long, total As Long

    On Error GoTo Fail

    names(1) = "配信管理"
    names(2) = "集信管理"
    types(1) = "SNDFILE"
    types(2) = "RCVFILE"

    For idx = 1 To 2
        Set ws = GetTargetSheet(names(idx))
        If Not ws Is Nothing Then
            selected(idx) = True
            total = total + 1

            Set rows = ReadExcelRecords(ws, types(idx))

            src(idx) = AskOriginalFile(names(idx))
            If Len(src(idx)) = 0 Then Exit Sub

            generated(idx) = PatchUsingOriginal( _
                src(idx), types(idx), rows, encoding(idx), _
                cardCount(idx), editCount(idx))
        End If
    Next idx

    If total = 0 Then
        MsgBox "配信管理・集信管理シートが見つかりません。" & vbCrLf & _
               "先に取込VBAで管理情報を読み込んでください。", vbExclamation
        Exit Sub
    End If

    baseFolder = AskOutputFolder()
    If Len(baseFolder) = 0 Then Exit Sub

    dst(1) = baseFolder & "\hulft_snd.txt"
    dst(2) = baseFolder & "\hulft_rcv.txt"

    ' Validate BOTH targets before writing either one.
    For idx = 1 To 2
        If selected(idx) Then
            If StrComp(dst(idx), src(idx), vbTextCompare) = 0 Then
                Err.Raise vbObjectError + 8201, , _
                    "元のTXTと出力先が同じです。別フォルダを選んでください。"
            End If
            If FileExists(dst(idx)) Then
                Err.Raise vbObjectError + 8202, , _
                    "出力先ファイルがすでに存在します: " & dst(idx) & vbCrLf & _
                    "元のTXTや以前の結果の上書き防止のため、" & _
                    "新しい空フォルダを選んでください。"
            End If
        End If
    Next idx

    If MsgBox("Excelの内容を元のutligenテキストへ反映します。" & vbCrLf & _
              "元のTXTは変更せず、別フォルダへ出力します。" & vbCrLf & vbCrLf & _
              "配信管理: " & cardCount(1) & "件 / 変更 " & editCount(1) & "項目" & vbCrLf & _
              "集信管理: " & cardCount(2) & "件 / 変更 " & editCount(2) & "項目" & vbCrLf & vbCrLf & _
              "出力を実行しますか？", vbYesNo + vbQuestion + vbDefaultButton2, _
              "HULFT定義カード出力") <> vbYes Then Exit Sub

    For idx = 1 To 2
        If selected(idx) Then
            SaveTextFile dst(idx), generated(idx), encoding(idx)
        End If
    Next idx

    report = "テキストの出力が完了しました。" & vbCrLf
    For idx = 1 To 2
        If selected(idx) Then
            report = report & vbCrLf & names(idx) & ": " & cardCount(idx) & _
                     "件 / Excel編集 " & editCount(idx) & "項目" & _
                     vbCrLf & dst(idx)
        End If
    Next idx
    MsgBox report & vbCrLf & vbCrLf & _
           "HULFTへの登録前に、元ファイルとの比較と" & _
           "テスト環境での検証を行ってください。", vbInformation
    Exit Sub

Fail:
    MsgBox "安全のため処理を中止しました。" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "※エラーが出力中に発生した場合は出力先も確認してください。", _
           vbExclamation, "HULFT定義カード出力"
End Sub

Private Function GetTargetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetTargetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function ReadExcelRecords(ByVal ws As Worksheet, _
                                  ByVal cardType As String) As Object
    Dim result As Object, row As Object, allowed As Object
    Dim seenHeaders As Object, found As Range
    Dim hdr() As String, v As String, fileID As String
    Dim colCount As Long, rowCount As Long
    Dim r As Long, c As Long, anyData As Boolean

    Set result = NewTextDictionary()
    Set allowed = GetAllowedParameters(cardType)
    Set seenHeaders = NewTextDictionary()

    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                             LookAt:=xlPart, SearchOrder:=xlByRows, _
                             SearchDirection:=xlPrevious, MatchCase:=False)
    If found Is Nothing Then
        Err.Raise vbObjectError + 8210, , ws.Name & "シートが空です。"
    End If
    rowCount = found.Row

    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                             LookAt:=xlPart, SearchOrder:=xlByColumns, _
                             SearchDirection:=xlPrevious, MatchCase:=False)
    colCount = found.Column

    If rowCount < 2 Or colCount < 2 Then
        Err.Raise vbObjectError + 8211, , _
            ws.Name & "シートにヘッダとデータが必要です。"
    End If

    ReDim hdr(1 To colCount)
    For c = 1 To colCount
        v = UCase$(Trim$(CellText(ws.Cells(1, c))))
        If Len(v) = 0 Then
            Err.Raise vbObjectError + 8212, , _
                ws.Name & ": 1行目の見出しが空欄です（" & c & "列）。"
        End If
        If seenHeaders.Exists(v) Then
            Err.Raise vbObjectError + 8213, , _
                ws.Name & ": 項目名が重複しています: " & v
        End If
        If v <> "FILEID" And Not allowed.Exists(v) Then
            Err.Raise vbObjectError + 8214, , _
                ws.Name & ": 不明な列名: " & v & vbCrLf & _
                "HULFTの管理情報項目のみ指定してください。"
        End If
        seenHeaders.Add v, True
        hdr(c) = v
    Next c

    If hdr(1) <> "FILEID" Then
        Err.Raise vbObjectError + 8215, , _
            ws.Name & "シートのA1はFILEIDにしてください。"
    End If

    For r = 2 To rowCount
        Set row = NewTextDictionary()
        anyData = False

        For c = 1 To colCount
            v = CellText(ws.Cells(r, c))
            If HasUnsafeLineBreak(v) Then
                Err.Raise vbObjectError + 8216, , _
                    ws.Name & "!" & ws.Cells(r, c).Address(False, False) & _
                    " に改行・NUL文字があります。"
            End If
            If Len(v) > 0 Then anyData = True
            row.Add hdr(c), v
        Next c

        fileID = CStr(row("FILEID"))
        If Len(fileID) = 0 Then
            If anyData Then
                Err.Raise vbObjectError + 8217, , _
                    ws.Name & "の" & r & "行目はFILEIDが空欄です。"
            End If
        Else
            If Trim$(fileID) <> fileID Or InStr(fileID, "=") > 0 _
               Or UCase$(fileID) = "_INSTANT_TRANSFER_" _
               Or LenB(StrConv(fileID, vbFromUnicode)) > 50 Then
                Err.Raise vbObjectError + 8218, , _
                    ws.Name & "の" & r & "行目のFILEIDが不正です: " & fileID
            End If
            If result.Exists(fileID) Then
                Err.Raise vbObjectError + 8219, , _
                    ws.Name & ": FILEIDが重複しています: " & fileID
            End If
            result.Add fileID, row
        End If
    Next r

    If result.Count = 0 Then
        Err.Raise vbObjectError + 8220, , _
            ws.Name & "に出力可能な管理情報がありません。"
    End If
    Set ReadExcelRecords = result
End Function

Private Function CellText(ByVal cell As Range) As String
    Dim raw As Variant
    If cell.HasFormula Then
        Err.Raise vbObjectError + 8221, , _
            cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " は数式です。値に変換してから出力してください。"
    End If
    raw = cell.Value2
    If IsError(raw) Then
        Err.Raise vbObjectError + 8222, , _
            cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " にExcelのエラー値があります。"
    End If
    If IsEmpty(raw) Then
        CellText = ""
    Else
        CellText = CStr(raw)
    End If
End Function

Private Function PatchUsingOriginal( _
    ByVal originalPath As String, ByVal cardType As String, _
    ByVal excelRecords As Object, ByRef encoding As String, _
    ByRef cardCount As Long, ByRef editCount As Long) As String

    Dim originalText As String, eol As String
    Dim sourceLines() As String, outputLines As Collection
    Dim row As Object, touched As Object, finalValues As Object
    Dim visitedIDs As Object, allowed As Object
    Dim currentID As String, name As String, oldValue As String
    Dim newValue As String, line As String, key As Variant
    Dim i As Long, rowNumber As Long, eq As Long
    Dim inside As Boolean, isField As Boolean, startTag As String
    Dim normalized As String, required As String

    originalText = LoadTextFile(originalPath, encoding)
    eol = DetectLineEnding(originalText)
    sourceLines = Split(originalText, eol)
    Set outputLines = New Collection
    Set visitedIDs = NewTextDictionary()
    Set allowed = GetAllowedParameters(cardType)

    If cardType = "SNDFILE" Then
        required = SND_REQUIRED
    Else
        required = RCV_REQUIRED
    End If

    For i = LBound(sourceLines) To UBound(sourceLines)
        line = sourceLines(i)
        normalized = Trim$(line)
        rowNumber = i + 1

        ' Only an active start tag begins a card. # comments are untouched.
        isField = ParseActiveField(line, name, oldValue)
        If isField Then
            If name = "SNDFILE" Or name = "RCVFILE" Then
                If inside Then
                    Err.Raise vbObjectError + 8240, , _
                        "元TXTのENDが不足しています。行=" & rowNumber
                End If
                If name <> cardType Then
                    Err.Raise vbObjectError + 8241, , _
                        "元TXTに異なる管理情報種別があります。行=" & rowNumber & _
                        " / 種別=" & name
                End If
                currentID = Trim$(oldValue)
                If Len(currentID) = 0 Then
                    Err.Raise vbObjectError + 8242, , _
                        "元TXTにFILEID空欄の定義があります。行=" & rowNumber
                End If
                If visitedIDs.Exists(currentID) Then
                    Err.Raise vbObjectError + 8243, , _
                        "元TXT内にFILEIDが重複しています: " & currentID
                End If
                If Not excelRecords.Exists(currentID) Then
                    Err.Raise vbObjectError + 8244, , _
                        "元TXTに存在しExcelに存在しないFILEIDです: " & currentID & vbCrLf & _
                        "削除・ID変更には対応しません。元データとExcelを照合してください。"
                End If

                Set row = excelRecords(currentID)
                Set touched = NewTextDictionary()
                Set finalValues = NewTextDictionary()
                visitedIDs.Add currentID, True
                cardCount = cardCount + 1
                inside = True
            End If
        End If

        If inside And UCase$(normalized) = "END" Then
            ' A new optional Excel column can be added before END, while
            ' leaving the original commented parameter in its old position.
            For Each key In row.Keys
                name = CStr(key)
                If name <> "FILEID" And Not touched.Exists(name) Then
                    newValue = CStr(row(name))
                    If UCase$(Trim$(newValue)) = "[MASKED]" Then
                        Err.Raise vbObjectError + 8245, , _
                            currentID & " / " & name & _
                            ": 元のTXTに該当する有効な項目がありません。" & _
                            vbCrLf & "[MASKED]をそのまま書き出せません。"
                    End If
                    If Len(newValue) > 0 Then
                        outputLines.Add name & "=" & newValue
                        finalValues(name) = newValue
                        editCount = editCount + 1
                    End If
                End If
            Next key

            CheckRequiredFields finalValues, required, currentID
            outputLines.Add line
            inside = False
            currentID = ""
            GoTo NextSourceLine
        End If

        If inside And isField Then
            If name <> cardType Then
                If Not allowed.Exists(name) Then
                    ' Unknown original parameters are preserved, not discarded.
                    ' They are not made editable by this macro.
                    outputLines.Add line
                    GoTo NextSourceLine
                End If

                If touched.Exists(name) Then
                    Err.Raise vbObjectError + 8246, , _
                        currentID & ": 元TXTで項目が重複しています: " & name
                End If
                touched.Add name, True

                If row.Exists(name) Then
                    newValue = CStr(row(name))
                    If UCase$(Trim$(newValue)) = "[MASKED]" Then
                        If name <> "PASSWORD" Then
                            Err.Raise vbObjectError + 8247, , _
                                currentID & " / " & name & _
                                ": [MASKED] はPASSWORD列以外では使用できません。"
                        End If
                        If UCase$(Trim$(oldValue)) = "%ERROR" Then
                            Err.Raise vbObjectError + 8248, , _
                                currentID & ": 元TXTにPASSWORD=%ERRORがあります。" & _
                                vbCrLf & "実際の暗号キーをExcel側で設定してください。"
                        End If
                        ' Preserve the original password line in place.
                        finalValues(name) = oldValue
                    ElseIf StrComp(Trim$(newValue), Trim$(oldValue), _
                                  vbBinaryCompare) = 0 Then
                        ' Unchanged: keep original spaces and full line.
                        finalValues(name) = oldValue
                    ElseIf Len(newValue) = 0 Then
                        ' Clearing an optional field comments it out.
                        If IsRequired(name, required) Then
                            Err.Raise vbObjectError + 8249, , _
                                currentID & ": 必須項目 " & name & " を空欄にできません。"
                        End If
                        line = "#" & line
                        editCount = editCount + 1
                    Else
                        eq = InStr(1, line, "=", vbBinaryCompare)
                        line = Left$(line, eq) & newValue
                        finalValues(name) = newValue
                        editCount = editCount + 1
                    End If
                Else
                    finalValues(name) = oldValue
                End If
            End If
        End If

        outputLines.Add line

NextSourceLine:
    Next i

    If inside Then
        Err.Raise vbObjectError + 8250, , _
            "元TXTの最終定義にEND行がありません。"
    End If
    If cardCount = 0 Then
        Err.Raise vbObjectError + 8251, , _
            "元TXTに" & cardType & "の定義カードが見つかりません。"
    End If
    If visitedIDs.Count <> excelRecords.Count Then
        For Each key In excelRecords.Keys
            If Not visitedIDs.Exists(CStr(key)) Then
                Err.Raise vbObjectError + 8252, , _
                    "Excelには存在するが元TXTに存在しないFILEIDです: " & CStr(key) & _
                    vbCrLf & "新規IDの追加・FILEID変更には対応しません。"
            End If
        Next key
    End If

    PatchUsingOriginal = JoinLines(outputLines, eol)
End Function

Private Sub CheckRequiredFields( _
    ByVal actual As Object, ByVal required As String, _
    ByVal fileID As String)
    Dim field As Variant, name As String
    For Each field In Split(required, " ")
        name = CStr(field)
        If Not actual.Exists(name) Then
            Err.Raise vbObjectError + 8260, , _
                fileID & ": 必須項目 " & name & " がありません。"
        End If
        If Len(Trim$(CStr(actual(name)))) = 0 Then
            Err.Raise vbObjectError + 8261, , _
                fileID & ": 必須項目 " & name & " が空欄です。"
        End If
    Next field
End Sub

Private Function IsRequired(ByVal name As String, _
                            ByVal required As String) As Boolean
    IsRequired = (InStr(1, " " & required & " ", _
                         " " & name & " ", vbTextCompare) > 0)
End Function

Private Function ParseActiveField(ByVal line As String, _
                                  ByRef name As String, _
                                  ByRef value As String) As Boolean
    Dim s As String, p As Long
    name = ""
    value = ""
    s = Trim$(line)
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "#" Then Exit Function
    p = InStr(1, s, "=", vbBinaryCompare)
    If p < 2 Then Exit Function
    name = UCase$(Trim$(Left$(s, p - 1)))
    value = Mid$(s, p + 1)
    ParseActiveField = True
End Function

Private Function GetAllowedParameters(ByVal cardType As String) As Object
    Dim d As Object, item As Variant, allNames As String
    Set d = NewTextDictionary()
    If cardType = "SNDFILE" Then
        allNames = SND_ALLOWED
    Else
        allNames = RCV_ALLOWED
    End If
    For Each item In Split(allNames, " ")
        d.Add CStr(item), True
    Next item
    Set GetAllowedParameters = d
End Function

Private Function NewTextDictionary() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set NewTextDictionary = d
End Function

Private Function HasUnsafeLineBreak(ByVal s As String) As Boolean
    HasUnsafeLineBreak = (InStr(s, vbCr) > 0 Or _
                          InStr(s, vbLf) > 0 Or InStr(s, Chr$(0)) > 0)
End Function

Private Function AskOriginalFile(ByVal sheetName As String) As String
    Dim picked As Variant
    picked = Application.GetOpenFilename( _
        "Text files (*.txt;*.prm;*.dat),*.txt;*.prm;*.dat,All files (*.*),*.*", _
        1, "元のutligen出力TXTを選択: " & sheetName, , False)
    If VarType(picked) <> vbBoolean Then AskOriginalFile = CStr(picked)
End Function

Private Function AskOutputFolder() As String
    Dim d As Object
    Set d = Application.FileDialog(4) ' Folder picker.
    With d
        .Title = "元のTXTとは別の空フォルダを指定"
        .AllowMultiSelect = False
        If .Show = -1 Then AskOutputFolder = .SelectedItems(1)
    End With
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function

Private Function LoadTextFile(ByVal filePath As String, _
                              ByRef charset As String) As String
    Dim raw As Object, txt As Object
    Dim bytes As Variant, pos As Long
    Dim s As String

    charset = "shift_jis"
    Set raw = CreateObject("ADODB.Stream")
    raw.Type = 1
    raw.Open
    raw.LoadFromFile filePath

    If raw.Size = 0 Then
        raw.Close
        Err.Raise vbObjectError + 8270, , "元のTXTが空です: " & filePath
    End If
    If raw.Size >= 2 Then
        bytes = raw.Read(IIf(raw.Size >= 3, 3, 2))
        pos = LBound(bytes)
        If bytes(pos) = 255 And bytes(pos + 1) = 254 Then
            charset = "unicode"
        ElseIf bytes(pos) = 254 And bytes(pos + 1) = 255 Then
            raw.Close
            Err.Raise vbObjectError + 8271, , _
                "UTF-16BEは未対応です。元TXTをShift-JISで取得してください。"
        ElseIf UBound(bytes) >= pos + 2 Then
            If bytes(pos) = 239 And bytes(pos + 1) = 187 _
                                 And bytes(pos + 2) = 191 Then
                charset = "utf-8"
            End If
        End If
    End If
    raw.Close

    Set txt = CreateObject("ADODB.Stream")
    txt.Type = 2
    txt.Charset = charset
    txt.Open
    txt.LoadFromFile filePath
    s = txt.ReadText(-1)
    txt.Close

    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = -257 Then s = Mid$(s, 2)
    End If
    LoadTextFile = s
End Function

Private Sub SaveTextFile(ByVal filePath As String, _
                         ByVal textData As String, ByVal charset As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = charset
    stream.Open
    stream.WriteText textData
    stream.SaveToFile filePath, 1 ' Create new; fail if it exists.
    stream.Close
End Sub

Private Function DetectLineEnding(ByVal textData As String) As String
    If InStr(1, textData, vbCrLf, vbBinaryCompare) > 0 Then
        DetectLineEnding = vbCrLf
    ElseIf InStr(1, textData, vbLf, vbBinaryCompare) > 0 Then
        DetectLineEnding = vbLf
    ElseIf InStr(1, textData, vbCr, vbBinaryCompare) > 0 Then
        DetectLineEnding = vbCr
    Else
        Err.Raise vbObjectError + 8280, , _
            "元TXTに改行がありません。utligen出力を選択してください。"
    End If
End Function

Private Function JoinLines(ByVal entries As Collection, _
                           ByVal eol As String) As String
    Dim a() As String, i As Long
    ReDim a(0 To entries.Count - 1)
    For i = 1 To entries.Count
        a(i - 1) = CStr(entries(i))
    Next i
    JoinLines = Join(a, eol)
End Function
