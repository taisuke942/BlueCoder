Option Explicit

' Version 5: read named process-environment settings; preserve Unicode names.
' Other keys retain the version 3 standalone-token behavior.

' Usage: cscript //nologo set_ini_test.vbs --env | --check
' Legacy 2, 3, and 5 positional arguments are also supported.
' charset: utf-8, shift_jis, or unicode (UTF-16 LE with BOM).
' This source is ASCII-only. Section names can contain Unicode escapes: \uXXXX.
' Only standalone key= tokens in active lines of the requested section are edited.

Const adTypeBinary = 1
Const adTypeText = 2
Const adSaveCreateNotExist = 1
Const adSaveCreateOverWrite = 2

Dim fso, tempPath, stage, errorNumber, errorText
tempPath = ""
stage = "initialization"

WScript.Echo "INI updater v5 (environment / ODBC Database tail)"

On Error Resume Next
Main
errorNumber = Err.Number
errorText = Err.Description
On Error GoTo 0

If errorNumber <> 0 Then
    WScript.Echo "ERROR during " & stage & ": " & errorText & " (0x" & Hex(errorNumber) & ")"
    On Error Resume Next
    If tempPath <> "" Then fso.DeleteFile tempPath, True
    On Error GoTo 0
    WScript.Quit 1
End If
WScript.Quit 0

Sub Main()
    Dim path, newValue, charset, raw, original, updated, roundTrip, outputBytes
    Dim hasBom, targetSection, targetKey, position, ending, line, trimmed, closePos, inSection
    Dim matched, changed, fileInfo, checkBytes, foundSection
    Dim shell, env, mode, dryRun, expected, countText, expectedText, i

    stage = "reading settings"
    Set fso = CreateObject("Scripting.FileSystemObject")
    dryRun = False
    expected = 0
    mode = ""
    If WScript.Arguments.Count = 1 Then mode = LCase(WScript.Arguments(0))
    If mode = "--env" Or mode = "--check" Then
        Set shell = CreateObject("WScript.Shell")
        Set env = shell.Environment("PROCESS")
        path = env.Item("INI_FILE")
        newValue = env.Item("NEW_VALUE")
        charset = LCase(Trim(env.Item("INI_CHARSET")))
        targetSection = DecodeUnicodeEscapes(Trim(env.Item("TARGET_SECTION")))
        targetKey = Trim(env.Item("TARGET_KEY"))
        expectedText = Trim(env.Item("EXPECTED_MATCHES"))
        dryRun = (mode = "--check")
        If charset = "" Then charset = "shift_jis"
        If expectedText <> "" Then
            If Len(expectedText) > 6 Then Fail "E015: EXPECTED_MATCHES is too large."
            For i = 1 To Len(expectedText)
                If InStr("0123456789", Mid(expectedText, i, 1)) = 0 Then Fail "E015: EXPECTED_MATCHES must be a nonnegative integer."
            Next
            expected = CLng(expectedText)
        End If
    Else
        If WScript.Arguments.Count <> 2 And WScript.Arguments.Count <> 3 And WScript.Arguments.Count <> 5 Then
            Fail "E001: Run the v5 BAT, or use --env / --check."
        End If
        path = WScript.Arguments(0)
        newValue = WScript.Arguments(1)
        charset = "utf-8"
        If WScript.Arguments.Count >= 3 Then charset = LCase(Trim(WScript.Arguments(2)))
        If WScript.Arguments.Count = 5 Then
            targetSection = DecodeUnicodeEscapes(Trim(WScript.Arguments(3)))
            targetKey = Trim(WScript.Arguments(4))
        Else
            targetSection = ChrW(&H8A2D) & ChrW(&H5B9A) & "1"
            targetKey = "test"
        End If
    End If
    If Len(path) = 0 Then Fail "E011: INI_FILE is empty. Run the v5 BAT."
    If Len(newValue) = 0 Then Fail "E012: NEW_VALUE is empty."
    If Len(targetSection) = 0 Then Fail "E013: TARGET_SECTION is empty."
    If InStr(targetSection, "[") > 0 Or InStr(targetSection, "]") > 0 Then Fail "E013: Do not include brackets in TARGET_SECTION."
    ValidateKey targetKey
    path = fso.GetAbsolutePathName(path)
    Select Case charset
        Case "utf8": charset = "utf-8"
        Case "sjis", "cp932", "shift-jis": charset = "shift_jis"
        Case "utf-16", "utf-16le": charset = "unicode"
        Case "utf-8", "shift_jis", "unicode"
        Case Else: Fail "Supported charsets: utf-8, shift_jis, unicode."
    End Select

    ValidateValue newValue
    If StrComp(targetKey, "Database", vbTextCompare) = 0 Then
        If InStr(newValue, "{") > 0 Or InStr(newValue, "}") > 0 Or InStr(newValue, "'") > 0 Then
            Fail "E003: Database name must not contain braces or quotes."
        End If
    End If
    If Not fso.FileExists(path) Then Fail "E101: INI file not found: " & path
    Set fileInfo = fso.GetFile(path)
    If fileInfo.Size = 0 Then Fail "E104: The INI file is empty."
    If (fileInfo.Attributes And 1) <> 0 Then Fail "E103: The INI file is read-only."

    stage = "reading the INI"
    raw = ReadBytes(path)
    stage = "checking the file BOM"
    hasBom = CheckBom(raw, charset)
    stage = "decoding the INI"
    original = DecodeBytes(raw, charset)
    If InStr(original, ChrW(0)) > 0 Then Fail "E102: NUL character found; check INI_CHARSET."

    ' Refuse a charset mismatch or a conversion that would alter other bytes.
    stage = "checking the original encoding"
    roundTrip = EncodeText(original, charset, hasBom)
    If Not SameBytes(raw, roundTrip) Then
        Fail "E102: The file cannot be preserved with INI_CHARSET=" & charset & ". Check its encoding."
    End If

    stage = "parsing the INI"
    foundSection = False
    inSection = False
    matched = 0
    changed = 0
    updated = ""

    ' Keep every original line ending, including a missing final newline.
    position = 1
    Do While position <= Len(original)
        SplitNextLine original, position, line, ending
        trimmed = Trim(Replace(line, vbTab, " "))
        If Len(trimmed) > 0 Then
            If Left(trimmed, 1) <> ";" And Left(trimmed, 1) <> "#" Then
                If Left(trimmed, 1) = "[" Then
                    inSection = False
                    closePos = InStr(2, trimmed, "]")
                    If closePos > 0 Then
                        inSection = (Trim(Mid(trimmed, 2, closePos - 2)) = targetSection)
                        If inSection Then foundSection = True
                    End If
                ElseIf inSection Then
                    line = RewriteLine(line, targetKey, newValue, matched, changed)
                End If
            End If
        End If
        updated = updated & line & ending
    Loop

    If Not foundSection Then Fail "E201: Target section not found. No file was changed."
    If matched = 0 Then Fail "E202: No matching active value. In Database mode, Database must be the final attribute. No file was changed."
    If expected > 0 Then
        If matched <> expected Then Fail "E204: Expected " & expected & " matching values, found " & matched & ". No file was changed."
    End If
    WScript.Echo "Matched: " & matched & " value(s)."
    If dryRun Then
        WScript.Echo "Would update: " & changed & " value(s)."
        WScript.Echo "CHECK ONLY: no files written."
        Exit Sub
    End If
    If changed = 0 Then
        WScript.Echo "No change needed. " & matched & " value(s) already match."
        Exit Sub
    End If

    stage = "preparing the updated file"
    outputBytes = EncodeText(updated, charset, hasBom)
    If DecodeBytes(outputBytes, charset) <> updated Then
        Fail "E301: The new value cannot be represented in INI_CHARSET=" & charset & "."
    End If
    tempPath = fso.BuildPath(fso.GetParentFolderName(path), "ini_edit_" & fso.GetTempName)
    WriteBytes tempPath, outputBytes, adSaveCreateNotExist
    checkBytes = ReadBytes(tempPath)
    If Not SameBytes(outputBytes, checkBytes) Then Fail "E304: Temporary file verification failed."

    checkBytes = ReadBytes(path)
    If Not SameBytes(raw, checkBytes) Then Fail "E302: The INI changed while this script was running. Update cancelled."

    ' Keep the existing pathname and file permissions by writing to that file.
    ' No backup is created. The temporary file contains new data, not a backup.
    stage = "writing the INI"
    WriteBytes path, outputBytes, adSaveCreateOverWrite
    checkBytes = ReadBytes(path)
    If Not SameBytes(outputBytes, checkBytes) Then Fail "E303: Updated file verification failed."
    fso.DeleteFile tempPath, True
    tempPath = ""

    WScript.Echo "Updated: " & changed & " value(s)."
    WScript.Echo "File: " & path
End Sub


' Convert \uXXXX notation (e.g. \u30BB\u30C3\u30C8 = a Japanese section name).
' Plain names such as MAIN or SETTING work without escapes.
Function DecodeUnicodeEscapes(ByVal value)
    Dim i, j, digits, result, ch
    i = 1
    result = ""
    Do While i <= Len(value)
        If Mid(value, i, 2) = "\u" Then
            If i + 5 > Len(value) Then Fail "E004: Invalid Unicode escape in TARGET_SECTION. Use \uXXXX."
            digits = Mid(value, i + 2, 4)
            For j = 1 To 4
                ch = UCase(Mid(digits, j, 1))
                If InStr(1, "0123456789ABCDEF", ch, vbBinaryCompare) = 0 Then
                    Fail "E004: Invalid Unicode escape in TARGET_SECTION. Use \uXXXX."
                End If
            Next
            result = result & ChrW(CLng("&H" & digits))
            i = i + 6
        Else
            result = result & Mid(value, i, 1)
            i = i + 1
        End If
    Loop
    DecodeUnicodeEscapes = result
End Function

Sub ValidateKey(ByVal keyName)
    Dim i, ch, charCode
    If Len(keyName) = 0 Then Fail "E014: TARGET_KEY is empty. Set it to Database in the BAT."
    For i = 1 To Len(keyName)
        ch = Mid(keyName, i, 1)
        charCode = AscW(ch)
        If (charCode >= 0 And charCode <= 32) Or ch = "=" Or ch = ";" Or ch = "#" Or ch = """" Then
            Fail "E014: TARGET_KEY must not contain spaces, =, ;, #, or quotes."
        End If
    Next
End Sub

Sub ValidateValue(ByVal value)
    Dim i, ch
    If Len(value) = 0 Then Fail "E012: NEW_VALUE must not be empty."
    For i = 1 To Len(value)
        ch = Mid(value, i, 1)
        If AscW(ch) >= 0 And AscW(ch) <= 32 Then Fail "E012: NEW_VALUE must not contain spaces or control characters."
        If ch = ";" Or ch = "#" Or ch = """" Then Fail "E012: NEW_VALUE must not contain semicolons, #, or double quotes."
    Next
End Sub

Function RewriteLine(ByVal line, ByVal targetKey, ByVal newValue, ByRef matched, ByRef changed)
    Dim i, n, ch, boundary, eqPos, valueStart, valueEnd, quoteEnd, copied, result
    If StrComp(targetKey, "Database", vbTextCompare) = 0 Then
        RewriteLine = RewriteDatabaseTail(line, newValue, matched, changed)
        Exit Function
    End If
    i = 1
    n = Len(line)
    copied = 1
    result = ""
    Do While i <= n
        ch = Mid(line, i, 1)
        If ch = ";" Or ch = "#" Then Exit Do
        If ch = """" Then
            quoteEnd = InStr(i + 1, line, """")
            If quoteEnd = 0 Then Exit Do
            i = quoteEnd + 1
        Else
            boundary = (i = 1)
            If i > 1 Then boundary = IsBlank(Mid(line, i - 1, 1))
            If boundary And StrComp(Mid(line, i, Len(targetKey)), targetKey, vbTextCompare) = 0 Then
                eqPos = i + Len(targetKey)
                Do While IsBlank(Mid(line, eqPos, 1))
                    eqPos = eqPos + 1
                Loop
                If Mid(line, eqPos, 1) = "=" Then
                    valueStart = eqPos + 1
                    Do While IsBlank(Mid(line, valueStart, 1))
                        valueStart = valueStart + 1
                    Loop
                    If Mid(line, valueStart, 1) = """" Then
                        valueStart = valueStart + 1
                        valueEnd = InStr(valueStart, line, """")
                        If valueEnd = 0 Then Fail "Unclosed quote in a " & targetKey & "= value."
                        If Mid(line, valueEnd + 1, 1) = """" Then Fail "Escaped quotes in a " & targetKey & "= value are not supported."
                        i = valueEnd + 1
                    Else
                        valueEnd = valueStart
                        Do While valueEnd <= n
                            ch = Mid(line, valueEnd, 1)
                            If IsBlank(ch) Or ch = ";" Or ch = "#" Then Exit Do
                            valueEnd = valueEnd + 1
                        Loop
                        If valueStart > eqPos + 1 Then
                            If InStr(Mid(line, valueStart, valueEnd - valueStart), "=") > 0 Then
                                Fail "Ambiguous empty " & targetKey & "= value followed by another setting. Update cancelled."
                            End If
                        End If
                        i = valueEnd
                    End If
                    matched = matched + 1
                    If Mid(line, valueStart, valueEnd - valueStart) <> newValue Then changed = changed + 1
                    result = result & Mid(line, copied, valueStart - copied) & newValue
                    copied = valueEnd
                Else
                    i = i + 1
                End If
            Else
                i = i + 1
            End If
        End If
    Loop
    RewriteLine = result & Mid(line, copied)
End Function

' Update only the final non-empty Database attribute in an active ODBC line.
' Do not treat semicolons inside braced or quoted values as separators.
Function RewriteDatabaseTail(ByVal line, ByVal newValue, ByRef matched, ByRef changed)
    Dim i, n, ch, closing, atValueStart, partStart, lastStart, lastEnd
    Dim eqPos, valueStart, valueEnd, keyName, trimmed, wrapper
    RewriteDatabaseTail = line
    trimmed = Trim(Replace(line, vbTab, " "))
    If Len(trimmed) = 0 Then Exit Function
    If Left(trimmed, 1) = ";" Or Left(trimmed, 1) = "#" Then Exit Function

    n = Len(line)
    i = 1
    partStart = 1
    lastStart = 0
    lastEnd = 0
    closing = ""
    atValueStart = False
    Do While i <= n
        ch = Mid(line, i, 1)
        If closing <> "" Then
            If ch = closing Then
                If Mid(line, i + 1, 1) = closing Then
                    i = i + 1
                Else
                    closing = ""
                End If
            End If
        ElseIf ch = ";" Then
            If Len(Trim(Replace(Mid(line, partStart, i - partStart), vbTab, " "))) > 0 Then
                lastStart = partStart
                lastEnd = i
            End If
            partStart = i + 1
            atValueStart = False
        ElseIf ch = "=" Then
            atValueStart = True
        ElseIf atValueStart And (ch = "{" Or ch = """" Or ch = "'") Then
            If ch = "{" Then
                closing = "}"
            Else
                closing = ch
            End If
            atValueStart = False
        ElseIf Not IsBlank(ch) Then
            atValueStart = False
        End If
        i = i + 1
    Loop
    If closing <> "" Then Fail "E203: Unclosed brace or quote in a connection line. No file was changed."

    If Len(Trim(Replace(Mid(line, partStart), vbTab, " "))) > 0 Then
        lastStart = partStart
        lastEnd = n + 1
    End If
    If lastStart = 0 Then Exit Function
    eqPos = InStr(lastStart, line, "=")
    If eqPos = 0 Or eqPos >= lastEnd Then Exit Function
    keyName = Trim(Replace(Mid(line, lastStart, eqPos - lastStart), vbTab, " "))
    If StrComp(keyName, "Database", vbTextCompare) <> 0 Then Exit Function

    valueStart = eqPos + 1
    Do While valueStart < lastEnd
        If Not IsBlank(Mid(line, valueStart, 1)) Then Exit Do
        valueStart = valueStart + 1
    Loop
    valueEnd = lastEnd
    Do While valueEnd > valueStart
        If Not IsBlank(Mid(line, valueEnd - 1, 1)) Then Exit Do
        valueEnd = valueEnd - 1
    Loop
    If valueStart < valueEnd Then
        ch = Mid(line, valueStart, 1)
        wrapper = ""
        If ch = "{" Then wrapper = "}"
        If ch = """" Or ch = "'" Then wrapper = ch
        If wrapper <> "" Then
            If Mid(line, valueEnd - 1, 1) <> wrapper Then Fail "E203: Invalid Database value wrapper. No file was changed."
            valueStart = valueStart + 1
            valueEnd = valueEnd - 1
        End If
    End If

    matched = matched + 1
    If Mid(line, valueStart, valueEnd - valueStart) <> newValue Then changed = changed + 1
    RewriteDatabaseTail = Left(line, valueStart - 1) & newValue & Mid(line, valueEnd)
End Function

Sub SplitNextLine(ByRef text, ByRef position, ByRef line, ByRef ending)
    Dim stopPos, ch
    stopPos = position
    Do While stopPos <= Len(text)
        ch = Mid(text, stopPos, 1)
        If ch = vbCr Or ch = vbLf Then Exit Do
        stopPos = stopPos + 1
    Loop
    line = Mid(text, position, stopPos - position)
    ending = Mid(text, stopPos, 1)
    If ending = vbCr Then
        If Mid(text, stopPos + 1, 1) = vbLf Then ending = vbCrLf
    End If
    position = stopPos + Len(ending)
End Sub

Function IsBlank(ByVal ch)
    IsBlank = (ch = " " Or ch = vbTab)
End Function

Function ReadBytes(ByVal path)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile path
    If stream.Size = 0 Then Fail "An empty file was encountered: " & path
    ReadBytes = stream.Read
    stream.Close
End Function

Function DecodeBytes(ByRef bytes, ByVal charset)
    Dim stream, value
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = adTypeBinary
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = adTypeText
    stream.Charset = charset
    value = stream.ReadText
    If Len(value) > 0 Then
        If Left(value, 1) = ChrW(&HFEFF) Then value = Mid(value, 2)
    End If
    DecodeBytes = value
    stream.Close
End Function

Function EncodeText(ByVal value, ByVal charset, ByVal hasBom)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = adTypeText
    stream.Charset = charset
    stream.Open
    stream.WriteText value
    stream.Position = 0
    stream.Type = adTypeBinary
    If charset = "utf-8" And Not hasBom Then stream.Position = 3
    EncodeText = stream.Read
    stream.Close
End Function

Sub WriteBytes(ByVal path, ByRef bytes, ByVal saveOption)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = adTypeBinary
    stream.Open
    stream.Write bytes
    stream.SaveToFile path, saveOption
    stream.Close
End Sub

Function CheckBom(ByRef bytes, ByVal charset)
    Dim length, utf8Bom, utf16Bom
    length = LenB(bytes)
    utf8Bom = False
    utf16Bom = False
    If length >= 3 Then
        utf8Bom = (AscB(MidB(bytes, 1, 1)) = &HEF And AscB(MidB(bytes, 2, 1)) = &HBB And AscB(MidB(bytes, 3, 1)) = &HBF)
    End If
    If length >= 2 Then
        utf16Bom = (AscB(MidB(bytes, 1, 1)) = &HFF And AscB(MidB(bytes, 2, 1)) = &HFE)
        If AscB(MidB(bytes, 1, 1)) = &HFE And AscB(MidB(bytes, 2, 1)) = &HFF Then Fail "UTF-16 BE is not supported."
    End If
    If length >= 4 Then
        If utf16Bom And AscB(MidB(bytes, 3, 1)) = 0 And AscB(MidB(bytes, 4, 1)) = 0 Then Fail "UTF-32 is not supported."
    End If
    If utf8Bom And charset <> "utf-8" Then Fail "The file has a UTF-8 BOM. Set INI_CHARSET=utf-8."
    If utf16Bom And charset <> "unicode" Then Fail "The file has a UTF-16 LE BOM. Set INI_CHARSET=unicode."
    If charset = "unicode" And Not utf16Bom Then Fail "UTF-16 LE input must have a BOM."
    CheckBom = (utf8Bom Or utf16Bom)
End Function

Function SameBytes(ByRef first, ByRef second)
    Dim i, length
    SameBytes = False
    length = LenB(first)
    If length <> LenB(second) Then Exit Function
    For i = 1 To length
        If AscB(MidB(first, i, 1)) <> AscB(MidB(second, i, 1)) Then Exit Function
    Next
    SameBytes = True
End Function

Sub Fail(ByVal message)
    Err.Raise vbObjectError + 1000, "set_ini_test.vbs", message
End Sub
