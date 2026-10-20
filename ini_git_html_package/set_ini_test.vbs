Option Explicit

' Version 2: inspect COM byte arrays with LenB / MidB / AscB.

' Usage: cscript //nologo set_ini_test.vbs "file.ini" "new-value" [charset]
' charset: utf-8 (default), shift_jis, or unicode (UTF-16 LE with BOM).
' This source is ASCII-only. ChrW builds the Japanese section name.
' Only standalone test= tokens in active lines of that section are edited.

Const adTypeBinary = 1
Const adTypeText = 2
Const adSaveCreateNotExist = 1
Const adSaveCreateOverWrite = 2

Dim fso, tempPath, backupPath, stage, errorNumber, errorText
tempPath = ""
backupPath = ""
stage = "initialization"

WScript.Echo "INI updater v2"

On Error Resume Next
Main
errorNumber = Err.Number
errorText = Err.Description
On Error GoTo 0

If errorNumber <> 0 Then
    WScript.Echo "ERROR during " & stage & ": " & errorText & " (0x" & Hex(errorNumber) & ")"
    If backupPath <> "" Then WScript.Echo "Backup: " & backupPath
    On Error Resume Next
    If tempPath <> "" Then fso.DeleteFile tempPath, True
    On Error GoTo 0
    WScript.Quit 1
End If
WScript.Quit 0

Sub Main()
    Dim path, newValue, charset, raw, original, updated, roundTrip, outputBytes
    Dim hasBom, targetSection, position, ending, line, trimmed, closePos, inSection
    Dim matched, changed, fileInfo, checkBytes, candidateBackup

    If WScript.Arguments.Count < 2 Or WScript.Arguments.Count > 3 Then
        Fail "Usage: cscript //nologo set_ini_test.vbs ""file.ini"" ""new-value"" [charset]"
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.GetAbsolutePathName(WScript.Arguments(0))
    newValue = WScript.Arguments(1)
    charset = "utf-8"
    If WScript.Arguments.Count = 3 Then charset = LCase(Trim(WScript.Arguments(2)))
    Select Case charset
        Case "utf8": charset = "utf-8"
        Case "sjis", "cp932", "shift-jis": charset = "shift_jis"
        Case "utf-16", "utf-16le": charset = "unicode"
        Case "utf-8", "shift_jis", "unicode"
        Case Else: Fail "Supported charsets: utf-8, shift_jis, unicode."
    End Select

    ValidateValue newValue
    If Not fso.FileExists(path) Then Fail "INI file not found: " & path
    Set fileInfo = fso.GetFile(path)
    If fileInfo.Size = 0 Then Fail "The INI file is empty."
    If (fileInfo.Attributes And 1) <> 0 Then Fail "The INI file is read-only."

    stage = "reading the INI"
    raw = ReadBytes(path)
    stage = "checking the file BOM"
    hasBom = CheckBom(raw, charset)
    stage = "decoding the INI"
    original = DecodeBytes(raw, charset)
    If InStr(original, ChrW(0)) > 0 Then Fail "NUL character found; check INI_CHARSET."

    ' Refuse a charset mismatch or a conversion that would alter other bytes.
    stage = "checking the original encoding"
    roundTrip = EncodeText(original, charset, hasBom)
    If Not SameBytes(raw, roundTrip) Then
        Fail "The file cannot be preserved with INI_CHARSET=" & charset & ". Check its encoding."
    End If

    stage = "parsing the INI"
    targetSection = ChrW(&H8A2D) & ChrW(&H5B9A) & "1"
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
                    End If
                ElseIf inSection Then
                    line = RewriteLine(line, newValue, matched, changed)
                End If
            End If
        End If
        updated = updated & line & ending
    Loop

    If matched = 0 Then Fail "No active test= token found in the target section. No file was changed."
    If changed = 0 Then
        WScript.Echo "No change needed. " & matched & " value(s) already match."
        Exit Sub
    End If

    stage = "preparing the updated file"
    outputBytes = EncodeText(updated, charset, hasBom)
    If DecodeBytes(outputBytes, charset) <> updated Then
        Fail "The new value cannot be represented in INI_CHARSET=" & charset & "."
    End If
    tempPath = fso.BuildPath(fso.GetParentFolderName(path), "ini_edit_" & fso.GetTempName)
    WriteBytes tempPath, outputBytes, adSaveCreateNotExist
    checkBytes = ReadBytes(tempPath)
    If Not SameBytes(outputBytes, checkBytes) Then Fail "Temporary file verification failed."

    'stage = "creating a backup"
    'candidateBackup = NextBackupPath(path)
    'fso.CopyFile path, candidateBackup, False
    'backupPath = candidateBackup
    'checkBytes = ReadBytes(backupPath)
    'If Not SameBytes(raw, checkBytes) Then Fail "The INI changed while this script was running. Update cancelled."
    checkBytes = ReadBytes(path)
    If Not SameBytes(raw, checkBytes) Then Fail "The INI changed while this script was running. Update cancelled."

    ' Keep the existing pathname and file permissions by writing to that file.
    ' The backup is retained if a write or verification fails.
    stage = "writing the INI (the backup is retained if this fails)"
    WriteBytes path, outputBytes, adSaveCreateOverWrite
    checkBytes = ReadBytes(path)
    If Not SameBytes(outputBytes, checkBytes) Then Fail "Updated file verification failed; use the backup to restore."
    fso.DeleteFile tempPath, True
    tempPath = ""

    WScript.Echo "Updated: " & changed & " value(s)."
    WScript.Echo "File: " & path
    'WScript.Echo "Backup: " & backupPath
End Sub

Sub ValidateValue(ByVal value)
    Dim i, ch
    If Len(value) = 0 Then Fail "The new value must not be empty."
    For i = 1 To Len(value)
        ch = Mid(value, i, 1)
        If AscW(ch) >= 0 And AscW(ch) <= 32 Then Fail "The new value must not contain spaces or control characters."
        If ch = ";" Or ch = "#" Or ch = """" Then Fail "The new value must not contain semicolons, #, or double quotes."
    Next
End Sub

Function RewriteLine(ByVal line, ByVal newValue, ByRef matched, ByRef changed)
    Dim i, n, ch, boundary, eqPos, valueStart, valueEnd, quoteEnd, copied, result
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
            If boundary And StrComp(Mid(line, i, 4), "test", vbTextCompare) = 0 Then
                eqPos = i + 4
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
                        If valueEnd = 0 Then Fail "Unclosed quote in a test= value."
                        If Mid(line, valueEnd + 1, 1) = """" Then Fail "Escaped quotes in a test= value are not supported."
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
                                Fail "Ambiguous empty test= value followed by another setting. Update cancelled."
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

Function NextBackupPath(ByVal path)
    Dim stamp, candidate, counter, t
    t = Now
    stamp = Year(t) & Pad2(Month(t)) & Pad2(Day(t)) & "_" & Pad2(Hour(t)) & Pad2(Minute(t)) & Pad2(Second(t))
    candidate = path & ".bak_" & stamp
    counter = 1
    Do While fso.FileExists(candidate)
        candidate = path & ".bak_" & stamp & "_" & counter
        counter = counter + 1
    Loop
    NextBackupPath = candidate
End Function

Function Pad2(ByVal value)
    Pad2 = Right("0" & CStr(value), 2)
End Function

Sub Fail(ByVal message)
    Err.Raise vbObjectError + 1000, "set_ini_test.vbs", message
End Sub
