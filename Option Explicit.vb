Option Explicit

' Longest prefix match against a 2-column range: Code | Label
Private Function LongestPrefixMatch(ByVal s As String, ByVal twoCol As Range, _
                                    ByRef outCode As String, ByRef outLabel As String) As Boolean
    Dim r As Range, bestLen As Long, c As String
    s = LCase$(s)
    LongestPrefixMatch = False
    If twoCol Is Nothing Then Exit Function
    For Each r In twoCol.Columns(1).Cells
        c = LCase$(Trim$(CStr(r.Value)))
        If Len(c) > 0 Then
            If Left$(s, Len(c)) = c Then
                If Len(c) > bestLen Then
                    bestLen = Len(c)
                    outCode = c
                    outLabel = CStr(r.Offset(0, 1).Value)
                    LongestPrefixMatch = True
                End If
            End If
        End If
    Next r
End Function

' Turn a..z into 1..26 (and keep the letter too)
Private Function LetterIndex(ByVal ch As String) As Long
    ch = LCase$(Trim$(ch))
    If ch Like "[a-z]" Then
        LetterIndex = Asc(ch) - Asc("a") + 1
    Else
        LetterIndex = 0
    End If
End Function

Public Sub ParseUTM_C10_To_E10_M10()
    Dim code As String, rest As String
    Dim codeX As String, labelX As String
    Dim yr As String, mo As String, dy As String
    Dim serialLetter As String, serialIdx As Long
    Dim locCode As String, locLabel As String
    Dim rngMenu As Range, rngItem As Range, rngPerson As Range, rngDevice As Range, rngLoc As Range

    ' === Read input ===
    code = Trim$(Range("C10").Value)
    If Len(code) = 0 Then
        MsgBox "C10 is empty.", vbExclamation: Exit Sub
    End If
    rest = code

    ' === Bind mapping ranges by their names ===
    On Error Resume Next
    Set rngMenu = Range("MENU_TABLE")
    Set rngItem = Range("ITEM_TABLE")
    Set rngPerson = Range("PERSON_TABLE")
    Set rngDevice = Range("DEVICE_TABLE")
    Set rngLoc = Range("LOCATION_TABLE")
    On Error GoTo 0

    If rngMenu Is Nothing Or rngItem Is Nothing Or rngPerson Is Nothing Or rngDevice Is Nothing Or rngLoc Is Nothing Then
        MsgBox "One or more mapping ranges (MENU_TABLE, ITEM_TABLE, PERSON_TABLE, DEVICE_TABLE, LOCATION_TABLE) are missing.", vbCritical
        Exit Sub
    End If

    ' Clear previous outputs
    Range("E10:M10").ClearContents

    ' 1) Menu (variable length, longest prefix)
    If Not LongestPrefixMatch(rest, rngMenu, codeX, labelX) Then
        MsgBox "Menu code not found at start of: " & rest, vbExclamation: Exit Sub
    End If
    Range("E10").Value = labelX
    rest = Mid$(rest, Len(codeX) + 1)

    ' 2) Item / Name (variable length)
    If Not LongestPrefixMatch(rest, rngItem, codeX, labelX) Then
        MsgBox "Item code not found after menu. Remaining: " & rest, vbExclamation: Exit Sub
    End If
    Range("F10").Value = labelX
    rest = Mid$(rest, Len(codeX) + 1)

    ' 3) Person (variable length)
    If Not LongestPrefixMatch(rest, rngPerson, codeX, labelX) Then
        MsgBox "Person code not found after item. Remaining: " & rest, vbExclamation: Exit Sub
    End If
    Range("G10").Value = labelX
    rest = Mid$(rest, Len(codeX) + 1)

    ' 4) Date (fixed 6 digits: YYMMDD)
    If Len(rest) < 6 Or Not IsNumeric(Left$(rest, 6)) Then
        MsgBox "Expected 6-digit date (YYMMDD). Remaining: " & rest, vbExclamation: Exit Sub
    End If
    yr = "20" & Left$(rest, 2)
    mo = Mid$(rest, 3, 2)
    dy = Mid$(rest, 5, 2)
    Range("H10").Value = yr
    Range("I10").Value = mo
    Range("J10").Value = dy
    rest = Mid$(rest, 7)

    ' 5) Device (variable length)
    If Not LongestPrefixMatch(rest, rngDevice, codeX, labelX) Then
        MsgBox "Device code not found after date. Remaining: " & rest, vbExclamation: Exit Sub
    End If
    Range("K10").Value = labelX
    rest = Mid$(rest, Len(codeX) + 1)

    ' 6) Serial letter (exactly 1 letter)
    If Len(rest) < 1 Or Not (Left$(LCase$(rest), 1) Like "[a-z]") Then
        MsgBox "Expected 1-letter serial after device. Remaining: " & rest, vbExclamation: Exit Sub
    End If
    serialLetter = Left$(LCase$(rest), 1)
    serialIdx = LetterIndex(serialLetter)
    Range("L10").Value = serialLetter & " (" & serialIdx & ")"
    rest = Mid$(rest, 2)

    ' 7) Location (exactly 1 digit 1..8, then map)
    If Len(rest) < 1 Or Not (Left$(rest, 1) Like "[1-8]") Then
        MsgBox "Expected 1-digit location (1..8). Remaining: " & rest, vbExclamation: Exit Sub
    End If
    locCode = Left$(rest, 1)
    ' Lookup location label
    If Not LongestPrefixMatch(locCode, rngLoc, codeX, labelX) Then
        labelX = "Unknown"
    End If
    Range("M10").Value = labelX
    ' rest should be empty now (optional check)
End Sub
