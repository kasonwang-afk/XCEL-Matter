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

Public Sub ParseUTM_L10_To_L12_T12()
    Dim code As String, rest As String
    Dim codeX As String, labelX As String
    Dim yr As String, mo As String, dy As String
    Dim serialLetter As String, serialIdx As Long
    Dim locCode As String, locLabel As String
    Dim rngMenu As Range, rngDevice As Range, rngTarget As Range, rngProduct As Range, rngLoc As Range

    ' === Read input ===
    code = Trim$(Range("L10").Value)
    If Len(code) = 0 Then
        MsgBox "L10 is empty.", vbExclamation: Exit Sub
    End If
    rest = code

    ' === Bind mapping ranges by their names ===
    On Error Resume Next
    Set rngMenu = Range("Data_Menu")
    Set rngDevice = Range("Data_Device")
    Set rngTarget = Range("Data_Target")
    Set rngProduct = Range("Data_Product")
    Set rngLoc = Range("Data_Company")
    On Error GoTo 0

    If rngMenu Is Nothing Or rngDevice Is Nothing Or rngTarget Is Nothing Or rngProduct Is Nothing Or rngLoc Is Nothing Then
        MsgBox "One or more mapping ranges (Data_Menu, Data_Device, Data_Target, Data_Product, Data_Company) are missing.", vbCritical
        Exit Sub
    End If

    ' Clear previous outputs
    Range("L12:T12").ClearContents

    ' 1) Menu (variable length, longest prefix)
    If Not LongestPrefixMatch(rest, rngMenu, codeX, labelX) Then
        MsgBox "Menu code not found at start of: " & rest, vbExclamation: Exit Sub
    End If
    Range("L12").Value = labelX
    rest = Mid$(rest, Len(codeX) + 1)

    ' 2) Device / Name (variable length)
    If Not LongestPrefixMatch(rest, rngDevice, codeX, labelX) Then
        MsgBox "Device code not found after menu. Remaining: " & rest, vbExclamation: Exit Sub
    End If
    Range("M12").Value = labelX
    rest = Mid$(rest, Len(codeX) + 1)

    ' 3) Target (variable length)
    If Not LongestPrefixMatch(rest, rngTarget, codeX, labelX) Then
        MsgBox "Target code not found after Device. Remaining: " & rest, vbExclamation: Exit Sub
    End If
    Range("N12").Value = labelX
    rest = Mid$(rest, Len(codeX) + 1)

    ' 4) Date (fixed 6 digits: YYMMDD)
    If Len(rest) < 6 Or Not IsNumeric(Left$(rest, 6)) Then
        MsgBox "Expected 6-digit date (YYMMDD). Remaining: " & rest, vbExclamation: Exit Sub
    End If
    yr = "20" & Left$(rest, 2)
    mo = Mid$(rest, 3, 2)
    dy = Mid$(rest, 5, 2)
    Range("O12").Value = yr
    Range("P12").Value = mo
    Range("Q12").Value = dy
    rest = Mid$(rest, 7)

    ' 5) Product (variable length)
    If Not LongestPrefixMatch(rest, rngProduct, codeX, labelX) Then
        MsgBox "Product code not found after date. Remaining: " & rest, vbExclamation: Exit Sub
    End If
    Range("R12").Value = labelX
    rest = Mid$(rest, Len(codeX) + 1)

    ' 6) Serial letter (exactly 1 letter)
    If Len(rest) < 1 Or Not (Left$(LCase$(rest), 1) Like "[a-z]") Then
        MsgBox "Expected 1-letter serial after Product. Remaining: " & rest, vbExclamation: Exit Sub
    End If
    serialLetter = Left$(LCase$(rest), 1)
    serialIdx = LetterIndex(serialLetter)
    Range("S12").Value = serialLetter & " (" & serialIdx & ")"
    rest = Mid$(rest, 2)

    ' 7) Company (exactly 1 digit 1..8, then map)
    If Len(rest) < 1 Or Not (Left$(rest, 1) Like "[1-8]") Then
        MsgBox "Expected 1-digit Company (1..8). Remaining: " & rest, vbExclamation: Exit Sub
    End If
    locCode = Left$(rest, 1)
    ' Lookup Company label
    If Not LongestPrefixMatch(locCode, rngLoc, codeX, labelX) Then
        labelX = "Unknown"
    End If
    Range("T12").Value = labelX
    ' rest should be empty now (optional check)
End Sub
