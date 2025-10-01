Option Explicit

Private Function SplitCode(ByVal code As String, _
                           ByRef prefix As String, _
                           ByRef letter As String, _
                           ByRef suffix As String) As Boolean
    Dim i As Long, nDigits As Long
    code = Trim$(code)
    If Len(code) < 2 Then Exit Function

    i = Len(code)
    Do While i >= 1 And Mid$(code, i, 1) Like "[0-9]"
        nDigits = nDigits + 1
        i = i - 1
    Loop
    If nDigits = 0 Or i < 1 Then Exit Function

    suffix = Right$(code, nDigits)
    letter = Mid$(code, i, 1)
    prefix = Left$(code, i - 1)

    If LCase$(letter) < "a" Or LCase$(letter) > "z" Then Exit Function
    SplitCode = True
End Function

Public Sub GenerateSerialWithRules()
    Dim baseCode As String, prefix As String, letter As String, suffix As String
    Dim count As Long, i As Long
    Dim startRow As Long
    Dim startCell As Range
    Dim newLetter As String
    Dim v As Variant

    ' 🧹 先清空 C14:C39
    Range("C14:C39").ClearContents

    ' 讀取基準碼和數量
    baseCode = Range("C10").Value
    v = Range("C12").Value
    If Len(Trim$(v)) = 0 Or Not IsNumeric(v) Then
        MsgBox "C12（數量）需為正整數。", vbExclamation: Exit Sub
    End If
    If v <> Fix(v) Or v <= 0 Then
        MsgBox "C12（數量）需為正整數。", vbExclamation: Exit Sub
    End If
    count = CLng(v)

    If Not SplitCode(baseCode, prefix, letter, suffix) Then
        MsgBox "C10 格式需為：<前綴><英文字母><數字>，例如：XXXXXXXXe1", vbExclamation
        Exit Sub
    End If

    ' a→C14, b→C15, …, z→C39
    startRow = 14 + (Asc(LCase$(letter)) - Asc("a"))
    If startRow < 14 Or startRow > 39 Then
        MsgBox "起始字母必須介於 a~z。", vbExclamation
        Exit Sub
    End If
    Set startCell = Range("C" & startRow)

    ' 產生序號
    For i = 0 To count - 1
        newLetter = Chr$(Asc(LCase$(letter)) + i)
        If newLetter > "z" Then
            MsgBox "超出 z，已停止。請減少數量或換起始字母。", vbExclamation
            Exit Sub
        End If
        startCell.Offset(i, 0).Value = prefix & newLetter & suffix
    Next i
End Sub