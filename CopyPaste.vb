Public Sub PasteAndCheck()
    Dim srcValue As Variant
    srcValue = Range("Q17").Value

    If IsEmpty(srcValue) Then
        MsgBox "Q17 is empty. Nothing to copy.", vbExclamation
        Exit Sub
    End If

    Range("L10").Value = srcValue
    MsgBox "✅ Value copied from Q17 to L10: " & srcValue, vbInformation

    ' Now run the "Check" macro
    Call  ParseUTM_L10_To_L12_T12()
End Sub