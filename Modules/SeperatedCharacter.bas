Attribute VB_Name = "SeperatedCharacter"
'Check SEPERATED CHARACTER
'@param strLine                   Line to check
'@param strSeperatedChar   Seperated character
'@param intNoOfCol             Defined number of column of file
Function checkSeperatedCharacter(ByVal strLine As String, strSeperatedChar As String, intNoOfCol As Integer) As Boolean
    Dim vntColumns As Variant
    Dim vntColumnWithQuote As Variant
    Dim strTmpLine As String
    Dim strTmpLineQuote As String
    Dim vntTmpLineListByQuote As Variant

    If InStr(strLine, Chr(34)) = 0 Then
        'Line without double quote
        vntColumns = Split(strLine, strSeperatedChar)
    Else
        'LIne with double quote
        vntColumnWithQuote = Split(strLine, strSeperatedChar)
        For i = LBound(vntColumnWithQuote) To UBound(vntColumnWithQuote)
            If InStr(vntColumnWithQuote(i), Chr(34)) = 0 Then
                'Column without double quote
                strTmpLine = strTmpLine & vntColumnWithQuote(i) & strSeperatedChar
            Else
                'Column with double quote Å® file whole column
                Do
                    strTmpLineQuote = strTmpLineQuote & vntColumnWithQuote(i)
                    vntTmpLineListByQuote = Split(strTmpLineQuote, Chr(34))
                    If UBound(vntTmpLineListByQuote) Mod 2 <> 0 Then
                        i = i + 1
                    End If
                Loop While UBound(vntTmpLineListByQuote) Mod 2 <> 0
                strTmpLine = strTmpLine & Replace(strTmpLineQuote, strSeperatedChar, "") & strSeperatedChar
            End If
        Next
        'Remove end comma
        strTmpLine = Left(strTmpLine, Len(strTmpLine) - 1)

        vntColumns = Split(strTmpLine, strSeperatedChar)
    End If

    'Compare number of columns
    If UBound(vntColumns) + 1 = intNoOfCol Then
        checkSeperatedCharacter = True
    Else
        checkSeperatedCharacter = False
    End If

End Function

'Get process line of MAC file
Function getProcessLine(ByVal vntLines As Variant, ByVal startRowData As Integer) As Variant
    Dim arrProcessLine() As Variant
    Dim vntListByQuote As Variant
    Dim strTmpLine As String
    Dim processLineIndex As Integer

    processLineIndex = 0

    For lineIndex = LBound(vntLines) To UBound(vntLines)
        'Limit 1000 record of data
        If (processLineIndex + 1) > (startRowData + 1000) Then
            Exit For
        End If

        If (processLineIndex + 1) > startRowData Then
            If lineIndex = UBound(vntLines) And vntLines(lineIndex) = "" Then Exit For

            ReDim Preserve arrProcessLine(processLineIndex)

            If InStr(vntLines(lineIndex), Chr(34)) = 0 Then
                'Not contain double quote

                    arrProcessLine(processLineIndex) = vntLines(lineIndex)

            ElseIf InStr(vntLines(lineIndex), Chr(34)) > 0 Then
                'Contain double quote
                Do
                    'Find whole line
                    strTmpLine = strTmpLine & vntLines(lineIndex)
                    vntListByQuote = Split(strTmpLine, Chr(34))
                    If UBound(vntListByQuote) Mod 2 <> 0 Then
                        lineIndex = lineIndex + 1
                    End If
                Loop While UBound(vntListByQuote) Mod 2 <> 0
                'Assign line to process line
                arrProcessLine(processLineIndex) = strTmpLine
                strTmpLine = ""
            End If
        End If
        processLineIndex = processLineIndex + 1
    Next

    getProcessLine = arrProcessLine
End Function
