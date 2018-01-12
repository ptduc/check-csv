Attribute VB_Name = "Validation"
'All checkbox are not checked
Function allCheckboxNotChecked() As Boolean
    
    allCheckboxNotChecked = True
    If ActiveSheet.CheckBoxes("chkAll").value <> 1 And ActiveSheet.CheckBoxes("chkAll").value <> 2 Then
        allCheckboxNotChecked = False
        MsgBox ALL_CHECKBOX_NOT_CHECKED_MSG
    End If
End Function

Function extractionEmptyFileSize(errorRows As String)
    errorRows = Replace(Trim(errorRows), " ", ",")
    If IsEmpty(errorRows) = False And errorRows <> "" Then
        MsgBox Replace(OVERVIEW_FILE_ERROR_MSG, "%{rowNum}", errorRows)
    End If
End Function

Function vaidateFileSize(ByVal filePath As String, ByVal limit As String, ByVal fileName As String, ByVal flagRecordSize As String) As Boolean
    vaidateFileSize = True
    
    Dim size As Double
    size = file.size(filePath)
    
    If (flagRecordSize = "ëSåè") Then
        If (size > limit) Then
            vaidateFileSize = False
            Log.ERROR (Replace(FILE_OVER_LIMIT_SIZE_MSG, "%{fileName}", fileName))
        End If
    End If
End Function

Function validateExtenstion(ByVal filePath As String, ByVal exension As String, ByVal fileName As String) As Boolean

    If (extension(filePath) = exension) Then
        validateExtenstion = True
    Else
        validateExtenstion = False
        Log.ERROR (Replace(ERROR_FILE_EXTENSION, "%{fileName}", fileName))
    End If
End Function

Function validateNewLineCharacter(ByVal newLineCharacter As String, ByVal newLineDeclare As String, ByVal fileName As String) As Boolean
    If (newLineCharacter = newLineDeclare) Then
        validateNewLineCharacter = True
    Else
        validateNewLineCharacter = False
        Log.ERROR (Replace(ERROR_END_LINE_CHARACTER, "%{fileName}", fileName))
    End If
End Function

Function validateNameRule(ByVal path As String, ByVal nameRule As String, ByVal fileName As String) As Boolean
    Dim file As String
    file = Dir(path & "*")
    'fileNameOrignal = Split(file, ".")(0)
    fileNameOrignal = getFileName(path)
    startFileName = Left$(fileNameOrignal, 1)
    
    'If Split(fileNameOrignal, "_")(0) <> Split(nameRule, "_")(0) Then
    If InStr(1, fileNameOrignal, nameRule) = 0 Then
        checkNameRule = False
        Log.ERROR (Replace(ERROR_NAME_RULE, "%{fileName}", fileName))
    Else
        checkNameRule = True
    End If
    
    If (startFileName = "_") Or IsLetter(startFileName) Then
        checkNameRule = True
    Else
        checkNameRule = False
        Log.ERROR (Replace(ERROR_NAME_RULE_APHABLE, "%{fileName}", fileName))
    End If
    
End Function

Function validateBom(ByVal detectBOM As Boolean, ByVal fileName As String) As Boolean
    validateBom = True
    If detectBOM = True Then
        validateBom = False
        Log.ERROR (Replace(ERROR_BOM, "%{fileName}", fileName))
    End If
End Function

Function validateEncoding(ByVal endcodingFile As String, ByVal endcodingType As String, ByVal fileName As String) As Boolean
    If (endcodingFile = endcodingType) Then
        validateEncoding = True
    Else
        validateEncoding = False
        Log.ERROR (Replace(ERROR_ENDCODING, "%{fileName}", fileName))
    End If
End Function

Function validateMaxRecord(ByVal fileRecord As Integer, ByVal maxRecordRule As Integer, ByVal fileName As String, ByVal flagRecordSize As String) As Boolean
    validateMaxRecord = True
    If (fileRecord > maxRecordRule) And (flagRecordSize = "ëSåè") Then
        validateMaxRecord = False
        Log.ERROR (Replace(ERROR_OVER_RECORD, "%{fileName}", fileName))
    End If
End Function

Function IsLetter(ByVal strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function

Function validateQuote(ByVal textData As String, ByVal regEx As RegExp)

    
    
    validateQuote = regEx.Test(textData)
    
End Function
'huynnp
Function duplicateNamePattern(errorPatterns As String)
    errorPatterns = Replace(Trim(errorPatterns), " ", ",")
    If IsEmpty(errorPatterns) = False And errorPatterns <> "" Then
        MsgBox Replace(ERROR_DUPLICATE_NAME_PATTERN, "%{namePattern}", errorPatterns)
    End If
End Function

