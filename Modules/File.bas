Attribute VB_Name = "File"
Function extension(ByVal path As String)
    Dim file As String
    file = Dir(path & fileName & "*")
    extension = Right$(file, Len(file) - InStrRev(file, "."))
End Function

'File size return number of KB
Function size(ByVal path As String) As Double
    size = Number.roundUp(FileLen(path) / 1024, 5)
End Function

Function newLineCharacter(ByVal path As String, vbVal As Integer)
    Dim buffer As String, original As String
    Dim file As Integer
    file = FreeFile

    Open path For Input As #file
    buffer = Input(LOF(file), file)
    Close #file

    If InStr(1, buffer, vbCrLf) > 0 Then
        If vbVal = 1 Then
            original = vbCrLf
        Else
            original = "CRLF"
        End If
    ElseIf InStr(1, buffer, vbLf) > 0 Then
        If vbVal = 1 Then
            original = vbLf
        Else
            original = "LF"
        End If
    ElseIf InStr(1, buffer, vbCr) > 0 Then
        If vbVal = 1 Then
            original = vbCr
        Else
            original = "CR"
        End If
    End If


    newLineCharacter = original
End Function

Function csvToArray(ByVal emiFilePath As String, ByVal delimiter As String, ByVal newLineChar As String) As Variant()

    Dim TextFile As Integer
    Dim filePath As String
    Dim FileContent As String
    Dim LineArray() As String
    Dim DataArray() As Variant
    'Inputs
    filePath = emiFilePath

    'Open the text file in a Read State
    TextFile = FreeFile
    Open filePath For Input As TextFile

    'Store file content inside a variable
    FileContent = Input(LOF(TextFile), TextFile)

    'Close Text File
    Close TextFile

    'Separate Out lines of data
    LineArray = Split(FileContent, newLineChar, -1, vbTextCompare)
    ReDim DataArray(LBound(LineArray) To UBound(LineArray))
    Dim i As Long
    'Separate fields inside the lines
    For i = LBound(LineArray) To UBound(LineArray)
      DataArray(i) = Split(LineArray(i), delimiter, -1, vbTextCompare)
    Next i
    csvToArray = DataArray
End Function

'Check if a file exists
Function exists(ByVal filePath As String) As Boolean

    Dim obj_fso As Object

    Set obj_fso = CreateObject("Scripting.FileSystemObject")
    exists = obj_fso.FileExists(filePath)
End Function

'Check file encoding
'File encoding with BOM always return true encoding, otherwise return ANSI
Function encoding(ByVal strFileName As String) As String

    Dim intFileNumber As Integer
    Dim lngFileSize As Long
    Dim strBuffer As String
    Dim lngCharNumber As Long
    Dim strCharacter As String * 1

    'Get the next available File Number
    intFileNumber = FreeFile

    Open strFileName For Binary Access Read As #intFileNumber

    lngFileSize = LOF(intFileNumber)    'How large is the File in Bytes?
    If lngFileSize > 10000 Then
        lngFileSize = 10000
    End If
    strBuffer = Space$(lngFileSize)     'Set Buffer Size to File Length

    Get #intFileNumber, , strBuffer     'Grab a Chunk of Data from the File
    Close #intFileNumber

    Dim bytes() As Byte
    ReDim bytes(0 To lngFileSize - 1)
    arrayIndex = 0

    'Display results on a Byte-by-Byte basic
    For lngCharNumber = 1 To lngFileSize
      strCharacter = LCase(Mid(strBuffer, lngCharNumber, 1))

      value = Hex$(Asc(strCharacter))
      'Debug.Print "value: " & value

      'for 1 byte
      If Len(value) = 2 Then
        bytes(arrayIndex) = CByte("&H" & value)
        arrayIndex = arrayIndex + 1
      'for 2 bytes
      ElseIf Len(value) = 4 Then
        firstByte = Left$(value, 2)
        lastByte = Right$(value, 2)
        bytes(arrayIndex) = CByte("&H" & firstByte)
        bytes(arrayIndex + 1) = CByte("&H" & lastByte)
        arrayIndex = arrayIndex + 1
      End If

    Next lngCharNumber

    tmp = JudgeCode(bytes)

    If tmp = "JIS" Or tmp = "SJIS" Or tmp = "EUC" Or tmp = "Shift-JIS" Then
        encoding = "Shift-JIS"
    ElseIf tmp = "UTF-8" Or tmp = "UNI" Or tmp = "UNICODE" Then
        encoding = "UTF-8"
    End If
    Debug.Print "encoding: " & encoding
End Function

'Return true -> finish, false -> continue
Function detectBOM(ByVal filePath As String) As Boolean

  Dim b1 As Byte, b2 As Byte
  Open filePath For Binary As #1
  Get #1, , b1
  Get #1, , b2
  Close #1

  If b1 = &HFF And b2 = &HFE Then
    detectBOM = True
  ElseIf b1 = &HFE And b2 = &HFF Then
    detectBOM = True
  ElseIf b1 = &HEF And b2 = &HBB Then
    detectBOM = True
  Else
    detectBOM = False
  End If
End Function

'Get Quantity Of File
Function getFileLine(ByVal filePath As String) As Long

    Dim s As String
    Dim n As Long

    Open filePath For Input As 1

    n = 0
    Do While Not EOF(1)
        Line Input #1, s
        n = n + 1
    Loop
    Close #1
    getFileLine = n

End Function
