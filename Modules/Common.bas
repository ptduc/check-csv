Attribute VB_Name = "Common"
'Public Const DATA_CHECK_SHEET = "データチェックツール"
'Public Const FILE_LIST_SHEET = "IFファイル一覧"
'Public Const CLEAR_SUCCESS_MSG = "対象をクリアしました。"
'Public Const CLEAR_CONFIRM_MSG = "対象一覧表の内容をクリアしますが、よろしいでしょうか。"
'Public Const FILE_RULE_NAME_RANGE = "D5:D28"

Function lastRowDataCheckSheet() As Long
    lastRowDataCheckSheet = Sheet.lastHasData(1, DATA_CHECK_SHEET)
End Function

Function lastRowFileListSheet() As Long
    lastRowFileListSheet = Sheet.lastHasData(1, FILE_LIST_SHEET)
End Function

Function dataCheckSheet() As Worksheet
    Set dataCheckSheet = Worksheets(DATA_CHECK_SHEET)
End Function

Function fileListSheet() As Worksheet
    Set fileListSheet = Worksheets(FILE_LIST_SHEET)
End Function

Function setRunning(rowNum As Long)
    dataCheckSheet.Range("I" & rowNum).value = "Running"
End Function

Function setFinish(rowNum As Long)
    dataCheckSheet.Range("I" & rowNum).value = "Finished"
End Function

Function setCancel(rowNum As Long)
    dataCheckSheet.Range("I" & rowNum).value = "Cancel"
End Function

Function definiteSheet(ByVal sheetName As String) As Worksheet
    Set definiteSheet = ActiveWorkbook.Worksheets(sheetName)
End Function

Function readFileByLine(ByVal filePath As String, ByVal limitLine As Long) As String
    Dim fileNo As Integer
    Dim textData As String
    Dim textRow As String
    
    fileNo = FreeFile
    Open filePath For Input As #fileNo
    
    currentLine = 1
    Do While Not EOF(fileNo) And currentLine <= limitLine
        Line Input #fileNo, textRow
        textData = textData & textRow
    Loop
    
    Close #fileNo
    readFileByLine = textData
    
End Function

Function getFileName(ByVal filePath As String) As String
    Dim fso As New FileSystemObject
    getFileName = fso.GetBaseName(filePath)
End Function
