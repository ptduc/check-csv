Attribute VB_Name = "Main"
'Sheet
Public Const DATA_CHECK_SHEET = "データチェックツール"
Public Const FILE_LIST_SHEET = "IFファイル一覧"
Public Const TABLE_DEFINITE_TEMPLATE = "【カラム定義】テンプレート"
'Range
Public Const DATA_CHECK_RANGE = "C6:M"
Public Const FILE_RULE_NAME_RANGE = "D5:D28"
Public Const DEFINITE_TABLE_FIRST_COL = 22
'Column Name "b→dash事前データチェックツール（Ver0.10）"
Public Const COL_DATA_CHECK_CHECK_BOX = "B"
Public Const COL_DATA_CHECK_NO = "C"
Public Const COL_DATA_CHECK_FILE_NAME = "D"
Public Const COL_DATA_CHECK_FILE_NAME_PATTERN = "E"
Public Const COL_DATA_CHECK_MAX_QUANTITY_RECORD = "G"
Public Const COL_DATA_CHECK_MAX_FILE_SIZE = "H"
Public Const COL_DATA_CHECK_FILE_PATH = "I"
Public Const COL_DATA_CHECK_STATUS_CHECK = "K"
Public Const COL_DATA_CHECK_DATE = "L"
Public Const COL_DATA_CHECK_SAVE = "M"
'Column Name "IFファイル一覧"
Public Const COL_FILE_LIST_ROW_NUM = "B"
Public Const COL_FILE_LIST_FILE_NAME = "C"
Public Const COL_FILE_LIST_NAME_PATTERN = "D"
Public Const COL_FILE_LIST_FILE_TYPE = "E"
Public Const COL_FILE_LIST_DELIMITER = "F"
Public Const COL_FILE_LIST_ENCODING = "G"
Public Const COL_FILE_LIST_HEADER = "H"
Public Const COL_FILE_LIST_HEADER_START_ROW_DATA = "I"
Public Const COL_FILE_LIST_NEW_LINE = "J"
Public Const COL_FILE_LIST_FLAG_RECORD_SIZE = "L"
Public Const COL_FILE_LIST_MAX_QUANTITY_RECORD = "M"
Public Const COL_FILE_LIST_MAX_FILE_SIZE = "N"
'Column Name of "DefiniteTable"
Public Const COL_DEFINITE_TABLE_No = "B"
Public Const COL_DEFINITE_TABLE_NAME = "D"
Public Const COL_DEFINITE_TABLE_DATA_TYPE = "F"
Public Const COL_DEFINITE_TABLE_PRIMARY_KEY = "G"
Public Const COL_DEFINITE_TABLE_NOT_NULL = "H"
Public Const COL_DEFINITE_TABLE_DATE_FORMAT = "L"
'Message
Public Const CLEAR_SUCCESS_MSG = "対象をクリアしました。"
Public Const CLEAR_CONFIRM_MSG = "対象一覧表の内容をクリアしますが、よろしいでしょうか。"
Public Const OVERVIEW_FILE_ERROR_MSG = "IFファイル一覧シートの行%{rowNum}には「ファイル概要」または「ファイル命名規則」または「最大レコード数」または「 最大ファイル容量」がまだ定義されていません。"
Public Const FILE_NOT_SELECT_MSG = "ファイルパスがまだ設定されてない行があります。チェックしたい行は「選択」ボタンをクリックし、チェック対象のファイルパスを設定してください。"
Public Const ALL_CHECKBOX_NOT_CHECKED_MSG = "チェック対象ファイルはまだ設定されていません。チェックしたいファイルの行頭の□をクリックしてチェックし対象ファイルを選択してください。"
Public Const FILE_NOT_EXISTS_MSG = "%{fileName}ファイルは存在していません。"
Public Const FILE_OVER_LIMIT_SIZE_MSG = "%{fileName}：最大ファイル容量を超えています。"
Public Const NOTIFICATE_NOT_SELECT_FILE = "選択ボタンをクリックし、チェック対象のファイルパスを設定してください。"
Public Const ERROR_FILE_NAME_RULE_DUPLICATE = "%{fileName}が複数存在します。ファイル命名規則は一意になるように設定してください。"
Public Const ERROR_FILE_EXTENSION = "%{fileName}：ファイルの拡張子は定義と異なっています。"
Public Const ERROR_END_LINE_CHARACTER = "%{fileName}：チェック対象行に定義と異なる改行文字が含まれています。"
Public Const ERROR_NAME_RULE_APHABLE = "%{fileName}：ファイル名の1文字目が半角英字もしくはアンダーバーのどちらかで始まるように設定してください。"
Public Const ERROR_NAME_RULE = "%{fileName}：ファイル名はIFファイル一覧で定義された命名規則に添っていません。"
Public Const ERROR_BOM = "%{fileName}：BOM付きのファイルです。"
Public Const ERROR_ENDCODING = "%{fileName}：チェック対象行に定義と異なる文字コードが含まれています。"
Public Const ERROR_SEPERATED_CHARACTER = "%{fileName}：行%{row}のカラム数はカラム定義シートで記載されているカラム数と一致していません。また、チェック対象行のなかに区切り文字が適切ではない行が含まれています。"
Public Const ERROR_NOT_DATA_DEFINITE_TABLE = "%{fileName}のカラム定義シートが正しく定義されていません。" & vbNewLine & "１カラム以上の情報を正しく定義してください。"
Public Const ERROR_OVER_RECORD = "%{fileName}：ファイルの最大レコードを超えています。"
Public Const CHECK_EXISTS_SHEET = "該当カラム定義シートは既に存在しています。"
Public Const ERROR_NOT_EXISTS_SHEET = "%{fileName}のカラム定義シートは存在していません。" & vbNewLine & "チェックした行は「作成」ボタンをクリックし、カラム定義シートを作成してください。"
Public Const ERROR_COLUMN_NOT_NULL = "%{fileName}：行%{row}に%{column}カラムにデータを設定してください。"
'huynnp
'<ファイル概要>：行{データが重複された行数}に{カラム理論名}カラムには重複データがあります。
Public Const ERROR_COLUMN_PRIMARY_KEY = "%{fileName}：行%{row}に%{column}カラムには重複データがあります。"

Public Const ERROR_COLUMN_DATE_FORMAT = "%{fileName}：行%{row}に%{column}カラムの値のフォーマットは定義と一致していません。"
Public Const ERROR_DOUBLE_QUOTE = "%{fileName}：行%{row}に、カラム入力規則と異なるカラムが存在します。"
'huynnp
Public Const ERROR_DUPLICATE_NAME_PATTERN = "%{namePattern}が複数存在します。ファイル命名規則は一意になるように設定してください。"
Public Const STATUS_PROCESSING = "実施中"
Public Const STATUS_PROCESS_COMPLETED = "チェック完了"

'"完了（正常）"
Public Const STATUS_PROCESS_COMPLETED_OK = STATUS_PROCESS_COMPLETED
'"完了（異常あり）"
Public Const STATUS_PROCESS_COMPLETED_NOK = STATUS_PROCESS_COMPLETED
Public Const STATUS_PROCESS_INIT_FILE = "未実施"
Public Const STATUS_PROCESS_STOP = "中断"

'Variant
Public dateColumns As Scripting.Dictionary
Public parseErrorRows As Scripting.Dictionary
Public currentCheckingRow As Integer
'Checkbox check all handle
Sub chkAll_Click()
    Dim CB As CheckBox
    For Each CB In ActiveSheet.CheckBoxes
        If CB.Name <> ActiveSheet.CheckBoxes("chkAll").Name Then
            rowNum = Split(CB.Name, " ")(1)
            fileOverView = Trim(Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & rowNum).value)
            'Uncheck all
            If ActiveSheet.CheckBoxes("chkAll").value <> 1 Then
                CB.value = ActiveSheet.CheckBoxes("chkAll").value
            'Check all
            ElseIf fileOverView <> "" Then
                CB.value = ActiveSheet.CheckBoxes("chkAll").value
            End If
      End If
    Next CB
End Sub

' Checkbox list handle
Sub Mixed_State()
    Dim CB As CheckBox
    For Each CB In ActiveSheet.CheckBoxes
        If CB.Name <> ActiveSheet.CheckBoxes("chkAll").Name And CB.value <> ActiveSheet.CheckBoxes("chkAll").value And ActiveSheet.CheckBoxes("chkAll").value <> 2 Then
            ActiveSheet.CheckBoxes("chkAll").value = 2
            Exit For
        Else
            ActiveSheet.CheckBoxes("chkAll").value = CB.value
        End If
    Next CB
End Sub

'Clear button handle
Sub btnClear_Click()
    clearContent
    MsgBox CLEAR_SUCCESS_MSG
End Sub

'Checking handle
Sub btnProcess_Click()
    Log.createFileLog
    'All input validation
    'Validate all checkbox are not checked
    checkboxValid = Validation.allCheckboxNotChecked()
    fileEmptyValid = True
    fileExistsValid = True
    fileNotExistsList = ""
    Set dateColumns = New Scripting.Dictionary
    Set parseErrorRows = New Scripting.Dictionary

    If checkboxValid Then
        'Loop all checkbox
        
        For Each CB In ActiveSheet.CheckBoxes
            If CB.Name <> ActiveSheet.CheckBoxes("chkAll").Name And CB.value = 1 Then
                'Row number processing
                rowNum = Split(CB.Name, " ")(1)
                currentCheckingRow = rowNum
                '1.4 validate definiteSheet
                fileNameOverView = Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & rowNum).value
                definiteSheetName = Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & rowNum).value
                If checkExistsSheet(definiteSheetName) = Fasle Then
                    MsgBox Replace(ERROR_NOT_EXISTS_SHEET, "%{fileName}", fileNameOverView)
                    Exit Sub
                End If
                'Get file path from list
                filePath = Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowNum).value
                filePath = Trim(CStr(filePath))
                
                'Validate file not select
                If IsEmpty(filePath) Or filePath = "" Or filePath = NOTIFICATE_NOT_SELECT_FILE Then
                    fileEmptyValid = False
                    Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowNum).value = NOTIFICATE_NOT_SELECT_FILE
                    Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowNum).Font.Color = vbRed
                 Else
                    'Validate file does not exists
                    If file.exists(filePath) = False Then
                        fileExistsValid = False
                        fileNotExistsList = fileNotExistsList & vbCrLf & Replace(FILE_NOT_EXISTS_MSG, "%{fileName}", Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & rowNum).value)
                    Else
                    End If
                End If
                currentCheckingRow = currentCheckingRow + 1
            End If
       Next CB
       
       '2 Validate Duplicate file name
       Dim fileRule As Worksheet
       Set fileRule = Common.fileListSheet()
       For Each E In fileRule.Range(FILE_RULE_NAME_RANGE)
        a = WorksheetFunction.CountIf(fileRule.Range(FILE_RULE_NAME_RANGE), E)
        If a >= 2 Then
            MsgBox Replace(ERROR_FILE_NAME_RULE_DUPLICATE, "%{fileName}", E)
            Exit Sub
        End If
       Next E
       
       If fileEmptyValid = False Then
            MsgBox FILE_NOT_SELECT_MSG
       ElseIf fileExistsValid = False Then
            MsgBox Replace(fileNotExistsList, vbCrLf, "", 1, 1)
       End If
    End If
    
    'Validate file content
    If checkboxValid And fileEmptyValid And fileExistsValid Then
        
        For Each CB In ActiveSheet.CheckBoxes
            Set dateColumns = New Scripting.Dictionary
           If CB.Name <> ActiveSheet.CheckBoxes("chkAll").Name And CB.value = 1 Then
               'Row number processing
                rowNum = Split(CB.Name, " ")(1)
                currentCheckingRow = rowNum
                Call updateStatusProcess(rowNum, STATUS_PROCESSING)
               
               'Get file path from list
                Dim dataCheck As Worksheet
                Set dataCheck = Common.dataCheckSheet()
                
                filePath = dataCheck.Range(COL_DATA_CHECK_FILE_PATH & rowNum).value
                filePath = Trim(CStr(filePath))
                fileOverView = dataCheck.Range(COL_DATA_CHECK_FILE_NAME & rowNum).value
                
                Dim fileList As Worksheet
                Set fileList = Common.fileListSheet()
                              
                fileRowIndex = dataCheck.Range(COL_DATA_CHECK_SAVE & rowNum).value
                limitSize = fileList.Range(COL_FILE_LIST_MAX_FILE_SIZE & fileRowIndex).value
                fileNameRule = fileList.Range(COL_FILE_LIST_NAME_PATTERN & fileRowIndex).value
                extensionFileList = fileList.Range(COL_FILE_LIST_FILE_TYPE & fileRowIndex).value
                delimiter = fileList.Range(COL_FILE_LIST_DELIMITER & fileRowIndex).value
                newLineDeclare = fileList.Range(COL_FILE_LIST_NEW_LINE & fileRowIndex).value
                endcodingType = fileList.Range(COL_FILE_LIST_ENCODING & fileRowIndex).value
                maxRecordRule = fileList.Range(COL_FILE_LIST_MAX_QUANTITY_RECORD & fileRowIndex).value
                flagRecordSize = fileList.Range(COL_FILE_LIST_FLAG_RECORD_SIZE & fileRowIndex).value
                isHeader = fileList.Range(COL_FILE_LIST_HEADER & fileRowIndex).value
                startRowData = fileList.Range(COL_FILE_LIST_HEADER_START_ROW_DATA & fileRowIndex).value
                If isHeader = "あり" And IsNumeric(startRowData) Then
                    If startRowData < 1 Then
                        startRowData = 0
                    End If
                Else
                    startRowData = 0
                End If
                'B.2 Get quantity Column
                Dim nameSheeetDefiniteTable As String
                nameSheeetDefiniteTable = dataCheck.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & rowNum).value
                Dim sheetOfDefiniteTable As Worksheet
                Set sheetOfDefiniteTable = Common.definiteSheet(nameSheeetDefiniteTable)
                lastRow = lastHasData(1, nameSheeetDefiniteTable, "D22:D500")
                quantityColumnTable = lastRow - DEFINITE_TABLE_FIRST_COL + 1
                If quantityColumnTable < 0 Then
                    MsgBox Replace(ERROR_NOT_DATA_DEFINITE_TABLE, "%{fileName}", fileOverView)
                    GoTo NextIterationCB
                End If
                
                Dim lstColumnNotNull
                Dim lstColPrimaryKey
                lstColumnNotNull = ""
                lstColPrimaryKey = ""
                For i = DEFINITE_TABLE_FIRST_COL To lastRow Step 1
                    no = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_No & i).value
                    typeDefinite = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_DATA_TYPE & i).value
                    isPrimaryKey = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_PRIMARY_KEY & i).value
                    isNotNull = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_NOT_NULL & i).value
                    dateFormat = Trim(sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_DATE_FORMAT & i).value)
                    If isPrimaryKey = "Yes" Then
                        lstColPrimaryKey = lstColPrimaryKey & " " & no
                    End If
                    If isNotNull = "Yes" Then
                        lstColumnNotNull = lstColumnNotNull & " " & no
                    End If
                    'get date columns for 11 validate
                    If dateFormat <> "" Then
                        dateColumns.Add no, dateFormat
                    End If
                Next
                lstColPrimaryKey = Split(Trim(lstColPrimaryKey), " ")
                lstColumnNotNull = Split(Trim(lstColumnNotNull), " ")
                
                'Validate Bom
                'IsValidateBom = validateBom(detectBOM(filePath), fileOverView)
                If detectBOM(filePath) = True Then
                    Log.ERROR (Replace(ERROR_BOM, "%{fileName}", fileOverView))
                    Call updateStatusProcess(rowNum, STATUS_PROCESS_COMPLETED_NOK)
                    GoTo NextIterationCB
                End If
                
                '3 Validate file size
                IsValid = vaidateFileSize(filePath, limitSize, fileOverView, flagRecordSize)
                
                '4 Validate file extension
                IsValidExension = validateExtenstion(filePath, extensionFileList, fileOverView)
                
                '5 Validate newLineCharacter
                IsValidateNewLineCharacter = validateNewLineCharacter(newLineCharacter(filePath, 0), newLineDeclare, fileOverView)
                
                '6 Validate rule's name
                fileNameRule = Split(fileNameRule, "<")(0)
                IsValidateFileNameRule = validateNameRule(filePath, fileNameRule, fileOverView)
                
                '7.1 Validate encode
                IsValidateEncoding = validateEncoding(encoding(filePath), endcodingType, fileOverView)
                
                '8 Validate sperated character
                newLineChar = file.newLineCharacter(filePath, 1)
                FileType = fileList.Range(COL_FILE_LIST_FILE_TYPE & fileRowIndex)
                If FileType = "tsv" Then
                 FileType = vbTab
                Else
                 FileType = ","
                End If
                
                csvContent = readCSV(filePath, FileType, 1000, quantityColumnTable, fileOverView, startRowData, currentCheckingRow)
                If IsNull(csvContent) Or IsEmpty(csvContent) Then
                    GoTo NextIterationCB
                End If
                
                checkSp = checkValidateSeperatedCharacter(filePath, FileType, quantityColumnTable, fileOverView, startRowData)
                
                'Check Primary key
                For Each pkey In lstColPrimaryKey
                    rowNumCSV = 1
                    errorRowNumCSV = ""
                    errorColumnName = ""
                    For Each Row In csvContent
                        Count = 0
                        For Each Row2 In csvContent
                            If Row(1, pkey) = Row2(1, pkey) Then
                                Count = Count + 1
                            End If
                        Next
                        If Count > 1 Then
                            'errorRowNumCSV = errorRowNumCSV & "、" & rowNumCSV
                            ColumnName = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_NAME & DEFINITE_TABLE_FIRST_COL + pkey + 1).value
                            'errorColumnName = ColumnName
                            Log.ERROR (Replace(Replace(Replace(ERROR_COLUMN_PRIMARY_KEY, "%{fileName}", fileOverView), "%{row}", rowNumCSV), "%{column}", ColumnName))
                        End If
                        rowNumCSV = rowNumCSV + 1
                    Next
                    
                Next pkey
                
                rowNumCSV = 1
                For Each csvRow In csvContent
                    'Check NOT NULL
                    For Each n In lstColumnNotNull
                        n = CInt(n)
                        If csvRow(1, n) = "" Then
                            ColumnName = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_NAME & DEFINITE_TABLE_FIRST_COL + n + 1).value
                            Log.ERROR (Replace(Replace(Replace(ERROR_COLUMN_NOT_NULL, "%{fileName}", fileOverView), "%{row}", rowNumCSV), "%{column}", ColumnName))
                        End If
                    Next n
                    ' 11 validate
                    Dim key As Variant
                    Dim FormatString As String
                    Dim OriginalValue As String
                    Dim FormattedValue As String
                    For Each key In dateColumns.Keys
                        FormatString = dateColumns(key)
                        OriginalValue = csvRow(1, key)
                        ColumnName = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_NAME & DEFINITE_TABLE_FIRST_COL + key - 1).value
                        
                        If IsDate(OriginalValue) Then
                            FormattedValue = Format(OriginalValue, FormatString)
                            If FormattedValue <> OriginalValue Then
                                Log.ERROR (Replace(Replace(Replace(ERROR_COLUMN_DATE_FORMAT, "%{fileName}", fileOverView), "%{row}", rowNumCSV), "%{column}", ColumnName))
                            End If
                        Else
                            Log.ERROR (Replace(Replace(Replace(ERROR_COLUMN_DATE_FORMAT, "%{fileName}", fileOverView), "%{row}", rowNumCSV), "%{column}", ColumnName))
                        End If
                    Next key
                    If rowNumCSV > 1000 Then
                        Call updateStatusProcess(rowNum, STATUS_PROCESS_COMPLETED_OK)
                        GoTo NextIterationCB
                    Else
                        rowNumCSV = rowNumCSV + 1
                    End If
                Next csvRow
                '9 Validate maxRecord
                fileRecordQuantity = getFileLine(filePath)
                IsValidateMaxRecord = validateMaxRecord(fileRecordQuantity, maxRecordRule, fileOverView, flagRecordSize)
                
                Call updateStatusProcess(rowNum, STATUS_PROCESS_COMPLETED_OK)
                currentCheckingRow = currentCheckingRow + 1
           End If
NextIterationCB:
        Next CB
    End If
End Sub

'Select file handle
Sub btnSelectFile_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd

    .AllowMultiSelect = False
    .Title = "Please select the file to kill his non colored cells"
    .Filters.Add "CSV", "*.csv"
    .Filters.Add "TSV", "*.tsv"

    If .Show = True Then
        rowClicked = Split(ActiveSheet.Shapes(Application.Caller).Name)(1)
        txtFileName = .SelectedItems(1)
        Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowClicked).value = txtFileName
        Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowClicked).Font.Color = vbBlack
        Common.dataCheckSheet.Range(COL_DATA_CHECK_STATUS_CHECK & rowClicked).value = STATUS_PROCESS_INIT_FILE
    End If

    End With
End Sub

'Get conditions handle
Sub btnGetextractionList_Click()
    mbResult = MsgBox(CLEAR_CONFIRM_MSG, vbYesNo)
    
    If mbResult = vbYes Then
        clearContent
        
        Dim errorRows As String
        errorRows = ""
        
        lastRow = Common.lastRowFileListSheet()
        If lastRow > 104 Then
            lastRow = 104
        End If
        addIndex = 6
        errorIndex = 0
        
        'huynnp
        Dim stringPattern As String
        stringPattern = ""
        Dim errorPatterns As String
        errorPatterns = ""
        
        For i = 5 To lastRow
            
            rowNum = Common.fileListSheet.Range(COL_FILE_LIST_ROW_NUM & i).value
            name_pattern = Common.fileListSheet.Range(COL_FILE_LIST_NAME_PATTERN & i).value
            fileOverView = Common.fileListSheet.Range(COL_FILE_LIST_FILE_NAME & i).value
            maxRecord = Common.fileListSheet.Range(COL_FILE_LIST_MAX_QUANTITY_RECORD & i).value
            maxFileSize = Common.fileListSheet.Range(COL_FILE_LIST_MAX_FILE_SIZE & i).value
            flagRecordSize = Common.fileListSheet.Range(COL_FILE_LIST_FLAG_RECORD_SIZE & i).value
            flagDuplicateNamePattern = False
            
            If IsEmpty(fileOverView) = False Then
                If (IsEmpty(maxRecord) Or maxRecord = 0 Or IsEmpty(maxFileSize) Or maxFileSize = 0) Then
                    If flagRecordSize = "全件" Then
                        errorRows = errorRows & " " & rowNum
                    Else
                    'huynnp
                        arrayPattern = Split(stringPattern, ";")
                        For j = 0 To UBound(arrayPattern)
                            If arrayPattern(j) = Split(name_pattern, "<")(0) Then
                                errorPatterns = errorPatterns & " " & name_pattern & ":" & rowNum
                                flagDuplicateNamePattern = True
                                'Exit Sub
                            End If
                        Next j
                        If flagDuplicateNamePattern = False Then
                            stringPattern = stringPattern & ";" & Split(name_pattern, "<")(0)
                        End If
                                                
                        If Trim(errorPatterns) = "" Then
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_NO & addIndex).value = rowNum
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & addIndex).value = Split(name_pattern, "<")(0)
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & addIndex).value = fileOverView
                            If flagRecordSize = "全件" Then
                                Common.dataCheckSheet.Range(COL_DATA_CHECK_MAX_QUANTITY_RECORD & addIndex).value = maxRecord
                                Common.dataCheckSheet.Range(COL_DATA_CHECK_MAX_FILE_SIZE & addIndex).value = maxFileSize
                            End If
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_SAVE & addIndex).value = i
                        End If
                        addIndex = addIndex + 1
                    End If
                ElseIf IsEmpty(name_pattern) Then
                    errorRows = errorRows & " " & rowNum
                Else
                'huynnp
                    arrayPattern = Split(stringPattern, ";")
                    For j = 0 To UBound(arrayPattern)
                        If arrayPattern(j) = Split(name_pattern, "<")(0) Then
                            errorPatterns = errorPatterns & " " & name_pattern & ":" & rowNum
                            flagDuplicateNamePattern = True
                            'Exit Sub
                        End If
                    Next j
                    If flagDuplicateNamePattern = False Then
                        stringPattern = stringPattern & ";" & Split(name_pattern, "<")(0)
                    End If
                    
                    If Trim(errorPatterns) = "" Then
                        Common.dataCheckSheet.Range(COL_DATA_CHECK_NO & addIndex).value = rowNum
                        Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & addIndex).value = Split(name_pattern, "<")(0)
                        Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & addIndex).value = fileOverView
                        If flagRecordSize = "全件" Then
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_MAX_QUANTITY_RECORD & addIndex).value = maxRecord
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_MAX_FILE_SIZE & addIndex).value = maxFileSize
                        End If
                        Common.dataCheckSheet.Range(COL_DATA_CHECK_SAVE & addIndex).value = i
                    End If

                    addIndex = addIndex + 1
                End If
            ElseIf (IsEmpty(fileOverView)) And checkRowEmpty(i) Then
                errorRows = errorRows & " " & rowNum
            End If
        Next i
'huynnp
        If IsEmpty(Trim(errorRows)) = False And Trim(errorRows) <> "" Then
            Validation.extractionEmptyFileSize (errorRows)
        Else
            Validation.duplicateNamePattern (errorPatterns)
        End If
        
    End If
End Sub

' Cancel checking handle
Sub btnCancel_Click()
    Call updateStatusProcess(currentCheckingRow, STATUS_PROCESS_STOP)
    Exit Sub
End Sub

'Clear without confirmation
Private Sub clearContent()
    lastRow = Common.lastRowDataCheckSheet()
    If lastRow > 5 Then
        Common.dataCheckSheet.Range(DATA_CHECK_RANGE & lastRow).ClearContents
    End If
    For Each CB In ActiveSheet.CheckBoxes
      CB.value = 0
    Next CB
End Sub

'Create Definite Sheet
Sub btnCreateDefiniteSheet_Click()
    rowClicked = Split(ActiveSheet.Shapes(Application.Caller).Name, " ")(1)
    sheetNamePattern = Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & rowClicked).value
    sheetNameRule = Common.fileListSheet.Range(COL_FILE_LIST_NAME_PATTERN & rowClicked - 1).value
    'Check exists sheets
    If sheetNamePattern = "" Then
        Exit Sub
    End If
    If checkExistsSheet(sheetNamePattern) Then
        MsgBox CHECK_EXISTS_SHEET
        Exit Sub
    End If
    'Copy sheet template
    Sheets(TABLE_DEFINITE_TEMPLATE).Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = sheetNamePattern
    Sheets(sheetNamePattern).Range("B2").value = Split(sheetNameRule, "<")(0) & "カラム定義シート"
End Sub

'Check exists sheet
Private Function checkExistsSheet(ByVal sheetName As String)
    For Each sh In ActiveWorkbook.Sheets
        If sh.Name = sheetName Then
            checkExistsSheet = True
            Exit Function
        End If
    Next sh
    checkExistsSheet = False
End Function
'Check row empty
Private Function checkRowEmpty(ByVal index As Integer)
    D = Common.fileListSheet.Range("D" & index).value
    E = Common.fileListSheet.Range("E" & index).value
    g = Common.fileListSheet.Range("G" & index).value
    H = Common.fileListSheet.Range("H" & index).value
    i = Common.fileListSheet.Range("i" & index).value
    j = Common.fileListSheet.Range("J" & index).value
    K = Common.fileListSheet.Range("K" & index).value
    L = Common.fileListSheet.Range("L" & index).value
    M = Common.fileListSheet.Range("M" & index).value
    n = Common.fileListSheet.Range("n" & index).value
    O = Common.fileListSheet.Range("O" & index).value
    P = Common.fileListSheet.Range("P" & index).value
    Q = Common.fileListSheet.Range("Q" & index).value
    r = Common.fileListSheet.Range("r" & index).value
    s = Common.fileListSheet.Range("s" & index).value
    T = Common.fileListSheet.Range("T" & index).value
    U = Common.fileListSheet.Range("U" & index).value
    v = Common.fileListSheet.Range("v" & index).value
    W = Common.fileListSheet.Range("W" & index).value
    X = Common.fileListSheet.Range("X" & index).value
    Y = Common.fileListSheet.Range("Y" & index).value
    Z = Common.fileListSheet.Range("Z" & index).value
    AA = Common.fileListSheet.Range("AA" & index).value
    AB = Common.fileListSheet.Range("AB" & index).value
    AC = Common.fileListSheet.Range("AC" & index).value
    AD = Common.fileListSheet.Range("AD" & index).value
    AE = Common.fileListSheet.Range("AE" & index).value
    AF = Common.fileListSheet.Range("AF" & index).value
    AG = Common.fileListSheet.Range("AG" & index).value
    AH = Common.fileListSheet.Range("AH" & index).value
    AI = Common.fileListSheet.Range("AI" & index).value
    AJ = Common.fileListSheet.Range("AJ" & index).value
    AK = Common.fileListSheet.Range("AK" & index).value
    AL = Common.fileListSheet.Range("AL" & index).value
    AM = Common.fileListSheet.Range("AM" & index).value
    
    If IsEmpty(D) = False Or IsEmpty(E) = False Or IsEmpty(g) = False Or IsEmpty(H) = False Or IsEmpty(i) = False Or IsEmpty(j) = False Or IsEmpty(K) = False Or IsEmpty(L) = False Or IsEmpty(M) = False Or IsEmpty(n) = False Or IsEmpty(O) = False Or IsEmpty(P) = False Or IsEmpty(Q) = False Or IsEmpty(r) = False Or IsEmpty(s) = False Or IsEmpty(T) = False Or IsEmpty(U) = False Or IsEmpty(v) = False Then
        checkRowEmpty = True
        Exit Function
    ElseIf IsEmpty(AD) = False Or IsEmpty(AE) = False Or IsEmpty(AF) = False Or IsEmpty(AG) = False Or IsEmpty(AH) = False Or IsEmpty(AI) = False Or IsEmpty(AJ) = False Or IsEmpty(AK) = False Or IsEmpty(AL) = False Or IsEmpty(AM) = False Or IsEmpty(W) = False Or IsEmpty(X) = False Or IsEmpty(Y) = False Or IsEmpty(Z) = False Or IsEmpty(AA) = False Or IsEmpty(AB) = False Or IsEmpty(AC) = False Then
        checkRowEmpty = True
        Exit Function
    End If
    checkRowEmpty = False
    
End Function



'Read CSV by binary
'@param strLine                   Line to check
'@param strSeperatedChar   Seperated character
'@param intNoOfCol             Defined number of column of file
Function checkValidateSeperatedCharacter(ByVal fullFileName As String, ByVal strSeperatedChar As String, ByVal noOfCol As Integer, ByVal fileOverView As String, ByVal startRowData As Integer)
    
    Dim intUnit As Integer
    Dim my_string As String
    Dim vntLines As Variant
    
    intUnit = FreeFile
    Open fullFileName For Binary Access Read As #intUnit
    my_string = Input(LOF(intUnit), intUnit)
    
    If InStr(my_string, vbCrLf) > 0 Then
        'Window file
        vntLines = Split(my_string, vbCrLf)
        vntLines = getProcessLine(vntLines, startRowData)
        
    ElseIf InStr(my_string, vbCr) > 0 Then
        'MAC file
        vntLines = Split(my_string, vbCr)
        vntLines = getProcessLine(vntLines, startRowData)
        
    Else
        'Unix file
        vntLines = Split(my_string, vbLf)
        vntLines = getProcessLine(vntLines, startRowData)
        
    End If
    Close intUnit
    checkValidateSeperatedCharacter = True
    'Checking
    For i = (LBound(vntLines) + startRowData) To UBound(vntLines)
        If checkSeperatedCharacter(vntLines(i), strSeperatedChar, noOfCol) = False Then
            checkValidateSeperatedCharacter = False
            Log.ERROR (Replace(Replace(ERROR_SEPERATED_CHARACTER, "%{fileName}", fileOverView), "%{row}", i + 1 + startRowData))
        End If
    Next
End Function

' Read and parse CSV file
' Return array 2d if parse success, otherwise return Null
Function readCSV(ByVal filePath As String, ByVal separater As String, ByVal limitLine As Long, ByVal columnCount As Integer, ByVal fileOverView As String, ByVal startRow As Integer, ByVal currentRowFile As Integer) As Variant
    Dim textData As String
    Dim textRow As String
    Dim fileNo As Integer
    Dim csv As Variant
    Dim csvArray As Variant
    Dim errorTmp As String
    Dim lineIndex As Integer
    
    ReDim csvArray(1 To limitLine) As Variant
    arrayIndex = 1
    lineIndex = 1
    IsValid = True
    csv = Null
    SetCSVUtilsAnyErrorIsFatal (True)
    
    fileNo = FreeFile
    
    Open filePath For Input As #fileNo
    
    If columnCount = 1 Then
        parttern = "^\""([^\""]|\""\"")*\""$"
                   
    ElseIf columnCount = 2 Then
        parttern = "^\""([^\""]|\""\"")*\""" & separater & "\""([^\""]|\""\"")*\""$"
    Else
        parttern = "^\""([^\""]|\""\"")*\""" & separater
    
        For i = 1 To columnCount - 2
            parttern = parttern & "\""([^\""]|\""\"")*\""" & separater
        Next i
        parttern = parttern & "\""([^\""]|\""\"")*\""$"
    End If
    
    Dim regEx As New RegExp
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = parttern
    End With
    
    Do While Not EOF(fileNo) And arrayIndex <= limitLine
        
        Line Input #fileNo, textRow
        
        If textData = "" Then
            textData = textRow
        Else
            textData = textData & textRow
        End If
        
        'try parse CSV
        csv = ParseCSVToArray(textData, separater, fileOverView, arrayIndex)
        
        'parse csv error
        If IsNull(csv) Or UBound(csv) = -1 Then
            errorTmp = arrayIndex
        'parse ok
        Else
            'Validate quote
            If lineIndex > startRow Then
                If validateQuote(textData, regEx) = False Then
                    Log.ERROR (Replace(Replace(ERROR_DOUBLE_QUOTE, "%{fileName}", fileOverView), "%{row}", lineIndex))
                End If
            End If
            'parse csv success and column matched
            If UBound(csv, 2) = columnCount Then
                If lineIndex > startRow Then
                    csvArray(arrayIndex) = csv
                    arrayIndex = arrayIndex + 1
                End If
                textData = ""
                errorTmp = ""
            'parse csv ok but column not match
            ElseIf UBound(csv, 2) <> columnCount Then
                'If parseErrorRows.exists(errorTmp) = False Then
                    'If errorTmp <> "" Then
                        'parseErrorRows.Add errorTmp, errorTmp
                    'End If
                'End If
                If parseErrorRows.exists(lineIndex) = False Then
                    parseErrorRows.Add lineIndex, lineIndex
                End If
                errorTmp = ""
                textData = ""
            End If
        End If
        lineIndex = lineIndex + 1
    Loop
    Close #fileNo
    
    'check has error
    If errorTmp <> "" Then
        If parseErrorRows.exists(errorTmp) = False Then
            parseErrorRows.Add errorTmp, errorTmp
        End If
        errorTmp = ""
    End If
    
    'Log if has error
    If textData <> "" Or parseErrorRows.Count > 0 Then
        IsValid = False
        For Each key In parseErrorRows.Keys
            Log.ERROR (Replace(Replace(ERROR_SEPERATED_CHARACTER, "%{fileName}", fileOverView), "%{row}", parseErrorRows(key) + startRow))
        Next
        Call updateStatusProcess(currentRowFile, STATUS_PROCESS_COMPLETED_NOK)
    End If
    
    'trim result array
    If IsValid = True Then
        Dim csvTmp As Variant
        If arrayIndex > 1 Then
            If (arrayIndex < limitLine) Then
                ReDim csvTmp(1 To arrayIndex - 1) As Variant
                For i = 1 To arrayIndex - 1
                    csvTmp(i) = csvArray(i)
                Next i
            Else
                csvTmp = csvArray
            End If
        End If
        'return value
        readCSV = csvTmp
        'free memory
        'Erase csvTmp
    Else
        readCSV = Null
        Call updateStatusProcess(currentRowFile, STATUS_PROCESS_COMPLETED_NOK)
    End If
    
    'free memory
    Erase csvArray
    parseErrorRows.RemoveAll
End Function

Sub updateStatusProcess(ByVal rowNum As Integer, ByVal statusProcess As String)
    Common.dataCheckSheet.Range(COL_DATA_CHECK_STATUS_CHECK & rowNum).value = statusProcess
    Common.dataCheckSheet.Range(COL_DATA_CHECK_DATE & rowNum).value = Format(Now, "yyyy-mm-dd hh:mm:ss")
End Sub
