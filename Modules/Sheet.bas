Attribute VB_Name = "Sheet"
Function lastHasData(lRowColCell As Long, _
                    Optional sSheet As String, _
                    Optional sRange As String)
'Find the last row, column, or cell using the Range.Find method
'lRowColCell: 1=Row, 2=Col, 3=Cell

Dim lRow As Long
Dim lCol As Long
Dim wsFind As Worksheet
Dim rFind As Range

    'Default to ActiveSheet if none specified
    On Error GoTo ErrExit
    
    If sSheet = "" Then
        Set wsFind = ActiveSheet
    Else
        Set wsFind = Worksheets(sSheet)
    End If

    'Default to all cells if range no specified
    If sRange = "" Then
        Set rFind = wsFind.Cells
    Else
        Set rFind = wsFind.Range(sRange)
    End If
    
    On Error GoTo 0

    Select Case lRowColCell
    
        Case 1 'Find last row
            On Error Resume Next
            lastHasData = rFind.Find(What:="*", _
                            After:=rFind.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
            On Error GoTo 0

        Case 2 'Find last column
            On Error Resume Next
            lastHasData = rFind.Find(What:="*", _
                            After:=rFind.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0

        Case 3 'Find last cell by finding last row & col
            On Error Resume Next
            lRow = rFind.Find(What:="*", _
                           After:=rFind.Cells(1), _
                           LookAt:=xlPart, _
                           LookIn:=xlFormulas, _
                           SearchOrder:=xlByRows, _
                           SearchDirection:=xlPrevious, _
                           MatchCase:=False).Row
            On Error GoTo 0

            On Error Resume Next
            lCol = rFind.Find(What:="*", _
                            After:=rFind.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0

            On Error Resume Next
            lastHasData = wsFind.Cells(lRow, lCol).Address(False, False)
            'If lRow or lCol = 0 then entire sheet is blank, return "A1"
            If Err.Number > 0 Then
                lastHasData = rFind.Cells(1).Address(False, False)
                Err.Clear
            End If
            On Error GoTo 0

    End Select
    
    Exit Function
    
ErrExit:

    MsgBox "Error setting the worksheet or range."

End Function

