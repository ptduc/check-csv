Attribute VB_Name = "Log"
'write log at current directory
Private Sub writeLog(logType As String, data As String)
    Dim filePath As String
    filePath = ActiveWorkbook.path & "/logs.txt"
    Open filePath For Append As #2
    data = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - [" & logType & "]: " & data
    Print #2, data
    Close #2
End Sub

Sub createFileLog()
    Dim F As Integer
    F = FreeFile
    Open ActiveWorkbook.path & "/logs.txt" For Output As #F
    Close #F
End Sub

Sub ERROR(data As String)
    writeLog "ERROR", data
End Sub

Sub warn(data As String)
    writeLog "WARN", data
End Sub

Sub info(data As String)
    writeLog "INFO", data
End Sub
