Attribute VB_Name = "FileEncoding"
Option Explicit
'****************************************************************************
' 機能名    : Module1.bas
' 機能説明  : 文字コード判定
' 備考      :
' 著作権    : Copyright(C) 2008 - 2009 のん All rights reserved
' ---------------------------------------------------------------------------
' 使用条件  : このサイトの内容を使用(流用/改変/転載/等全て)した成果物を不特定
'           : 多数に公開/配布する場合は、このサイトを参考にした旨を記述してく
'           : ださい。(例)WEBページやReadMeにリンクを貼ってください
' ---------------------------------------------------------------------------
'****************************************************************************
Private Const JUDGEFIX = 9999       '文字コード決定％
Private Const JUDGESIZEMAX = 1000   '文字コード判定バイト数
Private Const SingleByteWeight = 1  '１バイト　文字コードの一致重み
Private Const Multi_ByteWeight = 2  '複数バイト文字コードの一致重み
Private Enum JISMODE                'JISコードのモード
    ctrl = 0                        '制御コード
    asci = 1                        'ASCII
    roma = 2                        'JISローマ字
    kana = 3                        'JISカナ（半角カナ）
    kanO = 4                        '旧JIS漢字 (1978)
    kanN = 5                        '新JIS漢字 (1983/1990)
    kanH = 6                        'JIS補助漢字
End Enum

'----文字コード判定
' 関数名    : JudgeCode
' 返り値    : 判定結果文字コード名
' 引き数    : bytCode : 判定文字データ
' 機能説明  : 文字コードを判定する
' 備考      :
Public Function JudgeCode(ByRef bytCode() As Byte) As String
    JudgeCode = "SJIS"
    Dim lngSJIS As Long
    Dim lngJIS As Long
    Dim lngEUC As Long
    Dim lngUNI As Long
    Dim lngUTF7 As Long
    Dim lngUTF8 As Long
    
    lngJIS = JudgeJIS(bytCode, True): Debug.Print "JIS :" & lngJIS
    If lngJIS >= JUDGEFIX Then JudgeCode = "JIS": Exit Function
    
    lngUNI = JudgeUNI(bytCode, True): Debug.Print "UNI :" & lngUNI
    If lngUNI >= JUDGEFIX Then JudgeCode = "UNICODE": Exit Function
    
    lngUTF8 = JudgeUTF8(bytCode, True): Debug.Print "UTF8:" & lngUTF8
    If lngUTF8 >= JUDGEFIX Then JudgeCode = "UTF8": Exit Function

    lngUTF7 = JudgeUTF7(bytCode, True): Debug.Print "UTF7:" & lngUTF7
    If lngUTF7 >= JUDGEFIX Then JudgeCode = "UTF7": Exit Function
    
    lngSJIS = JudgeSJIS(bytCode, True): Debug.Print "SJIS:" & lngSJIS
    If lngSJIS >= JUDGEFIX Then JudgeCode = "SJIS": Exit Function
    
    lngEUC = JudgeEUC(bytCode, True): Debug.Print "EUC :" & lngEUC
    If lngEUC >= JUDGEFIX Then JudgeCode = "EUC": Exit Function
    Debug.Print "--------"

    If lngSJIS >= lngSJIS And lngSJIS >= lngUNI And lngSJIS >= lngJIS And _
       lngSJIS >= lngUTF7 And lngSJIS >= lngUTF8 And lngSJIS >= lngEUC Then
        JudgeCode = "SJIS"
        Exit Function
    End If
    
    If lngUNI >= lngSJIS And lngUNI >= lngUNI And lngUNI >= lngJIS And _
       lngUNI >= lngUTF7 And lngUNI >= lngUTF8 And lngUNI >= lngEUC Then
        JudgeCode = "UNICODE"
        Exit Function
    End If
    
    If lngJIS >= lngSJIS And lngJIS >= lngUNI And lngJIS >= lngJIS And _
       lngJIS >= lngUTF7 And lngJIS >= lngUTF8 And lngJIS >= lngEUC Then
        JudgeCode = "JIS"
        Exit Function
    End If
    
    If lngUTF7 >= lngSJIS And lngUTF7 >= lngUNI And lngUTF7 >= lngJIS And _
       lngUTF7 >= lngUTF7 And lngUTF7 >= lngUTF8 And lngUTF7 >= lngEUC Then
        JudgeCode = "UTF7"
        Exit Function
    End If
    
    If lngUTF8 >= lngSJIS And lngUTF8 >= lngUNI And lngUTF8 >= lngJIS And _
       lngUTF8 >= lngUTF7 And lngUTF8 >= lngUTF8 And lngUTF8 >= lngEUC Then
        JudgeCode = "UTF8"
        Exit Function
    End If
    
    If lngEUC >= lngSJIS And lngEUC >= lngUNI And lngEUC >= lngJIS And _
       lngEUC >= lngUTF7 And lngEUC >= lngUTF8 And lngEUC >= lngEUC Then
        JudgeCode = "EUC"
        Exit Function
    End If
    
End Function

'----SJIS関係
' 関数名    : JudgeSJIS
' 返り値    : 判定結果確率（％）
' 引き数    : bytCode : 判定文字データ
'           : fixFlag : 確定判断有無
' 機能説明  : SJISの文字コード判定(可能性)確率を計算する
' 備考      :
Private Function JudgeSJIS(ByRef bytCode() As Byte, _
                           Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        '81-9F,E0-EF(1バイト目)
        If (bytCode(i) >= &H81 And bytCode(i) <= &H9F) Or _
           (bytCode(i) >= &HE0 And bytCode(i) <= &HEF) Then
           If i <= UBound(bytCode) - 1 Then
                '40-7E,80-FC(2バイト目)
                If (bytCode(i + 1) >= &H40 And bytCode(i + 1) <= &H7E) Or _
                   (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HFC) Then
                    lngFit = lngFit + (2 * Multi_ByteWeight)
                    i = i + 1
                End If
            End If
        
        'A1-DF(1バイト目)
        ElseIf (bytCode(i) >= &HA1 And bytCode(i) <= &HDF) Then
            lngFit = lngFit + (1 * SingleByteWeight)
        
        '20-7E(1バイト目)
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            lngFit = lngFit + (1 * SingleByteWeight)
        
        '00-1F, 7F(1バイト目)
        ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                bytCode(i) = &H7F Then
            lngFit = lngFit + (1 * SingleByteWeight)
        End If
    Next i
    JudgeSJIS = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----JIS関係
' 関数名    : JudgeJIS
' 返り値    : 判定結果確率（％）
' 引き数    : bytCode : 判定文字データ
'           : fixFlag : 確定判断有無
' 機能説明  : JISの文字コード判定(可能性)確率を計算する
' 備考      :
Private Function JudgeJIS(ByRef bytCode() As Byte, _
                          Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngMode As JISMODE
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        '1B(1バイト目)
        If bytCode(i) = &H1B Then
           If i <= UBound(bytCode) - 2 Then
                '28 42(2・3バイト目)
                If bytCode(i + 1) = &H28 And bytCode(i + 1) <= &H42 Then
                    lngMode = asci
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '28 4A(2・3バイト目)
                If bytCode(i + 1) = &H28 And bytCode(i + 1) <= &H4A Then
                    lngMode = roma
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '28 49(2・3バイト目)
                If bytCode(i + 1) = &H28 And bytCode(i + 1) <= &H49 Then
                    lngMode = kana
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '24 40(2・3バイト目)
                If bytCode(i + 1) = &H24 And bytCode(i + 1) <= &H40 Then
                    lngMode = kanO
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '24 42(2・3バイト目)
                If bytCode(i + 1) = &H24 And bytCode(i + 1) <= &H42 Then
                    lngMode = kanN
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '24 44(2・3バイト目)
                If bytCode(i + 1) = &H24 And bytCode(i + 1) <= &H44 Then
                    lngMode = kanH
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
            End If
        Else
            Select Case lngMode
            Case ctrl, asci, roma
                '00-1F,7F
                If (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                    bytCode(i) = &H7F Then
                    lngFit = lngFit + (1 * SingleByteWeight)
                End If
                '20-7E
                If (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
                    lngFit = lngFit + (1 * SingleByteWeight)
                End If
            Case kana
                '21-5F
                If (bytCode(i) >= &H21 And bytCode(i) <= &H5F) Then
                    lngFit = lngFit + (1 * SingleByteWeight)
                End If
            Case kanO, kanN, kanH
               If i <= UBound(bytCode) - 1 Then
                    '21-7E
                    If (bytCode(i) >= &H21 And bytCode(i) <= &H7E) And _
                       (bytCode(i - 1) >= &H21 And bytCode(i - 1) <= &H7E) Then
                        lngFit = lngFit + (2 * Multi_ByteWeight)
                        i = i + 1
                    End If
                End If
            End Select
        End If
    Next i
    JudgeJIS = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----EUC関係
' 関数名    : JudgeEUC
' 返り値    : 判定結果確率（％）
' 引き数    : bytCode : 判定文字データ
'           : fixFlag : 確定判断有無
' 機能説明  : EUCの文字コード判定(可能性)確率を計算する
' 備考      :
Private Function JudgeEUC(ByRef bytCode() As Byte, _
                          Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        '8E(1バイト目) + A1-DF(2バイト目)
        If bytCode(i) = &H8E Then
            If i <= UBound(bytCode) - 1 Then
                If bytCode(i + 1) >= &HA1 And bytCode(i + 1) <= &HDF Then
                    lngFit = lngFit + (2 * Multi_ByteWeight)
                    i = i + 1
                End If
            End If
        
        '8F(1バイト目) + A1-0xFE(2・3バイト目)
        ElseIf bytCode(i) = &H8F Then
            If i <= UBound(bytCode) - 2 Then
                If (bytCode(i + 1) >= &HA1 And bytCode(i + 1) <= &HFE) And _
                   (bytCode(i + 2) >= &HA1 And bytCode(i + 2) <= &HFE) Then
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                End If
            End If
        
        'A1-FE(1バイト目) + A1-FE(2バイト目)
        ElseIf bytCode(i) >= &HA1 And bytCode(i) <= &HFE Then
            If i <= UBound(bytCode) - 1 Then
                If bytCode(i + 1) >= &HA1 And bytCode(i + 1) <= &HFE Then
                    lngFit = lngFit + (2 * Multi_ByteWeight)
                    i = i + 1
                End If
            End If
            
        '20-7E(1バイト目)
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            lngFit = lngFit + (1 * SingleByteWeight)

        '00-1F, 7F(1バイト目)
        ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                bytCode(i) = &H7F Then
            lngFit = lngFit + (1 * SingleByteWeight)
        End If
    Next i
    JudgeEUC = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----UNICODE関係
' 関数名    : JudgeUNI
' 返り値    : 判定結果確率（％）
' 引き数    : bytCode : 判定文字データ
'           : fixFlag : 確定判断有無
' 機能説明  : UTF16の文字コード判定(可能性)確率を計算する
' 備考      :
Private Function JudgeUNI(ByRef bytCode() As Byte, _
                          Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        If fixFlag Then
            'BOM
            If bytCode(i) = &HFF Then
                If i <= UBound(bytCode) - 1 Then
                    If bytCode(i + 1) = &HFE Then
                        JudgeUNI = JUDGEFIX
                        Exit Function
                    End If
                End If
            End If
            '半角の証
            'If bytCode(i) = &H0 Then
            '    JudgeUNI = JUDGEFIX
            '    Exit Function
            'End If
        End If
        
        If i <= UBound(bytCode) - 1 Then
            '00(2バイト目)
            If (bytCode(i + 1) = &H0) Then
                '00-FF(1バイト目)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            '01-33(2バイト目)
            ElseIf (bytCode(i + 1) >= &H1 And bytCode(i + 1) <= &H33) Then
                '00-FF(1バイト目)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            '34-4D(2バイト目)
            ElseIf (bytCode(i + 1) >= &H34 And bytCode(i + 1) <= &H4D) Then
                '00-FF(1バイト目)----空き----
                lngFit = 0
                Exit For
            
            '4E-9F(2バイト目)
            ElseIf (bytCode(i + 1) >= &H4E And bytCode(i + 1) <= &H9F) Then
                '00-FF(1バイト目)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            'A0-AB(2バイト目)
            ElseIf (bytCode(i + 1) >= &HA0 And bytCode(i + 1) <= &HAB) Then
                '00-FF(1バイト目)----空き----
                lngFit = 0
                Exit For
            
            'AC-D7(2バイト目)
            ElseIf (bytCode(i + 1) >= &HAC And bytCode(i + 1) <= &HD7) Then
                '00-FF(1バイト目)----ハングル----
                lngFit = 0
                Exit For
            
            'D8-DF(2バイト目)
            ElseIf (bytCode(i + 1) >= &HD8 And bytCode(i + 1) <= &HDF) Then
                '00-FF(1バイト目)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            'E0-F7(2バイト目)
            ElseIf (bytCode(i + 1) >= &HE0 And bytCode(i + 1) <= &HF7) Then
                '00-FF(1バイト目)----外字----
                lngFit = 0
                Exit For
            
            'F8-FF(2バイト目)
            ElseIf (bytCode(i + 1) >= &HF8 And bytCode(i + 1) <= &HFF) Then
                '00-FF(1バイト目)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            End If
            i = i + 1
        End If
    Next i
    JudgeUNI = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----UTF7関係
' 関数名    : JudgeUTF7
' 返り値    : 判定結果確率（％）
' 引き数    : bytCode : 判定文字データ
'           : fixFlag : 確定判断有無
' 機能説明  : UTF7の文字コード判定(可能性)確率を計算する
' 備考      :
Private Function JudgeUTF7(ByRef bytCode() As Byte, _
                           Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngWrk As Long
    Dim str64 As String
    Dim bln64 As Boolean
    str64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim lngUB As Long
    Dim lngBY As Long
    Dim lngXB As Long
    Dim lngXX As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    lngWrk = 0
    
    For i = 0 To lngUB
        '+～-まではBASE64ENCODE
        If bytCode(i) = Asc("+") And bln64 = False Then
            lngWrk = 1
            bln64 = True
        ElseIf bytCode(i) = Asc("-") Then
            If lngWrk <= 0 Then
                lngWrk = lngWrk + 1
                lngFit = lngFit + (lngWrk * SingleByteWeight)
            ElseIf lngWrk = 1 Then
                lngWrk = lngWrk + 1
                lngFit = lngFit + (lngWrk * Multi_ByteWeight)
            ElseIf lngWrk >= 4 And lngXB < 6 And _
                   ((InStr(str64, Chr(bytCode(i - 1))) - 1) And lngXX) = 0 Then
                lngWrk = lngWrk + 1
                lngFit = lngFit + (lngWrk * Multi_ByteWeight)
            End If
            lngWrk = 0
            bln64 = False
        Else
            If bln64 = True Then
                'BASE64ENCODE中
                If InStr(str64, Chr(bytCode(i))) > 0 Then
                    lngBY = Int((lngWrk * 6) / 8)
                    lngXB = (lngWrk * 6) - (lngBY * 8)
                    lngXX = (2 ^ lngXB) - 1
                    lngWrk = lngWrk + 1
                Else
                    lngWrk = 0
                    bln64 = False
                End If
            Else
                '20-7E(1バイト目)
                If (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
                    lngFit = lngFit + (1 * SingleByteWeight)
        
                '00-1F, 7F(1バイト目)
                ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                        bytCode(i) = &H7F Then
                     lngFit = lngFit + (1 * SingleByteWeight)
                End If
            End If
        End If
    Next i
    JudgeUTF7 = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----UTF8関係
' 関数名    : JudgeUTF8
' 返り値    : 判定結果確率（％）
' 引き数    : bytCode : 判定文字データ
'           : fixFlag : 確定判断有無
' 機能説明  : UTF8の文字コード判定(可能性)確率を計算する
' 備考      :
Private Function JudgeUTF8(ByRef bytCode() As Byte, _
                           Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        If fixFlag Then
            'BOM
            If bytCode(i) = &HEF Then
                If i <= UBound(bytCode) - 2 Then
                    If bytCode(i + 1) = &HBB And _
                       bytCode(i + 2) = &HBF Then
                        JudgeUTF8 = JUDGEFIX
                        Exit Function
                    End If
                End If
            End If
        End If
        
        'AND FC(1バイト目) + 80-BF(2-6バイト目)
        If (bytCode(i) And &HFC) = &HFC Then
            If i <= UBound(bytCode) - 5 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) And _
                   (bytCode(i + 3) >= &H80 And bytCode(i + 3) <= &HBF) And _
                   (bytCode(i + 4) >= &H80 And bytCode(i + 4) <= &HBF) And _
                   (bytCode(i + 5) >= &H80 And bytCode(i + 5) <= &HBF) Then
                    lngFit = lngFit + (6 * Multi_ByteWeight)
                    i = i + 5
                End If
            End If
        
        'AND F8(1バイト目) + 80-BF(2-5バイト目)
        ElseIf (bytCode(i) And &HF8) = &HF8 Then
            If i <= UBound(bytCode) - 4 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) And _
                   (bytCode(i + 3) >= &H80 And bytCode(i + 3) <= &HBF) And _
                   (bytCode(i + 4) >= &H80 And bytCode(i + 4) <= &HBF) Then
                    lngFit = lngFit + (5 * Multi_ByteWeight)
                    i = i + 4
                End If
            End If
            
        'AND F0(1バイト目) + 80-BF(2-4バイト目)
        ElseIf (bytCode(i) And &HF0) = &HF0 Then
            If i <= UBound(bytCode) - 3 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) And _
                   (bytCode(i + 3) >= &H80 And bytCode(i + 3) <= &HBF) Then
                    lngFit = lngFit + (4 * Multi_ByteWeight)
                    i = i + 3
                End If
            End If
        
        'AND E0(1バイト目) + 80-BF(2-3バイト目)
        ElseIf (bytCode(i) And &HE0) = &HE0 Then
            If i <= UBound(bytCode) - 2 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) Then
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                End If
            End If
        
        'AND C0(1バイト目) + 80-BF(2バイト目)
        ElseIf (bytCode(i) And &HC0) = &HC0 Then
            If i <= UBound(bytCode) - 1 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) Then
                    lngFit = lngFit + (2 * Multi_ByteWeight)
                    i = i + 1
                End If
            End If

        '20-7E(1バイト目)
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            lngFit = lngFit + (1 * SingleByteWeight)

        '00-1F, 7F(1バイト目)
        ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                bytCode(i) = &H7F Then
            lngFit = lngFit + (1 * SingleByteWeight)
        End If
    Next i
    JudgeUTF8 = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

