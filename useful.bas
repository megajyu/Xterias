Attribute VB_Name = "useful"
Option Explicit

'判定==========
'説明)指定されたCellで、テキストが存在するか判定する
Public Function IsNoTextCell(ByRef vCell As Range) As Boolean
    IsNoTextCell = True
    If vCell Is Nothing Then Exit Function
    If vCell.Text <> "" Then IsNoTextCell = False
End Function
'説明)(タブなどを除いた)テキストが存在するか判定する
Public Function IsNoVisibleText(ByRef vStr As String) As Boolean
    IsNoVisibleText = True
    
    Dim chkStr As String: chkStr = vStr
    chkStr = Replace(chkStr, vbTab, "") 'タブを消去
    chkStr = Replace(chkStr, " ", "") '半角スペースを消去
    chkStr = Replace(chkStr, "　", "") '全角スペースを消去
    chkStr = Replace(chkStr, vbCrLf, "") '改行を消去
    chkStr = Replace(chkStr, vbCr, "") 'キャリッジリターンを消去
    chkStr = Replace(chkStr, vbLf, "") 'LFを消去
    
    If chkStr <> "" Then IsNoVisibleText = False
    
End Function


'正規表現==========
'説明)正規表現でマッチしているか判定する
Public Function RegChk(ByVal vStr As String, ByVal vPattern) As Boolean
    With CreateObject("VBScript.RegExp")
        .Pattern = vPattern
        RegChk = .test(vStr)
    End With
End Function

'説明)正規表現でマッチした箇所を置き換える
Public Function RegReplace(ByVal vFrom As String, ByVal vPattern, ByVal vTo As String) As String
    Dim ret As String
    Dim oMatchs As Variant
    Dim oMatch As Variant
    
    With CreateObject("VBScript.RegExp")
        .Pattern = vPattern
        .Global = True
        ret = .Replace(vFrom, vTo)
    End With
    
    RegReplace = ret
End Function

'説明)マッチした文言の配列を取得する
Public Function RegMatch(ByVal vFrom As String, ByVal vPattern As String) As Variant
    Dim ret() As String
    Dim cnt As Long
    Dim oMatchs As Variant
    Dim oMatch As Variant
    
    With CreateObject("VBScript.RegExp")
        .Pattern = vPattern
        .Global = True
        
        cnt = 0
        Set oMatchs = .Execute(vFrom)
        For Each oMatch In oMatchs
            ReDim Preserve ret(cnt)
            ret(cnt) = oMatch.Value
            cnt = cnt + 1
        Next
    End With
    
    RegMatch = ret
    
End Function

'説明)サブマッチした文言の配列を取得する
Public Function RegSubMatch(ByVal vFrom As String, ByVal vPattern As String) As Variant
    Dim ret() As String
    Dim cnt As Long
    Dim tmp As Variant
    Dim oMatchs As Variant
    Dim oMatch As Variant
    
    With CreateObject("VBScript.RegExp")
        .Pattern = vPattern
        .Global = True
        
        cnt = 0
        Set oMatchs = .Execute(vFrom)
        For Each oMatch In oMatchs
            For Each tmp In oMatch.SubMatches
                ReDim Preserve ret(cnt)
                ret(cnt) = oMatch.Value
                cnt = cnt + 1
            Next
        Next
    End With
    
    '該当が無い場合は、空文字を返す
    If cnt = 0 Then
        ReDim ret(1)
        ret(0) = vbNullString
    End If
    
    RegSubMatch = ret
    
End Function

'文字列操作==========
'説明)後ろから調べて最初に第２引数が見つかった位置の一つ後ろから終わりまでを切り出す
'見つからなかった場合は、空文字を返す
Public Function EsRight(ByVal vStr As String, ByVal vSearchStr As String) As String
    Dim Idx As Long
    Idx = InStrRev(vStr, vSearchStr)
    
    '見つかった位置の一つ後ろから、文字列長ー見つかった位置
    EsRight = Mid(vStr, Idx + 1, Len(vStr) - Idx)
    
    '見つからなかった場合は、空文字を返す
    If Idx = 0 Then
        EsRight = vbNullString
    End If
End Function

'説明)前から調べて最初に第２引数が見つかった位置の一つ前までを切り出す
'見つからなかった場合は、空文字を返す
Public Function EsLeft(ByVal vStr As String, ByVal vSearchStr As String) As String
    Dim Idx As Long
    Idx = InStr(vStr, vSearchStr)
    
    '見つかった位置の一つ後ろから、文字列長ー見つかった位置
    EsLeft = Left(vStr, Idx - 1)
    
    '見つからなかった場合は、空文字を返す
    If Idx = 0 Then
        EsLeft = vbNullString
    End If
    
End Function

'説明)エラーになりにくいSplit
Public Function EsSplit(ByVal vStr As String, ByVal vDelimiter As String, Optional ByVal vIdx As Long, Optional vDefault As String) As Variant
    Dim ret As Variant  '返す値が、「文字列」「配列」の２種類あるため
    
    '文字列の指定なし
    '→vIdxが指定されている:vbNullString:文字が欲しいので
    '→vIdxが指定されていない:空配列:for eachで使いたいので
    If vStr = vbNullString Then
        If Not vIdx = vbNullString Then
            EsSplit = vbNullString
        Else
            EsSplit = Array("")
        End If
        Exit Function
    End If
    
    '文字列を分割(配列)
    ret = Split(vStr, vDilimiter)
    
    'vIdxが指定されている場合は、インデックスが指す値を返す
    Dim Idx As Long
    If Not vIdx = vbNullString Then
        Idx = CLng(vIdx)
        If Idx <= UBoung(ret) Then
            ret = ret(Idx)
        Else
            '指定なし:vbNullString、指定あり:指定された値
            ret = vDefault
        End If
    End If
    
    EsSplit = ret
End Function

'日付==========
'説明)現在の日付(YYYYMMDDhhmmss)を返す
'20161231125959
Public Function YYYYMMDDhhmmss() As String
    YYYYMMDDhhmmss = Year(Now) & _
                                        Right("00" & Month(Now), 2) & _
                                        Right("00" & Day(Now), 2) & _
                                        Right("00" & Hour(Now), 2) & _
                                        Right("00" & Minute(Now), 2) & _
                                        Right("00" & Second(Now), 2)
End Function

'説明)現在の日付(乱数付き)(YYYYMMDDhhmmssRRR)を返す
'20161231125959034
Public Function YYYYMMDDhhmmssRRR() As String
    YYYYMMDDhhmmssRRR = YYYYMMDDhhmmss & Right("000" & CInt(Rnd() * 100), 3)
End Function

'セル選択状態の判定==========
'説明)渡されたRangeが、どのような状態で有るか判定する
Public Function GetRangeType(ByRef vRange As Range) As String
 '途中
End Function

