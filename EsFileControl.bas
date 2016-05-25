Attribute VB_Name = "EsFileControl"
Option Explicit
'★デリミタの扱いをどうするか？
'★ネットワークパスの考慮

'情報取得==========
'説明)フルファイルパスから、ベースファイル名を取得する
'C:\DirName\FileName.txt→FileName
Public Function BName(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        BName = .GetBaseName(vFPath)
    End With
End Function

'説明)フルファイルパスから、ファイル名を取得する
'C:\DirName\FileName.txt→FileName.txt
Public Function FName(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        FName = .GetFileName(vFPath)
    End With
End Function

'説明)フルファイルパスから、フォルダパス(デリミタなし)を取得する
'C:\DirName\FileName.txt→C:\DirName
Public Function DPath(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        DPath = .GetParentFolderName(vFPath)
    End With
End Function

'説明)フルファイルパスから、フォルダパス(デリミタ付き)を取得する
'C:\DirName\FileName.txt→C:\DirName\
Public Function DPath_(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        DPath_ = .GetParentFolderName(vFPath) & FILE_DELIMITER
    End With
End Function

'説明)フルファイルパスから、フォルダ名を取得する
'C:\DirName\FileName.txt→DirName
Public Function DName(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        Dim tmp As String
        tmp = .GetParentFolderName(vFPath)
        DName = EsRight(tmp, FILE_DELIMITER)
        '★渡されたパスが、パス形式でなかった場合の考慮
        '★C:\などの、トップフォルダだった場合の考慮
    End With
End Function

'説明)デリミタがなければデリミタをつけたパスを返す
'C:\DirName→C:\DirName\
Public Function D_(ByVal vDPath As String) As String
    '文字列がなければ、空文字を返す
    D_ = vbNullString
    If vbNullString = vDPath Then Exit Function
    
    'パスの最後の文字がデリミタでなければ、デリミタを付加して返す
    Dim tLastStr As String: tLastStr = Right(vDPath, 1)
    Dim tAddStr As String: tAddStr = ""
    If tLastStr <> FILE_DELIMITER Then tAddStr = FILE_DELIMITER
    D_ = vDPath & tAddStr
End Function

'説明)拡張子を取得する
'パスの後ろから、最初に「.」が見つかった位置の一つ後ろから終わりまでを取り出す
'見つからなかった場合は、空文字を返す
'C:\DirName\FileName.txt→txt
'FileName.txt→txt
Public Function getExtention(ByVal vFPath As String) As String
    getExtention = EsRight(vFPath, ".")
End Function

'判定==========
'説明)指定されたファイルパスの、ファイルが存在するか判定する
Public Function IsExistFile(ByVal vFPath As String) As Boolean
    With CreateObject("Scripting.Filesystemobject")
        IsExistFile = .FileExists(vFPath)
    End With
End Function

'説明)指定されたフォルダパスの、フォルダが存在するか判定する
Public Function IsExistDir(ByVal vDPath As String) As Boolean
    With CreateObject("Scripting.Filesystemobject")
        IsExistDir = .FolderExists(vDPath)
    End With
End Function

'説明)指定されたパスが、存在するか判定する
'ファイルとフォルダで、両方判定する
Public Function IsExist(ByVal vFPath As String) As Boolean
    Dim ret As Boolean: ret = False
    
    '指定パスが、ファイルとして判定する
    ret = IsExistFile(vFPath)
    
    'ファイルとしての判定がFALSEであれば、フォルダとして判定する
    If Not (ret) Then
        ret = IsExistDir(vFPath)
    End If
    
    '双方の結果を返す
    IsExist = ret
End Function

'説明)指定されたファイルパスの、親フォルダが存在するか判定する
Public Function IsExistParent(ByVal vFPath As String) As Boolean
    '親フォルダのパスを取得して、存在するか確認
    IsExistParent = IsExistDir(DPath(vFPath))
End Function

'ファイル名操作==========
'説明)指定ファイル名を、日付(YYYYMMDDhhmmss)つきファイル名に変換したものを返す
'FileName.txt→FileName_20161231125959.txt
Public Function FName_YYYYMMDDhhmmss(ByVal vFName As String) As String
    Dim tBName As String: tBName = BName(vFName)
    Dim tExtention As String: tExtention = getExtention(vFName)
    
    FName_YYYYMMDDhhmmss = tBName & "_" & YYYYMMDDhhmmss & "." & tExtention
    
End Function

'説明)指定ファイル名を、日付(乱数)(YYYYMMDDhhmmssRRR)つきファイル名に変換したものを返す
'FileName.txt→FileName_20161231125959.txt
Public Function FName_YYYYMMDDhhmmssRRR(ByVal vFName As String) As String
    Dim tBName As String: tBName = BName(vFName)
    Dim tExtention As String: tExtention = getExtention(vFName)
    
    FName_YYYYMMDDhhmmssRRR = tBName & "_" & YYYYMMDDhhmmssRRR & "." & tExtention
    
End Function

