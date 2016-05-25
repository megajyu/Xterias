Attribute VB_Name = "EsFileControl"
Option Explicit
'���f���~�^�̈������ǂ����邩�H
'���l�b�g���[�N�p�X�̍l��

'���擾==========
'����)�t���t�@�C���p�X����A�x�[�X�t�@�C�������擾����
'C:\DirName\FileName.txt��FileName
Public Function BName(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        BName = .GetBaseName(vFPath)
    End With
End Function

'����)�t���t�@�C���p�X����A�t�@�C�������擾����
'C:\DirName\FileName.txt��FileName.txt
Public Function FName(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        FName = .GetFileName(vFPath)
    End With
End Function

'����)�t���t�@�C���p�X����A�t�H���_�p�X(�f���~�^�Ȃ�)���擾����
'C:\DirName\FileName.txt��C:\DirName
Public Function DPath(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        DPath = .GetParentFolderName(vFPath)
    End With
End Function

'����)�t���t�@�C���p�X����A�t�H���_�p�X(�f���~�^�t��)���擾����
'C:\DirName\FileName.txt��C:\DirName\
Public Function DPath_(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        DPath_ = .GetParentFolderName(vFPath) & FILE_DELIMITER
    End With
End Function

'����)�t���t�@�C���p�X����A�t�H���_�����擾����
'C:\DirName\FileName.txt��DirName
Public Function DName(ByVal vFPath As String) As String
    With CreateObject("Scripting.Filesystemobject")
        Dim tmp As String
        tmp = .GetParentFolderName(vFPath)
        DName = EsRight(tmp, FILE_DELIMITER)
        '���n���ꂽ�p�X���A�p�X�`���łȂ������ꍇ�̍l��
        '��C:\�Ȃǂ́A�g�b�v�t�H���_�������ꍇ�̍l��
    End With
End Function

'����)�f���~�^���Ȃ���΃f���~�^�������p�X��Ԃ�
'C:\DirName��C:\DirName\
Public Function D_(ByVal vDPath As String) As String
    '�����񂪂Ȃ���΁A�󕶎���Ԃ�
    D_ = vbNullString
    If vbNullString = vDPath Then Exit Function
    
    '�p�X�̍Ō�̕������f���~�^�łȂ���΁A�f���~�^��t�����ĕԂ�
    Dim tLastStr As String: tLastStr = Right(vDPath, 1)
    Dim tAddStr As String: tAddStr = ""
    If tLastStr <> FILE_DELIMITER Then tAddStr = FILE_DELIMITER
    D_ = vDPath & tAddStr
End Function

'����)�g���q���擾����
'�p�X�̌�납��A�ŏ��Ɂu.�v�����������ʒu�̈��납��I���܂ł����o��
'������Ȃ������ꍇ�́A�󕶎���Ԃ�
'C:\DirName\FileName.txt��txt
'FileName.txt��txt
Public Function getExtention(ByVal vFPath As String) As String
    getExtention = EsRight(vFPath, ".")
End Function

'����==========
'����)�w�肳�ꂽ�t�@�C���p�X�́A�t�@�C�������݂��邩���肷��
Public Function IsExistFile(ByVal vFPath As String) As Boolean
    With CreateObject("Scripting.Filesystemobject")
        IsExistFile = .FileExists(vFPath)
    End With
End Function

'����)�w�肳�ꂽ�t�H���_�p�X�́A�t�H���_�����݂��邩���肷��
Public Function IsExistDir(ByVal vDPath As String) As Boolean
    With CreateObject("Scripting.Filesystemobject")
        IsExistDir = .FolderExists(vDPath)
    End With
End Function

'����)�w�肳�ꂽ�p�X���A���݂��邩���肷��
'�t�@�C���ƃt�H���_�ŁA�������肷��
Public Function IsExist(ByVal vFPath As String) As Boolean
    Dim ret As Boolean: ret = False
    
    '�w��p�X���A�t�@�C���Ƃ��Ĕ��肷��
    ret = IsExistFile(vFPath)
    
    '�t�@�C���Ƃ��Ă̔��肪FALSE�ł���΁A�t�H���_�Ƃ��Ĕ��肷��
    If Not (ret) Then
        ret = IsExistDir(vFPath)
    End If
    
    '�o���̌��ʂ�Ԃ�
    IsExist = ret
End Function

'����)�w�肳�ꂽ�t�@�C���p�X�́A�e�t�H���_�����݂��邩���肷��
Public Function IsExistParent(ByVal vFPath As String) As Boolean
    '�e�t�H���_�̃p�X���擾���āA���݂��邩�m�F
    IsExistParent = IsExistDir(DPath(vFPath))
End Function

'�t�@�C��������==========
'����)�w��t�@�C�������A���t(YYYYMMDDhhmmss)���t�@�C�����ɕϊ��������̂�Ԃ�
'FileName.txt��FileName_20161231125959.txt
Public Function FName_YYYYMMDDhhmmss(ByVal vFName As String) As String
    Dim tBName As String: tBName = BName(vFName)
    Dim tExtention As String: tExtention = getExtention(vFName)
    
    FName_YYYYMMDDhhmmss = tBName & "_" & YYYYMMDDhhmmss & "." & tExtention
    
End Function

'����)�w��t�@�C�������A���t(����)(YYYYMMDDhhmmssRRR)���t�@�C�����ɕϊ��������̂�Ԃ�
'FileName.txt��FileName_20161231125959.txt
Public Function FName_YYYYMMDDhhmmssRRR(ByVal vFName As String) As String
    Dim tBName As String: tBName = BName(vFName)
    Dim tExtention As String: tExtention = getExtention(vFName)
    
    FName_YYYYMMDDhhmmssRRR = tBName & "_" & YYYYMMDDhhmmssRRR & "." & tExtention
    
End Function

