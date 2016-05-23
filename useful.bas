Attribute VB_Name = "useful"
Option Explicit

'����==========
'����)�w�肳�ꂽCell�ŁA�e�L�X�g�����݂��邩���肷��
Public Function IsNoTextCell(ByRef vCell As Range) As Boolean
    IsNoTextCell = True
    If vCell Is Nothing Then Exit Function
    If vCell.Text <> "" Then IsNoTextCell = False
End Function
'����)(�^�u�Ȃǂ�������)�e�L�X�g�����݂��邩���肷��
Public Function IsNoVisibleText(ByRef vStr As String) As Boolean
    IsNoVisibleText = True
    
    Dim chkStr As String: chkStr = vStr
    chkStr = Replace(chkStr, vbTab, "") '�^�u������
    chkStr = Replace(chkStr, " ", "") '���p�X�y�[�X������
    chkStr = Replace(chkStr, "�@", "") '�S�p�X�y�[�X������
    chkStr = Replace(chkStr, vbCrLf, "") '���s������
    chkStr = Replace(chkStr, vbCr, "") '�L�����b�W���^�[��������
    chkStr = Replace(chkStr, vbLf, "") 'LF������
    
    If chkStr <> "" Then IsNoVisibleText = False
    
End Function


'���K�\��==========
'����)���K�\���Ń}�b�`���Ă��邩���肷��
Public Function RegChk(ByVal vStr As String, ByVal vPattern) As Boolean
    With CreateObject("VBScript.RegExp")
        .Pattern = vPattern
        RegChk = .test(vStr)
    End With
End Function

'����)���K�\���Ń}�b�`�����ӏ���u��������
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

'����)�}�b�`���������̔z����擾����
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

'����)�T�u�}�b�`���������̔z����擾����
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
    
    '�Y���������ꍇ�́A�󕶎���Ԃ�
    If cnt = 0 Then
        ReDim ret(1)
        ret(0) = vbNullString
    End If
    
    RegSubMatch = ret
    
End Function

'�����񑀍�==========
'����)��납�璲�ׂčŏ��ɑ�Q���������������ʒu�̈��납��I���܂ł�؂�o��
'������Ȃ������ꍇ�́A�󕶎���Ԃ�
Public Function EsRight(ByVal vStr As String, ByVal vSearchStr As String) As String
    Dim Idx As Long
    Idx = InStrRev(vStr, vSearchStr)
    
    '���������ʒu�̈��납��A�����񒷁[���������ʒu
    EsRight = Mid(vStr, Idx + 1, Len(vStr) - Idx)
    
    '������Ȃ������ꍇ�́A�󕶎���Ԃ�
    If Idx = 0 Then
        EsRight = vbNullString
    End If
End Function

'����)�O���璲�ׂčŏ��ɑ�Q���������������ʒu�̈�O�܂ł�؂�o��
'������Ȃ������ꍇ�́A�󕶎���Ԃ�
Public Function EsLeft(ByVal vStr As String, ByVal vSearchStr As String) As String
    Dim Idx As Long
    Idx = InStr(vStr, vSearchStr)
    
    '���������ʒu�̈��납��A�����񒷁[���������ʒu
    EsLeft = Left(vStr, Idx - 1)
    
    '������Ȃ������ꍇ�́A�󕶎���Ԃ�
    If Idx = 0 Then
        EsLeft = vbNullString
    End If
    
End Function

'����)�G���[�ɂȂ�ɂ���Split
Public Function EsSplit(ByVal vStr As String, ByVal vDelimiter As String, Optional ByVal vIdx As Long, Optional vDefault As String) As Variant
    Dim ret As Variant  '�Ԃ��l���A�u������v�u�z��v�̂Q��ނ��邽��
    
    '������̎w��Ȃ�
    '��vIdx���w�肳��Ă���:vbNullString:�������~�����̂�
    '��vIdx���w�肳��Ă��Ȃ�:��z��:for each�Ŏg�������̂�
    If vStr = vbNullString Then
        If Not vIdx = vbNullString Then
            EsSplit = vbNullString
        Else
            EsSplit = Array("")
        End If
        Exit Function
    End If
    
    '������𕪊�(�z��)
    ret = Split(vStr, vDilimiter)
    
    'vIdx���w�肳��Ă���ꍇ�́A�C���f�b�N�X���w���l��Ԃ�
    Dim Idx As Long
    If Not vIdx = vbNullString Then
        Idx = CLng(vIdx)
        If Idx <= UBoung(ret) Then
            ret = ret(Idx)
        Else
            '�w��Ȃ�:vbNullString�A�w�肠��:�w�肳�ꂽ�l
            ret = vDefault
        End If
    End If
    
    EsSplit = ret
End Function

'���t==========
'����)���݂̓��t(YYYYMMDDhhmmss)��Ԃ�
'20161231125959
Public Function YYYYMMDDhhmmss() As String
    YYYYMMDDhhmmss = Year(Now) & _
                                        Right("00" & Month(Now), 2) & _
                                        Right("00" & Day(Now), 2) & _
                                        Right("00" & Hour(Now), 2) & _
                                        Right("00" & Minute(Now), 2) & _
                                        Right("00" & Second(Now), 2)
End Function

'����)���݂̓��t(�����t��)(YYYYMMDDhhmmssRRR)��Ԃ�
'20161231125959034
Public Function YYYYMMDDhhmmssRRR() As String
    YYYYMMDDhhmmssRRR = YYYYMMDDhhmmss & Right("000" & CInt(Rnd() * 100), 3)
End Function

'�Z���I����Ԃ̔���==========
'����)�n���ꂽRange���A�ǂ̂悤�ȏ�ԂŗL�邩���肷��
Public Function GetRangeType(ByRef vRange As Range) As String
 '�r��
End Function

