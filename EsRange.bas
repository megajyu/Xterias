Attribute VB_Name = "EsRange"
Option Explicit

'Cell:Ce:�P�Ƃ̃Z��
'Box:Bx:�l�̋l�܂����P��(�s)�͈̔�
'Range:Rg:��Z�����܂މ\���̂���1��(�s)�͈̔�
'Area:��ƍs�̂����ꂩ�ł�2��(�s)�ł���͈�

'�w��͈͂�"�s"����w�蕶����T���āArange��Ԃ�
'�w��͈͂�"��"����w�蕶����T���āArange��Ԃ�
'���K�\���ŕ�������v����range��Ԃ�

'�w��Z�����܂ޗ�ɂ����āA(�ォ��)�󗓂łȂ��擪��(������)�󗓂ł͂Ȃ��Ō��Range��Ԃ�
'Retu

'�w��Z�����܂ލs�ɂ����āA(������)�󗓂łȂ��擪��(�E����)�󗓂ł͂Ȃ��Ō��Range��Ԃ�
'Gyou

'Start
'End

'VPinch�E�E�E�I�v�V�����Ŕ͈͂��w��H
'HPinch

'deck
'packet
'pack
'Box
'bag

'�w��Area�̒��ŁA�f�[�^�̂���ŏ㕔�ɐ؂�l�߂�
'----------------------------------
'����)�w��Cell�����Ɍ��āA�󗓂łȂ��Ō��Cell��Ԃ�
Public Function CeHead(ByRef vCell As Range) As Range
    Set CeHead = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    '�V�[�g�̒�ł���΁A�w��Cell��Ԃ�
    If vCell.Row = MIN_ROW_COUNT Then Exit Function
    
    '�w��Cell�̈�オ��Cell�Ȃ�΁A�w��Cell��Ԃ�
    If IsNoTextCell(vCell.Offset(-1, 0)) Then Exit Function
    
    '��L�ȊO�́A�w��Cell����Ctrl�{����Cell��Ԃ�
    Set CeHead = vCell.End(xlUp)

End Function

'����)�w��Cell���牺�Ɍ��āA�󗓂łȂ��Ō��Cell��Ԃ�
Public Function CeFoot(ByRef vCell As Range) As Range
    Set CeFoot = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    '�V�[�g�̒�ł���΁A�w��Cell��Ԃ�
    If vCell.Row = MAX_ROW_COUNT Then Exit Function
    
    '�w��Cell�̈������Cell�Ȃ�΁A�w��Cell��Ԃ�
    If IsNoTextCell(vCell.Offset(1, 0)) Then Exit Function
    
    '��L�ȊO�́A�w��Cell����Ctrl�{����Cell��Ԃ�
    Set CeFoot = vCell.End(xlDown)

End Function
'����)�w��Cell���獶�Ɍ��āA�󗓂łȂ��Ō��Cell��Ԃ�
Public Function CeLeftHand(ByRef vCell As Range) As Range
    Set CeLeftHand = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    '�V�[�g�̍��[�ł���΁A�w��Cell��Ԃ�
    If vCell.Row = MIN_COLUMN_COUNT Then Exit Function
    
    '�w��Cell�̈������Cell�Ȃ�΁A�w��Cell��Ԃ�
    If IsNoTextCell(vCell.Offset(0, -1)) Then Exit Function
    
    '��L�ȊO�́A�w��Cell����Ctrl�{����Cell��Ԃ�
    Set CeLeftHand = vCell.End(xlToLeft)

End Function

'����)�w��Cell����E�Ɍ��āA�󗓂łȂ��Ō��Cell��Ԃ�
Public Function CeRightHand(ByRef vCell As Range) As Range
    Set CeRightHand = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    '�V�[�g�̉E�[�ł���΁A�w��Cell��Ԃ�
    If vCell.Row = MAX_COLUMN_COUNT Then Exit Function
    
    '�w��Cell�̈�E����Cell�Ȃ�΁A�w��Cell��Ԃ�
    If IsNoTextCell(vCell.Offset(0, 1)) Then Exit Function
    
    '��L�ȊO�́A�w��Cell����Ctrl�{����Cell��Ԃ�
    Set CeRightHand = vCell.End(xlToRight)

End Function
'----------------------------------
'����)�w��Cell�����Ɍ��āA�󗓂łȂ��Ō��Cell�܂ł�Pack��Ԃ�
'���͒l)Cell:�N�_
'�߂�l)Box:�N�_����󗓂łȂ��Ō��Cell�܂ł͈̔�
Public Function PkUpperBody(ByRef vCell As Range) As Range
    Set PkUpperBody = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    Set PkUpperBody = vCell.Worksheet.Range(CeHead(vCell), vCell)
End Function
'����)�w��Cell���牺�Ɍ��āA�󗓂łȂ��Ō��Cell�܂ł�Pack��Ԃ�
'���͒l)Cell:�N�_
'�߂�l)Box:�N�_����󗓂łȂ��Ō��Cell�܂ł͈̔�
Public Function PkLowerBody(ByRef vCell As Range) As Range
    Set PkLowerBody = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    Set PkLowerBody = vCell.Worksheet.Range(vCell, CeFoot(vCell))
End Function
Public Function PkHung(ByRef vCell As Range) As Range
    Set PkHung = PkLowerBody(vCell)
End Function

'����)�w��Cell���獶�Ɍ��āA�󗓂łȂ��Ō��Cell�܂ł�Pack��Ԃ�
'���͒l)Cell:�N�_
'�߂�l)Box:�N�_����󗓂łȂ��Ō��Cell�܂ł͈̔�
Public Function PkLeftArm(ByRef vCell As Range) As Range
    Set PkLeftArm = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    Set PkLeftArm = vCell.Worksheet.Range(CeLeftHand(vCell), vCell)
End Function

'����)�w��Cell����E�Ɍ��āA�󗓂łȂ��Ō��Cell�܂ł�Pack��Ԃ�
'���͒l)Cell:�N�_
'�߂�l)Box:�N�_����󗓂łȂ��Ō��Cell�܂ł͈̔�
Public Function PkRightArm(ByRef vCell As Range) As Range
    Set PkRightArm = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    Set PkRightArm = vCell.Worksheet.Range(vCell, CeRightHand(vCell))
End Function
Public Function PkTail(ByRef vCell As Range) As Range
    Set PkTail = PkRightArm(vCell)
End Function

'----------------------------------
'����)�w��Cell����㉺�Ɍ��āA�󗓂łȂ�Pack��Ԃ�
Public Function PkBody(ByRef vCell As Range) As Range
    Set PkBody = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    '���E�Ɍ��āA�l�̘A������[��Cell���擾
    Dim tSCell As Range: Set tSCell = CeLeftHand(vCell)
    Dim tECell As Range: Set tECell = CeLeftRight(vCell)
    
    Set PkBody = vCell.Worksheet.Range(tSCell, tECell)
End Function

'����)�w��Cell���獶�E�Ɍ��āA�󗓂łȂ�Pack��Ԃ�
Public Function PkArms(ByRef vCell As Range) As Range
    Set PkArms = Nothing
    
    '�������s���Ȃ�΁ANothing��Ԃ�
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '�w��Cell����Cell�Ȃ�΁ANothing��Ԃ�
    If IsNoTextCell(vCell) Then Exit Function
    
    '���E�Ɍ��āA�l�̘A������[��Cell���擾
    Dim tSCell As Range: Set tSCell = CeLeftHand(vCell)
    Dim tECell As Range: Set tECell = CeLeftRight(vCell)
    
    Set PkArms = vCell.Worksheet.Range(tSCell, tECell)
End Function
Public Function PkWing(ByRef vCell As Range) As Range
    Set PkWing = PkArms(vCell)
End Function

'----------------------------------
'����)�w��Cell���܂ޗ�ŁA�P�s�ڂ��牺�Ɍ��ċ󔒂ł͂Ȃ��ŏ���Cell��Ԃ�
Public Function CeTop(ByRef vCell As Range) As Range
    Dim tRg1st As Range '1�s�߂�Range
    Dim tRgFind As Range '���Ɍ������Č�������Range
    
    With vCell.Worksheet
        '�w��Cell���܂ޗ�̂P�s�߂�Cell���擾
        Set tRg1st = .Cells(MIN_ROW_COUNT, vRg.Column)
    
        '�l������΁A�����Ԃ�
        If Not IsNoTextCell(tRg1st) Then
            Set CeTop = tRg1st
            Exit Function
        End If
        
        '���ɒl��T��
        Set tRgFind = tRg1st.End(xlDown)
        
        '�󗓂��A�ŉ��s�Ȃ�Η�ɒl�������Ɣ��f
        If IsNoTextCell(tRgFind) And tRgFind.Row = MAX_ROW_COUNT Then
            Set CeTop = Nothing
            Exit Function
        End If
    
        Set CeTop = tRgFind
    End With
End Function
'����)�w��Cell���܂ޗ�ŁA�ŉ��s�����Ɍ��ċ󔒂ł͂Ȃ��ŏ���Cell��Ԃ�
Public Function CeBottom(ByRef vCell As Range) As Range
    Dim tRgLast As Range '�ŏI�s��Range
    Dim tRgFind As Range '��Ɍ������Č�������Range
    
    With vCell.Worksheet
        '�w��Cell���܂ޗ�̍ŏI�s��Cell���擾
        Set tRgLast = .Cells(MAX_ROW_COUNT, vRg.Column)
    
        '�l������΁A�����Ԃ�
        If Not IsNoTextCell(tRgLast) Then
            Set CeBottom = tRgLast
            Exit Function
        End If
        
        '���ɒl��T��
        Set tRgFind = tRgLast.End(xlUp)
        
        '�󗓂��A�P�s�߂Ȃ�Η�ɒl�������Ɣ��f
        If IsNoTextCell(tRgFind) And tRgFind.Row = MIN_ROW_COUNT Then
            Set CeBottom = Nothing
            Exit Function
        End If
    
        Set CeBottom = tRgFind
    End With
End Function
'����)�w��Cell���܂ޗ�ŁA�P��ڂ���E�Ɍ��ċ󔒂ł͂Ȃ��ŏ���Cell��Ԃ�
Public Function CeLeftEdge(ByRef vCell As Range) As Range
    Dim tRg1st As Range '1��߂�Range
    Dim tRgFind As Range '�E�Ɍ������Č�������Range
    
    With vCell.Worksheet
        '�w��Cell���܂ދƂ̂P��ڂ�Cell���擾
        Set tRg1st = .Cells(vRg.Row, MIN_COLUMN_COUNT)
    
        '�l������΁A�����Ԃ�
        If Not IsNoTextCell(tRg1st) Then
            Set CeLeftEdge = tRg1st
            Exit Function
        End If
        
        '�E�ɒl��T��
        Set tRgFind = tRg1st.End(xlToRight)
        
        '�󗓂��A�ŏI��Ȃ�΍s�ɒl�������Ɣ��f
        If IsNoTextCell(tRgFind) And tRgFind.Column = MAX_COLUMN_COUNT Then
            Set CeLeftEdge = Nothing
            Exit Function
        End If
    
        Set CeLeftEdge = tRgFind
    End With
End Function
'����)�w��Cell���܂ލs�ŁA�ŏI�񂩂獶�Ɍ��ċ󔒂ł͂Ȃ��ŏ���Cell��Ԃ�
Public Function CeRightEdge(ByRef vCell As Range) As Range
    Dim tRgLast As Range '�Ő����Range
    Dim tRgFind As Range '���Ɍ������Č�������Range
    
    With vCell.Worksheet
        '�w��Cell���܂ލs�̍ŏI���Cell���擾
        Set tRgLast = .Cells(MAX_ROW_COUNT, vRg.Column)
    
        '�l������΁A�����Ԃ�
        If Not IsNoTextCell(tRgLast) Then
            Set CeRightEdge = tRgLast
            Exit Function
        End If
        
        '���ɒl��T��
        Set tRgFind = tRgLast.End(xlToLeft)
        
        '�󗓂��A�P�s�߂Ȃ�Η�ɒl�������Ɣ��f
        If IsNoTextCell(tRgFind) And tRgFind.Column = MIN_COLUMN_COUNT Then
            Set CeRightEdge = Nothing
            Exit Function
        End If
    
        Set CeRightEdge = tRgFind
    End With
End Function
'----------------------------------

'����)�w��Cell���܂ޗ�ŁA
'�P�s�߂��牺�Ɍ��ċ󗓂łȂ��ŏ���Cell����
'�ŉ��s�����Ɍ��ċ󗓂łȂ��ŏ���Cell�܂ł�Box��߂�
'���Q��ȏ�̏ꍇ�̑Ή����K�v(�p�̗񂾂���ΏۂƂ���H)
Public Function BxPinchColumn(ByRef vCell As Range) As Range
    Set BxPinchColumn = Nothing
    
    Dim SRg As Range: Set SRg = CeTop(vCell)
    Dim ERg As Range: Set ERg = CeBottom(vCell)
    
    If SRg Is Nothing Then Exit Function
    If ERg Is Nothing Then Exit Function
    
    Set BxPinchColumn = vCell.Worksheet.Range(SRg.Address, ERg.Address)

End Function

'����)�w��Cell���܂ލs�ŁA
'�P��߂���E�Ɍ��ċ󗓂łȂ��ŏ���Cell����
'�ŏI�񂩂獶�Ɍ��ċ󗓂łȂ��ŏ���Cell�܂ł�Box��߂�
'���Q�s�ȏ�̏ꍇ�̑Ή����K�v(�p�̍s������ΏۂƂ���H)
Public Function BxPinchRow(ByRef vCell As Range) As Range
    Set BxPinchRow = Nothing

    Dim SRg As Range: Set SRg = CeLeftEdge(vCell)
    Dim ERg As Range: Set ERg = CeRightEdge(vCell)
    
    If SRg Is Nothing Then Exit Function
    If ERg Is Nothing Then Exit Function
    
    Set BxPinchRow = vCell.Worksheet.Range(SRg.Address, ERg.Address)

End Function

'----------------------------------
'����)�w��Area�̐擪Cell��߂�
Public Function CeFirst(ByRef vArea As Range) As Range
    Set CeFirst = vArea(1)
End Function
'����)�w��Area�̍ŏICell��߂�
Public Function CeLast(ByRef vArea As Range) As Range
    Set CeLast = vArea(vArea.Count)
End Function

'----------------------------------
'����)�w��Area���A�l�̂���s�܂ŏ�[��؂�l�߂�Area��߂�
Public Function ArCeil(ByRef vArea As Range) As Range
    Set ArCeil = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    Dim SRg As Range: Set SRg = CeFirst(vArea)
    Dim ERg As Range: Set ERg = CeLast(vArea)
    
    '�ォ��P�s���E���Ă����B�����ꂩ��Cell�ɒl������΁A������
    Dim tRg As Range
    Dim cnt As Long: cnt = 0
    For Each tRg In vArea.Rows
        If Not IsNoVisibleText(JoinRgText(tRg)) Then Exit For
        cnt = cnt + 1
    Next
    
    '�ŉ��s�܂Ō�����Ȃ�������ANothing
    If cnt = vArea.Rows.Count Then Exit Function
    
    '�w��Area����A�l�����������s�܂ŃI�t�Z�b�g�`�ŉ��s�܂ł�߂�
    Set ArCeil = vArea.Worksheet.Range(SRg.Offset(cnt, 0), ERg)
End Function

'����)�w��Area���A�l�̂���s�܂ŉ��[��؂�l�߂�Area��߂�
Public Function ArFloor(ByRef vArea As Range) As Range
    Set ArFloor = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    Dim SRg As Range: Set SRg = CeFirst(vArea)
    
    '������P�s���E���Ă����B�����ꂩ��Cell�ɒl������΁A������
    Dim tRg As Range
    Dim cnt As Long: cnt = 0
    For cnt = vArea.Rows.Count To 0 Step -1
        If cnt = 0 Then Exit For 'Rows��1�n�܂�B0�͌�����Ȃ���������
        Set tRg = vArea.Rows(cnt)
        If Not IsNoVisibleText(JoinRgText(tRg)) Then Exit For
    Next
    
    '�ŏ�s�܂Ō�����Ȃ�������ANothing
    If cnt = 0 Then Exit Function
    
    '�w��Area����A�ŏ�s�`�l�����������s�܂ł�߂�
    Set ArFloor = vArea.Worksheet.Range(SRg, CeLast(SRg.Offset(cnt, 0)))
End Function

'����)�w��Area���A�l�̂����܂ŉE�[��؂�l�߂�Area��߂�
Public Function ArRightWall(ByRef vArea As Range) As Range
    Set ArRightWall = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    Dim SRg As Range: Set SRg = CeFirst(vArea)
    
    '������P�s���E���Ă����B�����ꂩ��Cell�ɒl������΁A������
    Dim tRg As Range
    Dim cnt As Long: cnt = 0
    For cnt = vArea.Rows.Columns To 0 Step -1
        If cnt = 0 Then Exit For 'Rows��1�n�܂�B0�͌�����Ȃ���������
        Set tRg = vArea.Columns(cnt)
        If Not IsNoVisibleText(JoinRgText(tRg)) Then Exit For
    Next
    
    '�ŏ�s�܂Ō�����Ȃ�������ANothing
    If cnt = 0 Then Exit Function
    
    '�w��Area����A�ŏ�s�`�l�����������s�܂ł�߂�
    Set ArRightWall = vArea.Worksheet.Range(SRg, CeLast(SRg.Offset(0, cnt)))

End Function
'����)�w��Area���A�l�̂����܂ō��[��؂�l�߂�Area��߂�
Public Function ArLeftWall(ByRef vArea As Range) As Range
    Set ArLeftWall = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    Dim SRg As Range: Set SRg = CeFirst(vArea)
    Dim ERg As Range: Set ERg = CeLast(vArea)
    
    '������P�񂸂E���Ă����B�����ꂩ��Cell�ɒl������΁A������
    Dim tRg As Range
    Dim cnt As Long: cnt = 0
    For Each tRg In vArea.Columns
        If Not IsNoVisibleText(JoinRgText(tRg)) Then Exit For
        cnt = cnt + 1
    Next
    
    '�ŉ��s�܂Ō�����Ȃ�������ANothing
    If cnt = vArea.Columns.Count Then Exit Function
    
    '�w��Area����A�l�����������s�܂ŃI�t�Z�b�g�`�ŉ��s�܂ł�߂�
    Set ArLeftWall = vArea.Worksheet.Range(SRg.Offset(0, cnt), ERg)
    
End Function

'----------------------------------
'----------------------------------
'����)�Q�̎w��Area�̏d�Ȃ�Area��߂�
Public Function ArIntersect(ByRef vAreaA As Range, ByRef vAreaB As Range) As Range
    Set ArIntersect = Nothing
    
    If vAreaA Is Nothing Then Exit Function
    If vAreaB Is Nothing Then Exit Function
    
    On Error Resume Next    '�d�Ȃ�ӏ��������ƁAIntersect���G���[���o���̂ŁB
    Set RgIntersect = Intersect(vAreaA, vAreaB)
    
End Function

'����)�w��Area�̎w��Ԗڂ̍s��Box��߂�
Public Function BxRow(ByRef vArea As Range, ByRef vIdx As Long) As Range
    Set BxRow = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    '�w��Ԗڂ��A�w��Area�̍s�����������ꍇ��Nothing��߂�
    If vIdx > vArea.Rows.Count Then Exit Function
        
    '�w��Ԗڂ̍s��Area�̏d�Ȃ�͈͂��擾����
    Set BxRow = ArIntersect(tRgRow.Offset(vIdx, 0).EntireRow, vArea)

End Function

'����)�w��Area�̎w��Ԗڂ̗��Box��߂�
Public Function BxColumn(ByRef vArea As Range, ByRef vIdx As Long) As Range
    Set BxColumn = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    '�w��Ԗڂ��A�w��Area�̗񐔂��������ꍇ��Nothing��߂�
    If vIdx > vArea.Columns.Count Then Exit Function
        
    '�w��Ԗڂ̗��Area�̏d�Ȃ�͈͂��擾����
    Set BxColumn = ArIntersect(tRgRow.Offset(0, vIdx).EntireColumn, vArea)
End Function

'----------------------------------
'����)�w��Area�̎g�p�͈�(CurrentRegion)��߂�
Public Function ArActive(ByRef vArea As Range) As Range
    Set ArActive = CeFirst(vArea).CurrentRegion
End Function
'����)�w��Area�̑�����g�p�͈�(CurrentRegion)�̐擪Cell��߂�
Public Function CeActiveFirst(ByRef vArea As Range) As Range
    Set CeActiveFirst = CeFirst(ArActive(vArea))
End Function
'����)�w��Area�̑�����g�p�͈�(CurrentRegion)�̍ŏICell��߂�
Public Function CeActiveLast(ByRef vArea As Range) As Range
    Set CeActiveLast = CeLast(ArActive(vArea))
End Function

'----------------------------------
'����)�w��Cell����E�Ɍ��Ēl������cell�����݂��邩�����߂�
Public Function IsExistTextAtRight(ByRef vCell As Range) As Boolean
    IsExistTextAtRight = Not IsNoTextCell(vCell.End(xlToRight))
End Function
'����)�w��Cell���獶�Ɍ��Ēl������cell�����݂��邩�����߂�
Public Function IsExistTextAtLeft(ByRef vCell As Range) As Boolean
    IsExistTextAtLeft = Not IsNoTextCell(vCell.End(xlToLeft))
End Function
'����)�w��Cell�����Ɍ��Ēl������cell�����݂��邩�����߂�
Public Function IsExistTextAtUp(ByRef vCell As Range) As Boolean
    IsExistTextAtUp = Not IsNoTextCell(vCell.End(xlUp))
End Function
'����)�w��Cell���牺�Ɍ��Ēl������cell�����݂��邩�����߂�
Public Function IsExistTextAtDown(ByRef vCell As Range) As Boolean
    IsExistTextAtDown = Not IsNoTextCell(vCell.End(xlDown))
End Function


'----------------------------------
'����)���Z��(Area)�����ɍi�荞��
Public Function ArVisible(ByRef vArea As Range) As Range
    Set ArVisible = vArea.SpecialCells(xlCellTypeVisible)
End Function

'����)��P�����͈̔͂��㉺�ɁA��Q�����͈̔͂����E�ɍL���āA��������͈�(Area)��߂�
Public Function ArCross(ByRef vAreaA As Range, ByRef vAreaB As Range) As Range
    Set ArCross = ArIntersect(vAreaA.EntireColumn, vAreaB.EntireRow)
End Function

'����)�w��Cell������̋N�_�Ƃ��āA�E�Ɏw�蒷�A���Ɏw�蒷�ɐL�΂���Area��߂�
Public Function ArSpread(ByRef vCell As Range, ByVal vRightLength As Long, ByVal vDownLength As Long) As Range
    Set ArSpread = vCell.Worksheet.renge(vCell, vCell.Offset(vDownLength, vRightLength))
End Function

'����)�w��Area�̓��e���N���A����
Public Function ClearRgContents(ByRef vArea As Range) As Range
    Set ClearRgContents = vArea '�߂�l�́A���̂܂܂̃G���A��߂�
    If vArea Is Nothing Then Exit Function
    vArea.ClearComments
End Function

'----------------------------------
'����)�w��Area���AJoin�����������߂�
'Cell�̊Ԃ́A�w�肳�ꂽ�f���~�^(default��Tab)�łȂ���B
'�����s�̏ꍇ�A���s��؂�
Public Function JoinRgText(ByRef vArea As Range, Optional ByVal vDelimiter As String = vbTab, Optional vAddEndReturn As Boolean = True) As String
    Dim ret() As String: Dim retRow() As String
    Dim tRg As Range: Dim tRgRow As Range
    Dim cntRow As Long: Dim cnt As Long
    
    cntRow = 0
    For Each tRgRow In vArea.Rows
    
        '�s�Ő؂�o���A�eCell�̒l���Ȃ���
        cnt = 0
        For Each tRg In tRgRow.Cells
            ReDim Preserve ret(cnt)
            ret(cnt) = tRg.Text
            cnt = cnt + 1
        Next
        
        '1�s���Ƃ̌��ʂ��i�[
        ReDim Preserve retRow(cntRow)
        retRow(cntRow) = Join(ret, vDelimiter)
        cntRow = cntRow + 1
    Next
    
    JoinRgText = Join(retRow, vbCrLf)
    
    If vAddEndReturn Then
        JoinRgText = JoinRgText & vbCrLf
    End If
End Function

'----------------------------------
'����)�w��Area����A�E�����Ɏw�蕶�����������A��������Cell��߂�
Public Function CeFindRight(ByRef vArea As Range, ByVal vString As String) As Range
    Set CeFindRight = vArea.Find(What:=vString, LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByRows, MatchByte:=False)
End Function

'����)�w��Area����A�������Ɏw�蕶�����������A��������Cell��߂�
Public Function CsFindDown(ByRef vArea As Range, ByVal vString As String) As Range
    Set CeFindDown = vArea.Find(What:=vString, LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, MatchByte:=False)
End Sub

'����)�w��Area����A�w�蕶�����܂�Cell�̏W����߂�
Public Function CsFind(ByRef vArea As Range, ByVal vString As String) As Range
    Dim ret As Range
    Dim tRg As Range
    
    For Each tRg In vArea
        If InStr(tRg.Value, vString) > 0 Then
            If ret Is Nothing Then
                Set ret = tRg
            Else
                Set ret = Union(ret, tRg)
            End If
        End If
    Next
    Set CsFind = ret
End Function

'����)�w��Area����A���K�\���ɊY������Cell�̏W����߂�
Public Function CsFindReg(ByRef vArea As Range, ByVal vReg As String) As Range
    Dim ret As Range
    Dim tRg As Range
    For Each tRg In vArea
        If RegChk(tRg.Value, vReg) Then
            If ret Is Nothing Then
                Set ret = tRg
            Else
                Set ret = Union(ret, tRg)
            End If
        End If
    Next
    
    Set CsFindReg = ret
End Function


