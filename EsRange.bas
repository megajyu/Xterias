Attribute VB_Name = "EsRange"
Option Explicit

'Cell:Ce:単独のセル
'Box:Bx:値の詰まった１列(行)の範囲
'Range:Rg:空セルを含む可能性のある1列(行)の範囲
'Area:列と行のいずれかでも2列(行)である範囲

'指定範囲で"行"から指定文言を探して、rangeを返す
'指定範囲で"列"から指定文言を探して、rangeを返す
'正規表現で文字が一致するrangeを返す

'指定セルを含む列において、(上から)空欄でない先頭と(下から)空欄ではない最後のRangeを返す
'Retu

'指定セルを含む行において、(左から)空欄でない先頭と(右から)空欄ではない最後のRangeを返す
'Gyou

'Start
'End

'VPinch・・・オプションで範囲を指定？
'HPinch

'deck
'packet
'pack
'Box
'bag

'指定Areaの中で、データのある最上部に切り詰める
'----------------------------------
'説明)指定Cellから上に見て、空欄でない最後のCellを返す
Public Function CeHead(ByRef vCell As Range) As Range
    Set CeHead = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    'シートの底であれば、指定Cellを返す
    If vCell.Row = MIN_ROW_COUNT Then Exit Function
    
    '指定Cellの一つ上が空Cellならば、指定Cellを返す
    If IsNoTextCell(vCell.Offset(-1, 0)) Then Exit Function
    
    '上記以外は、指定CellからCtrl＋↓のCellを返す
    Set CeHead = vCell.End(xlUp)

End Function

'説明)指定Cellから下に見て、空欄でない最後のCellを返す
Public Function CeFoot(ByRef vCell As Range) As Range
    Set CeFoot = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    'シートの底であれば、指定Cellを返す
    If vCell.Row = MAX_ROW_COUNT Then Exit Function
    
    '指定Cellの一つ下が空Cellならば、指定Cellを返す
    If IsNoTextCell(vCell.Offset(1, 0)) Then Exit Function
    
    '上記以外は、指定CellからCtrl＋↓のCellを返す
    Set CeFoot = vCell.End(xlDown)

End Function
'説明)指定Cellから左に見て、空欄でない最後のCellを返す
Public Function CeLeftHand(ByRef vCell As Range) As Range
    Set CeLeftHand = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    'シートの左端であれば、指定Cellを返す
    If vCell.Row = MIN_COLUMN_COUNT Then Exit Function
    
    '指定Cellの一つ左が空Cellならば、指定Cellを返す
    If IsNoTextCell(vCell.Offset(0, -1)) Then Exit Function
    
    '上記以外は、指定CellからCtrl＋↓のCellを返す
    Set CeLeftHand = vCell.End(xlToLeft)

End Function

'説明)指定Cellから右に見て、空欄でない最後のCellを返す
Public Function CeRightHand(ByRef vCell As Range) As Range
    Set CeRightHand = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    'シートの右端であれば、指定Cellを返す
    If vCell.Row = MAX_COLUMN_COUNT Then Exit Function
    
    '指定Cellの一つ右が空Cellならば、指定Cellを返す
    If IsNoTextCell(vCell.Offset(0, 1)) Then Exit Function
    
    '上記以外は、指定CellからCtrl＋↓のCellを返す
    Set CeRightHand = vCell.End(xlToRight)

End Function
'----------------------------------
'説明)指定Cellから上に見て、空欄でない最後のCellまでのPackを返す
'入力値)Cell:起点
'戻り値)Box:起点から空欄でない最後のCellまでの範囲
Public Function PkUpperBody(ByRef vCell As Range) As Range
    Set PkUpperBody = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    Set PkUpperBody = vCell.Worksheet.Range(CeHead(vCell), vCell)
End Function
'説明)指定Cellから下に見て、空欄でない最後のCellまでのPackを返す
'入力値)Cell:起点
'戻り値)Box:起点から空欄でない最後のCellまでの範囲
Public Function PkLowerBody(ByRef vCell As Range) As Range
    Set PkLowerBody = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    Set PkLowerBody = vCell.Worksheet.Range(vCell, CeFoot(vCell))
End Function
Public Function PkHung(ByRef vCell As Range) As Range
    Set PkHung = PkLowerBody(vCell)
End Function

'説明)指定Cellから左に見て、空欄でない最後のCellまでのPackを返す
'入力値)Cell:起点
'戻り値)Box:起点から空欄でない最後のCellまでの範囲
Public Function PkLeftArm(ByRef vCell As Range) As Range
    Set PkLeftArm = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    Set PkLeftArm = vCell.Worksheet.Range(CeLeftHand(vCell), vCell)
End Function

'説明)指定Cellから右に見て、空欄でない最後のCellまでのPackを返す
'入力値)Cell:起点
'戻り値)Box:起点から空欄でない最後のCellまでの範囲
Public Function PkRightArm(ByRef vCell As Range) As Range
    Set PkRightArm = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    Set PkRightArm = vCell.Worksheet.Range(vCell, CeRightHand(vCell))
End Function
Public Function PkTail(ByRef vCell As Range) As Range
    Set PkTail = PkRightArm(vCell)
End Function

'----------------------------------
'説明)指定Cellから上下に見て、空欄でないPackを返す
Public Function PkBody(ByRef vCell As Range) As Range
    Set PkBody = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    '左右に見て、値の連続する端のCellを取得
    Dim tSCell As Range: Set tSCell = CeLeftHand(vCell)
    Dim tECell As Range: Set tECell = CeLeftRight(vCell)
    
    Set PkBody = vCell.Worksheet.Range(tSCell, tECell)
End Function

'説明)指定Cellから左右に見て、空欄でないPackを返す
Public Function PkArms(ByRef vCell As Range) As Range
    Set PkArms = Nothing
    
    '引数が不正ならば、Nothingを返す
    If vCell Is Nothing Then Exit Function
    If IsEmpty(vCell) Then Exit Function
    
    '指定Cellが空Cellならば、Nothingを返す
    If IsNoTextCell(vCell) Then Exit Function
    
    '左右に見て、値の連続する端のCellを取得
    Dim tSCell As Range: Set tSCell = CeLeftHand(vCell)
    Dim tECell As Range: Set tECell = CeLeftRight(vCell)
    
    Set PkArms = vCell.Worksheet.Range(tSCell, tECell)
End Function
Public Function PkWing(ByRef vCell As Range) As Range
    Set PkWing = PkArms(vCell)
End Function

'----------------------------------
'説明)指定Cellを含む列で、１行目から下に見て空白ではない最初のCellを返す
Public Function CeTop(ByRef vCell As Range) As Range
    Dim tRg1st As Range '1行めのRange
    Dim tRgFind As Range '下に検索して見つかったRange
    
    With vCell.Worksheet
        '指定Cellを含む列の１行めのCellを取得
        Set tRg1st = .Cells(MIN_ROW_COUNT, vRg.Column)
    
        '値があれば、それを返す
        If Not IsNoTextCell(tRg1st) Then
            Set CeTop = tRg1st
            Exit Function
        End If
        
        '下に値を探す
        Set tRgFind = tRg1st.End(xlDown)
        
        '空欄かつ、最下行ならば列に値が無いと判断
        If IsNoTextCell(tRgFind) And tRgFind.Row = MAX_ROW_COUNT Then
            Set CeTop = Nothing
            Exit Function
        End If
    
        Set CeTop = tRgFind
    End With
End Function
'説明)指定Cellを含む列で、最下行から上に見て空白ではない最初のCellを返す
Public Function CeBottom(ByRef vCell As Range) As Range
    Dim tRgLast As Range '最終行のRange
    Dim tRgFind As Range '上に検索して見つかったRange
    
    With vCell.Worksheet
        '指定Cellを含む列の最終行のCellを取得
        Set tRgLast = .Cells(MAX_ROW_COUNT, vRg.Column)
    
        '値があれば、それを返す
        If Not IsNoTextCell(tRgLast) Then
            Set CeBottom = tRgLast
            Exit Function
        End If
        
        '下に値を探す
        Set tRgFind = tRgLast.End(xlUp)
        
        '空欄かつ、１行めならば列に値が無いと判断
        If IsNoTextCell(tRgFind) And tRgFind.Row = MIN_ROW_COUNT Then
            Set CeBottom = Nothing
            Exit Function
        End If
    
        Set CeBottom = tRgFind
    End With
End Function
'説明)指定Cellを含む列で、１列目から右に見て空白ではない最初のCellを返す
Public Function CeLeftEdge(ByRef vCell As Range) As Range
    Dim tRg1st As Range '1列めのRange
    Dim tRgFind As Range '右に検索して見つかったRange
    
    With vCell.Worksheet
        '指定Cellを含む業の１列目のCellを取得
        Set tRg1st = .Cells(vRg.Row, MIN_COLUMN_COUNT)
    
        '値があれば、それを返す
        If Not IsNoTextCell(tRg1st) Then
            Set CeLeftEdge = tRg1st
            Exit Function
        End If
        
        '右に値を探す
        Set tRgFind = tRg1st.End(xlToRight)
        
        '空欄かつ、最終列ならば行に値が無いと判断
        If IsNoTextCell(tRgFind) And tRgFind.Column = MAX_COLUMN_COUNT Then
            Set CeLeftEdge = Nothing
            Exit Function
        End If
    
        Set CeLeftEdge = tRgFind
    End With
End Function
'説明)指定Cellを含む行で、最終列から左に見て空白ではない最初のCellを返す
Public Function CeRightEdge(ByRef vCell As Range) As Range
    Dim tRgLast As Range '最数列のRange
    Dim tRgFind As Range '左に検索して見つかったRange
    
    With vCell.Worksheet
        '指定Cellを含む行の最終列のCellを取得
        Set tRgLast = .Cells(MAX_ROW_COUNT, vRg.Column)
    
        '値があれば、それを返す
        If Not IsNoTextCell(tRgLast) Then
            Set CeRightEdge = tRgLast
            Exit Function
        End If
        
        '左に値を探す
        Set tRgFind = tRgLast.End(xlToLeft)
        
        '空欄かつ、１行めならば列に値が無いと判断
        If IsNoTextCell(tRgFind) And tRgFind.Column = MIN_COLUMN_COUNT Then
            Set CeRightEdge = Nothing
            Exit Function
        End If
    
        Set CeRightEdge = tRgFind
    End With
End Function
'----------------------------------

'説明)指定Cellを含む列で、
'１行めから下に見て空欄でない最初のCellから
'最下行から上に見て空欄でない最初のCellまでのBoxを戻す
'★２列以上の場合の対応が必要(角の列だけを対象とする？)
Public Function BxPinchColumn(ByRef vCell As Range) As Range
    Set BxPinchColumn = Nothing
    
    Dim SRg As Range: Set SRg = CeTop(vCell)
    Dim ERg As Range: Set ERg = CeBottom(vCell)
    
    If SRg Is Nothing Then Exit Function
    If ERg Is Nothing Then Exit Function
    
    Set BxPinchColumn = vCell.Worksheet.Range(SRg.Address, ERg.Address)

End Function

'説明)指定Cellを含む行で、
'１列めから右に見て空欄でない最初のCellから
'最終列から左に見て空欄でない最初のCellまでのBoxを戻す
'★２行以上の場合の対応が必要(角の行だけを対象とする？)
Public Function BxPinchRow(ByRef vCell As Range) As Range
    Set BxPinchRow = Nothing

    Dim SRg As Range: Set SRg = CeLeftEdge(vCell)
    Dim ERg As Range: Set ERg = CeRightEdge(vCell)
    
    If SRg Is Nothing Then Exit Function
    If ERg Is Nothing Then Exit Function
    
    Set BxPinchRow = vCell.Worksheet.Range(SRg.Address, ERg.Address)

End Function

'----------------------------------
'説明)指定Areaの先頭Cellを戻す
Public Function CeFirst(ByRef vArea As Range) As Range
    Set CeFirst = vArea(1)
End Function
'説明)指定Areaの最終Cellを戻す
Public Function CeLast(ByRef vArea As Range) As Range
    Set CeLast = vArea(vArea.Count)
End Function

'----------------------------------
'説明)指定Areaを、値のある行まで上端を切り詰めたAreaを戻す
Public Function ArCeil(ByRef vArea As Range) As Range
    Set ArCeil = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    Dim SRg As Range: Set SRg = CeFirst(vArea)
    Dim ERg As Range: Set ERg = CeLast(vArea)
    
    '上から１行ずつ拾っていく。いずれかのCellに値があれば、抜ける
    Dim tRg As Range
    Dim cnt As Long: cnt = 0
    For Each tRg In vArea.Rows
        If Not IsNoVisibleText(JoinRgText(tRg)) Then Exit For
        cnt = cnt + 1
    Next
    
    '最下行まで見つからなかったら、Nothing
    If cnt = vArea.Rows.Count Then Exit Function
    
    '指定Areaから、値が見つかった行までオフセット～最下行までを戻す
    Set ArCeil = vArea.Worksheet.Range(SRg.Offset(cnt, 0), ERg)
End Function

'説明)指定Areaを、値のある行まで下端を切り詰めたAreaを戻す
Public Function ArFloor(ByRef vArea As Range) As Range
    Set ArFloor = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    Dim SRg As Range: Set SRg = CeFirst(vArea)
    
    '下から１行ずつ拾っていく。いずれかのCellに値があれば、抜ける
    Dim tRg As Range
    Dim cnt As Long: cnt = 0
    For cnt = vArea.Rows.Count To 0 Step -1
        If cnt = 0 Then Exit For 'Rowsは1始まり。0は見つからなかった扱い
        Set tRg = vArea.Rows(cnt)
        If Not IsNoVisibleText(JoinRgText(tRg)) Then Exit For
    Next
    
    '最上行まで見つからなかったら、Nothing
    If cnt = 0 Then Exit Function
    
    '指定Areaから、最上行～値が見つかった行までを戻す
    Set ArFloor = vArea.Worksheet.Range(SRg, CeLast(SRg.Offset(cnt, 0)))
End Function

'説明)指定Areaを、値のある列まで右端を切り詰めたAreaを戻す
Public Function ArRightWall(ByRef vArea As Range) As Range
    Set ArRightWall = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    Dim SRg As Range: Set SRg = CeFirst(vArea)
    
    '下から１行ずつ拾っていく。いずれかのCellに値があれば、抜ける
    Dim tRg As Range
    Dim cnt As Long: cnt = 0
    For cnt = vArea.Rows.Columns To 0 Step -1
        If cnt = 0 Then Exit For 'Rowsは1始まり。0は見つからなかった扱い
        Set tRg = vArea.Columns(cnt)
        If Not IsNoVisibleText(JoinRgText(tRg)) Then Exit For
    Next
    
    '最上行まで見つからなかったら、Nothing
    If cnt = 0 Then Exit Function
    
    '指定Areaから、最上行～値が見つかった行までを戻す
    Set ArRightWall = vArea.Worksheet.Range(SRg, CeLast(SRg.Offset(0, cnt)))

End Function
'説明)指定Areaを、値のある列まで左端を切り詰めたAreaを戻す
Public Function ArLeftWall(ByRef vArea As Range) As Range
    Set ArLeftWall = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    Dim SRg As Range: Set SRg = CeFirst(vArea)
    Dim ERg As Range: Set ERg = CeLast(vArea)
    
    '左から１列ずつ拾っていく。いずれかのCellに値があれば、抜ける
    Dim tRg As Range
    Dim cnt As Long: cnt = 0
    For Each tRg In vArea.Columns
        If Not IsNoVisibleText(JoinRgText(tRg)) Then Exit For
        cnt = cnt + 1
    Next
    
    '最下行まで見つからなかったら、Nothing
    If cnt = vArea.Columns.Count Then Exit Function
    
    '指定Areaから、値が見つかった行までオフセット～最下行までを戻す
    Set ArLeftWall = vArea.Worksheet.Range(SRg.Offset(0, cnt), ERg)
    
End Function

'----------------------------------
'----------------------------------
'説明)２つの指定Areaの重なるAreaを戻す
Public Function ArIntersect(ByRef vAreaA As Range, ByRef vAreaB As Range) As Range
    Set ArIntersect = Nothing
    
    If vAreaA Is Nothing Then Exit Function
    If vAreaB Is Nothing Then Exit Function
    
    On Error Resume Next    '重なる箇所が無いと、Intersectがエラーを出すので。
    Set RgIntersect = Intersect(vAreaA, vAreaB)
    
End Function

'説明)指定Areaの指定番目の行のBoxを戻す
Public Function BxRow(ByRef vArea As Range, ByRef vIdx As Long) As Range
    Set BxRow = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    '指定番目が、指定Areaの行数よりも多い場合もNothingを戻す
    If vIdx > vArea.Rows.Count Then Exit Function
        
    '指定番目の行とAreaの重なる範囲を取得する
    Set BxRow = ArIntersect(tRgRow.Offset(vIdx, 0).EntireRow, vArea)

End Function

'説明)指定Areaの指定番目の列のBoxを戻す
Public Function BxColumn(ByRef vArea As Range, ByRef vIdx As Long) As Range
    Set BxColumn = Nothing
    
    If vArea Is Nothing Then Exit Function
    
    '指定番目が、指定Areaの列数よりも多い場合もNothingを戻す
    If vIdx > vArea.Columns.Count Then Exit Function
        
    '指定番目の列とAreaの重なる範囲を取得する
    Set BxColumn = ArIntersect(tRgRow.Offset(0, vIdx).EntireColumn, vArea)
End Function

'----------------------------------
'説明)指定Areaの使用範囲(CurrentRegion)を戻す
Public Function ArActive(ByRef vArea As Range) As Range
    Set ArActive = CeFirst(vArea).CurrentRegion
End Function
'説明)指定Areaの属する使用範囲(CurrentRegion)の先頭Cellを戻す
Public Function CeActiveFirst(ByRef vArea As Range) As Range
    Set CeActiveFirst = CeFirst(ArActive(vArea))
End Function
'説明)指定Areaの属する使用範囲(CurrentRegion)の最終Cellを戻す
Public Function CeActiveLast(ByRef vArea As Range) As Range
    Set CeActiveLast = CeLast(ArActive(vArea))
End Function

'----------------------------------
'説明)指定Cellから右に見て値を持つcellが存在するか判定を戻す
Public Function IsExistTextAtRight(ByRef vCell As Range) As Boolean
    IsExistTextAtRight = Not IsNoTextCell(vCell.End(xlToRight))
End Function
'説明)指定Cellから左に見て値を持つcellが存在するか判定を戻す
Public Function IsExistTextAtLeft(ByRef vCell As Range) As Boolean
    IsExistTextAtLeft = Not IsNoTextCell(vCell.End(xlToLeft))
End Function
'説明)指定Cellから上に見て値を持つcellが存在するか判定を戻す
Public Function IsExistTextAtUp(ByRef vCell As Range) As Boolean
    IsExistTextAtUp = Not IsNoTextCell(vCell.End(xlUp))
End Function
'説明)指定Cellから下に見て値を持つcellが存在するか判定を戻す
Public Function IsExistTextAtDown(ByRef vCell As Range) As Boolean
    IsExistTextAtDown = Not IsNoTextCell(vCell.End(xlDown))
End Function


'----------------------------------
'説明)可視セル(Area)だけに絞り込む
Public Function ArVisible(ByRef vArea As Range) As Range
    Set ArVisible = vArea.SpecialCells(xlCellTypeVisible)
End Function

'説明)第１引数の範囲を上下に、第２引数の範囲を左右に広げて、交差する範囲(Area)を戻す
Public Function ArCross(ByRef vAreaA As Range, ByRef vAreaB As Range) As Range
    Set ArCross = ArIntersect(vAreaA.EntireColumn, vAreaB.EntireRow)
End Function

'説明)指定Cellを左上の起点として、右に指定長、下に指定長に伸ばしたAreaを戻す
Public Function ArSpread(ByRef vCell As Range, ByVal vRightLength As Long, ByVal vDownLength As Long) As Range
    Set ArSpread = vCell.Worksheet.renge(vCell, vCell.Offset(vDownLength, vRightLength))
End Function

'説明)指定Areaの内容をクリアする
Public Function ClearRgContents(ByRef vArea As Range) As Range
    Set ClearRgContents = vArea '戻り値は、そのままのエリアを戻す
    If vArea Is Nothing Then Exit Function
    vArea.ClearComments
End Function

'----------------------------------
'説明)指定Areaを、Joinした文字列を戻す
'Cellの間は、指定されたデリミタ(defaultはTab)でつなげる。
'複数行の場合、改行区切り
Public Function JoinRgText(ByRef vArea As Range, Optional ByVal vDelimiter As String = vbTab, Optional vAddEndReturn As Boolean = True) As String
    Dim ret() As String: Dim retRow() As String
    Dim tRg As Range: Dim tRgRow As Range
    Dim cntRow As Long: Dim cnt As Long
    
    cntRow = 0
    For Each tRgRow In vArea.Rows
    
        '行で切り出し、各Cellの値をつなげる
        cnt = 0
        For Each tRg In tRgRow.Cells
            ReDim Preserve ret(cnt)
            ret(cnt) = tRg.Text
            cnt = cnt + 1
        Next
        
        '1行ごとの結果を格納
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
'説明)指定Areaから、右方向に指定文言を検索し、見つかったCellを戻す
Public Function CeFindRight(ByRef vArea As Range, ByVal vString As String) As Range
    Set CeFindRight = vArea.Find(What:=vString, LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByRows, MatchByte:=False)
End Function

'説明)指定Areaから、下方向に指定文言を検索し、見つかったCellを戻す
Public Function CsFindDown(ByRef vArea As Range, ByVal vString As String) As Range
    Set CeFindDown = vArea.Find(What:=vString, LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, MatchByte:=False)
End Sub

'説明)指定Areaから、指定文言を含むCellの集合を戻す
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

'説明)指定Areaから、正規表現に該当するCellの集合を戻す
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


