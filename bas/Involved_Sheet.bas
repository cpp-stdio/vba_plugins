Attribute VB_Name = "Involved_Sheet"
Option Explicit
'##############################################################################################################################
'
'   シート関連
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2024/07/05
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   その名前がシート名に適切な名前であるか検査する
'
'   戻り値 : OK(True), NG(False)
'
'   sheetName : シート名
'==============================================================================================================================
Public Function LEGACY_checkSheetName(ByVal sheetname As String) As Boolean
    LEGACY_checkSheetName = False
    '条件その1 : 空の名前ではない。
    If StrComp(sheetname, "", vbBinaryCompare) = 0 Then Exit Function
    '条件その2 : 含んではいけない文字列がない。
    Dim textFor As Variant
    For Each textFor In Array(":", "\", "/", "?", "*", "[", "]")
        If InStr(sheetname, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    '条件その3 : 名前は31文字以内である。
    If Len(sheetname) > 31 Then Exit Function
    '条件その4 : 同名のシートは存在出来ない。
    'aNewSheetにて不具合が発生したので分割する。
    LEGACY_checkSheetName = True
End Function
'==============================================================================================================================
'   等しい名前のシートを探す
'
'   戻り値 : 等しい名前を持つシート。ない場合は、Nothingが返却される
'
'   sheetName : シート名
'   book : 対象のブック（任意）
'==============================================================================================================================
Public Function LEGACY_sheetToEqualsName(ByVal sheetname As String, Optional ByRef book As Workbook = Nothing) As Worksheet

    Dim searchBook As Workbook
    Set searchBook = isBook(book)

    Dim sheet As Worksheet
    For Each sheet In searchBook.sheets
        If StrComp(sheet.name, sheetname, vbBinaryCompare) = 0 Then
            Set LEGACY_sheetToEqualsName = sheet
            Exit Function
        End If
    Next
    Set LEGACY_sheetToEqualsName = Nothing
End Function
'==============================================================================================================================
'   新たなシートを作成
'
'   戻り値 :新規作成されたWorksheetが返却され、作成済の場合はそのWorksheetが返却される。
'           作成出来なかった場合はNothingが返却される
'
'   sheetName : シート名
'   book : 対象のブック（任意、未入力の場合ThisWorkbook）
'==============================================================================================================================
Public Function LEGACY_aNewSheet(ByVal sheetname As String, Optional ByRef book As Workbook = Nothing) As Worksheet
    Set LEGACY_aNewSheet = Nothing
    '適切な名前でない場合
    If Not LEGACY_checkSheetName(sheetname) Then Exit Function
    '対象のブックが入力されていない場合
    Dim addBook As Workbook
    Set addBook = isBook(book)
    '作成済みかを検索
    Dim sheet As Worksheet
    Set sheet = LEGACY_sheetToEqualsName(sheetname, addBook)
    If Not sheet Is Nothing Then
        Set LEGACY_aNewSheet = sheet
        Exit Function
    End If
    '新たなシートを末尾へ作成する
    addBook.sheets.add After:=Worksheets(Worksheets.count)
    Set sheet = addBook.sheets.Item(addBook.sheets.count)
    sheet.name = sheetname
    sheet.Activate 'アクティブ化しておいた方が見た目は良い。
    Set LEGACY_aNewSheet = sheet
End Function
'==============================================================================================================================
'   ブック内にある全シート名を取得
'
'   戻り値 :見つかったシート名を配列Stringとして返却する
'           作成出来なかった場合はNothingが返却される
'
'   book : 対象のブック（任意、未入力の場合ThisWorkbook）
'==============================================================================================================================
Public Function LEGACY_getSheetNames(Optional ByRef book As Workbook = Nothing) As String()
    Dim r() As String
    Dim l As Long: l = 0
    '対象のブックが入力されていない場合
    Dim getBook As Workbook
    Set getBook = isBook(book)
    Dim sheet As Worksheet
    For Each sheet In getBook.sheets
        ReDim Preserve r(l)
        r(l) = sheet.name
        l = l + 1
    Next
    LEGACY_getSheetNames = r
End Function
'==============================================================================================================================
'       long型の数値から列番号(AX等)が必要になる場合がセルに数式を埋め込んで速度上昇を狙い際にいるが一行で出来た方がいいと判断したため
'
'   column : 変換したいLong型
'==============================================================================================================================
Public Function LEGACY_isColumnNumber_toString(column As Long) As String
    LEGACY_isColumnNumber_toString = ""

    If column <= 0 Then Exit Function
    
    Dim tmp As Variant
    tmp = Split(Cells(1, column).Address(True, False), "$")
    LEGACY_isColumnNumber_toString = tmp(0)

End Function
'==============================================================================================================================
'   Long型等の数値からString型(AX10)等のアルファベット文字列型の変更する
'   行情報セットタイプ
'
'   戻り値 : NG／空白 , OK／空白以外のアルファベット数文字
'
'       row    : 変換したいLong型
'   column : 変換したいLong型
'==============================================================================================================================
Public Function LEGACY_isColumnNumberAndRow_toString(row As Long, column As Long) As String
    LEGACY_isColumnNumberAndRow_toString = ""

    If row <= 0 Then Exit Function
    If column <= 0 Then Exit Function
    
    Dim tmp As Variant
    tmp = Split(Cells(row, column).Address(True, False), "$")
    LEGACY_isColumnNumberAndRow_toString = tmp(0) + tmp(1)

End Function
'==============================================================================================================================
'   シートを削除する
'
'   戻り値 : 成功(True), 失敗(False)
'
'   sheet : 削除するシート。成功した場合、アクセス不可になるので注意が必要
'   book  : 対象のブック（任意）
'==============================================================================================================================
Public Function LEGACY_aDeletedSheet(ByVal sheetname As String, Optional ByRef book As Workbook = Nothing) As Boolean
    Dim sheet As Worksheet
    Set sheet = LEGACY_sheetToEqualsName(sheetname, book)
    LEGACY_aDeletedSheet = LEGACY_aDeletedSheetEx(sheet, book)
    Set sheet = Nothing
End Function

Public Function LEGACY_aDeletedSheetEx(ByRef sheet As Worksheet, Optional ByRef book As Workbook = Nothing) As Boolean
    LEGACY_aDeletedSheetEx = False
    
    If sheet Is Nothing Then
        'Nothingなので、既に削除済みと仮定する。
        LEGACY_aDeletedSheetEx = True
        Exit Function
    End If
    
    '削除するタイミングでメッセージが表示されるが機能的に不必要なので非表示にしておく
    Application.DisplayAlerts = False
    Dim deleteBook As Workbook
    Set deleteBook = isBook(book)
    
    Dim deleteSheet As Worksheet
    For Each deleteSheet In deleteBook.sheets
        If StrComp(sheet.name, deleteSheet.name, vbBinaryCompare) = 0 Then
            Call deleteBook.sheets(sheet.name).delete
            Set sheet = Nothing  'シートを削除する
            LEGACY_aDeletedSheetEx = True '戻り値を変更
            Exit For
        End If
    Next
    
    'メッセージを表示状態に戻す
    Application.DisplayAlerts = True
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   シートの情報を全て削除する
'
'   sheet : 対象シート
'------------------------------------------------------------------------------------------------------------------------------
Public Function LEGACY_aInfoErasureSheet(ByRef sheet As Worksheet)
    Dim i As Long: i = 0
    'セルを全て削除
    sheet.Cells.clear
    sheet.Columns.clear
    sheet.Rows.clear
    'テーブルの情報を削除
    For i = sheet.ListObjects.count To 1 Step -1
        Call sheet.ListObjects.Item(i).delete
    Next i
    '埋め込みグラフを削除
    For i = sheet.ChartObjects.count To 1 Step -1
        Call sheet.ChartObjects(i).delete
    Next i
    '印刷時のページ区切りを削除
    'sheet.DisplayPageBreaks = False
    'ピボットテーブルを削除
    For i = sheet.PivotTables.count To 1 Step -1
        Call sheet.PivotTables(i).ClearTable
    Next i
    '図、クリップアート、図形、SmartArtの削除
    For i = sheet.Shapes.count To 1 Step -1
        Call sheet.Shapes.Item(i).delete
    Next i
    'ヘッター、フッターは完全に削除することは不可能らしい
    With sheet.PageSetup
        For i = .Pages.count To 1 Step -1
            .Pages.Item(i).CenterFooter = ""
            .Pages.Item(i).CenterHeader = ""
            .Pages.Item(i).LeftFooter = ""
            .Pages.Item(i).LeftHeader = ""
            .Pages.Item(i).RightFooter = ""
            .Pages.Item(i).RightHeader = ""
        Next i
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .DifferentFirstPageHeaderFooter = True
    End With
    
End Function

'==============================================================================================================================
'   シート内で入力されている数式を値に変換する
'   ※1　利用する場合は「Involved_Other.bas」のインポートをお願いします
'   ※2　数式部分の自動再計算を行うので動作が重くなる可能性があります
'   ※3　「Involved_Call」に単一セルの数式を値に変換するプログラムがあります。
'
'   戻り値 : 変換完了(True), NG(False)
'
'   sheetName : シート名
'==============================================================================================================================
Public Function LEGACY_aSheetDeleteFormula(ByVal sheetname As String, Optional ByRef book As Workbook = Nothing) As Boolean
    LEGACY_aSheetDeleteFormula = False

    Dim sheet As Worksheet
    Set sheet = LEGACY_sheetToEqualsName(sheetname, book)
    If sheet Is Nothing Then Exit Function
    
    LEGACY_aSheetDeleteFormula = LEGACY_aSheetDeleteFormulaDx(sheet)
    Set sheet = Nothing
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   シート用 ver.
'
'   sheet : シートを挿入(Nothingの場合無効)
'------------------------------------------------------------------------------------------------------------------------------
Public Function LEGACY_aSheetDeleteFormulaDx(ByRef sheet As Worksheet) As Boolean
    LEGACY_aSheetDeleteFormulaDx = False
    If sheet Is Nothing Then Exit Function
    
    Dim base As Range
    Dim cell As Range
    Dim row As Long
    Dim rowMax As Long
    Dim column As Long
    Dim columnMax As Long
    Dim text As String
    Dim value As Variant
    
    Set base = sheet.UsedRange.Range("A1")
    rowMax = sheet.UsedRange.Rows.count - 1
    columnMax = sheet.UsedRange.Columns.count - 1
    
    For row = rowMax To 0 Step -1
        For column = columnMax To 0 Step -1
            Set cell = base.Offset(row, column)
            
            If WorksheetFunction.IsFormula(cell) Then
                cell.Calculate '再計算
                text = cell.value
                '数値の場合はそのまま"数値"として表示させる（日付、金額等は対象外）
                If checkNumericalValue(text, value) Then
                    cell.value = value
                Else
                    cell.NumberFormatLocal = "@"
                    cell.value = text
                End If
            End If
        Next
    Next
    
    Set base = Nothing
    Set cell = Nothing
    LEGACY_aSheetDeleteFormulaDx = True
End Function
'==============================================================================================================================
'   特定のシート名を別ブックのとして特定の場所に保存する
'
'   戻り値 : Workbook（NGの場合はNothingが返却される）
'
'   sheetname   : シート名
'   filename    : 保存名（空白の場合は「シート名.xlsx」となる）
'   pathname    : パス名（空白の場合は本体のブックと同階層になる）
'   book        : ブック（Nothingの場合はThisWorkbookとしてみなす）
'==============================================================================================================================
Public Function LEGACY_saveSheet(ByVal sheetname As String, Optional ByVal fileName As String = "", _
                          Optional ByVal pathname As String = "", Optional ByRef book As Workbook = Nothing) As Workbook
    
    Set saveSheet = Nothing

    Dim sheet As Worksheet
    Set sheet = LEGACY_sheetToEqualsName(sheetname, book)
    
    Set LEGACY_saveSheet = LEGACY_saveSheetEx(sheet, fileName, pathname)
    Set sheet = Nothing
    
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   シート用 ver.
'
'   sheet       : シート本体
'   filename    : 保存名（空白の場合は「シート名.xlsx」となる）
'   pathname    : パス名（空白の場合は本体のブックと同階層になる）
'------------------------------------------------------------------------------------------------------------------------------
Public Function LEGACY_saveSheetEx(ByRef sheet As Worksheet, Optional fileName As String = "", _
                                                      Optional pathname As String = "") As Workbook
                            
    Set LEGACY_saveSheetEx = Nothing

    If sheet Is Nothing Then Exit Function
    
    If StrComp(fileName, "", vbBinaryCompare) = 0 Then
        fileName = sheet.name + ".xlsx"
    End If
    
    If StrComp(pathname, "", vbBinaryCompare) = 0 Then
        pathname = ThisWorkbook.path
    End If
    
    sheet.copy                        '別のブックへコピー
    
    Application.DisplayAlerts = False '下の関数を動かすとメッセージが表示されてしまうため
    Call ActiveWorkbook.SaveAs(pathname + "\" + fileName)
    'Call ActiveWorkbook.Activate
    Application.DisplayAlerts = True  'メッセージ表示防止解除

    Set LEGACY_saveSheetEx = ActiveWorkbook
    
End Function
    
'==============================================================================================================================
'   ブックの有無
'==============================================================================================================================
Private Function isBook(ByRef book As Workbook) As Workbook
    If book Is Nothing Then
        Set isBook = ThisWorkbook
    Else
        Set isBook = book
    End If
End Function

