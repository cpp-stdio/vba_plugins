Attribute VB_Name = "Involved_Call"
Option Explicit
'##############################################################################################################################
'
'   シート関連
'
'   新規作成日 : 2022/11/10
'   最終更新日 : 2022/11/21
'
'   新規作成エクセルバージョン : Office Professional Plus 2016 , 16.0.5.56.1000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2016 , 16.0.5.56.1000(32ビット)
'
'##############################################################################################################################
'==============================================================================================================================
'   aDeleteStrikethroughで使用される構造体
'   参考URL：https://vbabeginner.net/remove-strikethroughs-preserving-fonts/
'==============================================================================================================================
Private Type ST_FONT
    Background      As Long
    Bold            As Boolean
    Color           As Double
    ColorIndex      As Long
    FontStyle       As String
    Italic          As Boolean
    Name            As String
    OutlineFont     As Boolean
    Shadow          As Boolean
    Size            As Double
    Strikethrough   As Boolean
    Subscript       As Boolean
    Superscript     As Boolean
    ThemeColor      As Variant
    ThemeFont       As XlThemeFont
    TintAndShade    As Double
    Underline       As Long
End Type
'==============================================================================================================================
'   そのシートで入力されている数式を値に変換する
'   利用する場合は「Involved_Sheet.bas」のインポートをお願いします
'   ※数式部分の自動再計算を行うので動作が重くなる可能性があります
'
'   戻り値 : 変換完了(True), NG(False)
'
'   sheetName : シート名
'==============================================================================================================================
Public Function aCallDeleteFormula(ByRef range As range) As Boolean
    aCallDeleteFormula = False
    'If sheet = Nothing Then Exit Function
    
    'Dim cell As range
    'Dim row As Long
    'Dim rowMax As Long
    'Dim column As Long
    'Dim columnMax As Long
    'Dim value As String
    
    'Set base = sheet.UsedRange.range("A1")
    'rowMax = sheet.UsedRange.Rows.count - 1
    'columnMax = sheet.UsedRange.Columns.count - 1
    
    'For row = rowMax To 0 Step -1
    '    For column = columnMax To 0 Step -1
    '        Set cell = base.Offset(row, column)
    '
    '        If WorksheetFunction.IsFormula(cell) Then
    '            'cell.Calculate '再計算
    '            value = cell.value
    '            cell.NumberFormatLocal = "@"
    '            cell.value = value
    '        End If
    '    Next
    'Next
    
    'Set base = Nothing
    'Set cell = Nothing
    aCallDeleteFormula = True
End Function
'==============================================================================================================================
'   取り消し線のついた文字のみ削除する
'   参考URL：https://vbabeginner.net/remove-strikethroughs-preserving-fonts/
'
'   使い方：
'               Dim r   As range    '// セル
'               For Each r In Selection
'                   Call aCallDeleteStrikethrough(r)
'               Next
'
'   戻り値 : 削除されたセル, エラーの場合はNothing
'
'   r : 対象セル
'==============================================================================================================================
Public Function aDeleteStrikethrough(ByRef r As range) As range

On Error GoTo ErrorHandler '下記で謎エラーが発生することがある
 
    Dim i       As Long         '// 文字列長ループカウンタ
    Dim iLen    As Long         '// セル文字列長
    Dim C       As characters   '// 文字列のCharactersオブジェクト
    Dim f       As Font         '// 1文字ごとのFontオブジェクト
    Dim fAr()   As ST_FONT      '// Fontオブジェクト設定値保持用の構造体配列
    Dim s       As String       '// 取り消し線除去済みの文字列
    Dim iFont   As Long         '// Fontオブジェクト設定用配列のインデックス
    
    '// セル未設定時は処理終了
    If (r.value = "") Then Exit Function
    
    iFont = 0
    iLen = Len(r.value)
    ReDim fAr(iLen)
    
    '// セル文字列を１文字ずつループ
    For i = 1 To iLen
        '// 1文字分のCharactersオブジェクトを取得
        Set C = r.characters(i, 1)
        '// Fontオブジェクトを取得
        Set f = C.Font
        
        '// 対象の１文字に取り消し線が設定されていない場合
        If f.Strikethrough = False Then
            '// 取り消し線未設定の文字列を取得
            s = s & C.text
            
            '// Fontオブジェクトの各プロパティを保持
            fAr(iFont).Name = f.Name
            fAr(iFont).FontStyle = f.FontStyle
            fAr(iFont).Size = f.Size
            fAr(iFont).Strikethrough = f.Strikethrough
            fAr(iFont).Superscript = f.Superscript
            fAr(iFont).Subscript = f.Subscript
            fAr(iFont).OutlineFont = f.OutlineFont
            fAr(iFont).Shadow = f.Shadow
            fAr(iFont).Underline = f.Underline
            'fAr(iFont).ThemeColor = f.ThemeColor
            fAr(iFont).Color = f.Color
            fAr(iFont).TintAndShade = f.TintAndShade
            fAr(iFont).ThemeFont = f.ThemeFont
 
            iFont = iFont + 1
        End If
    Next
    
    '// 取り消し線を除いた文字列をセルに設定
    r.FormulaR1C1 = s
    
    '// 再度セルの文字列長を取得
    iLen = Len(s)
    
    '// 取り消し線を除いた文字列を１文字ずつループ
    For i = 1 To iLen
        '// 1文字分のFontオブジェクトを再設定のため取得
        Set f = r.characters(Start:=i, length:=1).Font
        
        '// インデックス取得
        iFont = i - 1
        
        '// Fontオブジェクトの各プロパティを保持しておいた値で再設定
        f.Name = fAr(iFont).Name
        f.FontStyle = fAr(iFont).FontStyle
        f.Size = fAr(iFont).Size
        f.Strikethrough = fAr(iFont).Strikethrough
        f.Superscript = fAr(iFont).Superscript
        f.Subscript = fAr(iFont).Subscript
        f.OutlineFont = fAr(iFont).OutlineFont
        f.Shadow = fAr(iFont).Shadow
        f.Underline = fAr(iFont).Underline
        'f.ThemeColor = fAr(iFont).ThemeColor
        f.Color = fAr(iFont).Color
        f.TintAndShade = fAr(iFont).TintAndShade
        f.ThemeFont = fAr(iFont).ThemeFont
    Next
    
    Set aDeleteStrikethrough = r
    Exit Function
ErrorHandler:
    Set aDeleteStrikethrough = Nothing
End Function
