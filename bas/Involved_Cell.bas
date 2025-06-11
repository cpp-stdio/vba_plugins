Attribute VB_Name = "Involved_Cell"
Option Explicit
'##############################################################################################################################
'
'   シート関連
'
'   新規作成日 : 2022/11/10
'   最終更新日 : 2024/01/30
'
'   新規作成エクセルバージョン : Office Professional Plus 2016 , 16.0.5.56.1000(32ビット)
'   最終更新エクセルバージョン : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   LEGACY_aDeleteStrikethroughで使用される構造体
'   参考URL：https://vbabeginner.net/remove-strikethroughs-preserving-fonts/
'==============================================================================================================================
Private Type ST_FONT
    Background      As Long
    Bold            As Boolean
    Color           As Double
    ColorIndex      As Long
    FontStyle       As String
    Italic          As Boolean
    name            As String
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
'   取り消し線のついた文字のみ削除する（バグ有だったが修正済み）
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
Public Function LEGACY_aDeleteStrikethrough(ByRef r As range) As range

    ' 戻り値を仮設定
    Set LEGACY_aDeleteStrikethrough = r

    Dim strike  As Variant
    strike = r.Font.Strikethrough
    
    'セル未設定時は処理終了
    If StrComp(CStr(r), "", vbBinaryCompare) = 0 Then Exit Function
    
    '取り消し線が設定されているフラグがFalseの場合処理をしない
    If Not IsNull(strike) And strike = False Then Exit Function
    
On Error GoTo aDeleteStrikethrough_ErrorHandler '下記で謎エラーが発生することがある

    Dim i       As Long         ' 文字列長ループカウンタ
    Dim iLen    As Long         ' セル文字列長
    Dim c       As Characters   ' 文字列のCharactersオブジェクト
    Dim f       As Font         ' 1文字ごとのFontオブジェクト
    Dim fAr()   As ST_FONT      ' Fontオブジェクト設定値保持用の構造体配列
    Dim s       As String       ' 取り消し線除去済みの文字列
    Dim iFont   As Long         ' Fontオブジェクト設定用配列のインデックス

    iFont = 0
    iLen = Len(CStr(r.value))
    ReDim fAr(iLen)
    
    '// セル文字列を１文字ずつループ
    For i = 1 To iLen
        '// 1文字分のCharactersオブジェクトを取得
        Set c = r.Characters(i, 1)
        
        '// Fontオブジェクトを取得
        Set f = c.Font
        
        '// 対象の１文字に取り消し線が設定されていない場合
        If f.Strikethrough = False And Not StrComp("", CStr(c.text), vbBinaryCompare) = 0 Then
            '// 取り消し線未設定の文字列を取得
            s = s & c.text
            
            '// Fontオブジェクトの各プロパティを保持
            fAr(iFont).name = f.name
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
        Set f = r.Characters(Start:=i, length:=1).Font
        
        '// インデックス取得
        iFont = i - 1
        
        '// Fontオブジェクトの各プロパティを保持しておいた値で再設定
        f.name = fAr(iFont).name
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
    
    Set LEGACY_aDeleteStrikethrough = r
    Exit Function
    
aDeleteStrikethrough_ErrorHandler:
    Set LEGACY_aDeleteStrikethrough = Nothing
    
End Function

'==============================================================================================================================
'   取り消し線のついた文字のみ削除する（軽量版）
'   ※注意：セルの書式までコピーできない
'
'   参考URL：https://stabucky.com/wp/archives/3209
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
Public Function LEGACY_aDeleteStrikethrough_verLight(ByRef r As range) As range

    ' 戻り値を仮設定
    Set LEGACY_aDeleteStrikethrough_verLight = r

    Dim strike  As Variant
    strike = r.Font.Strikethrough
    
    'セル未設定時は処理終了
    If StrComp(CStr(r), "", vbBinaryCompare) = 0 Then Exit Function
    
    '取り消し線が設定されているフラグがFalseの場合処理をしない
    If Not IsNull(strike) And strike = False Then Exit Function
    
    Dim i As Long
    ' テキストを取得
    Dim textBefore As String: textBefore = CStr(r)
    Dim textAfter As String: textAfter = ""
    
    For i = 1 To Len(textBefore)
        ' Strikethroughの値がFalseの場合のみ取り出す
        If r.Characters(Start:=i, length:=1).Font.Strikethrough = False Then
            textAfter = textAfter + Mid(textBefore, i, 1)
        End If
    Next i
    
    '戻り値の値の方にのみにセット
    LEGACY_aDeleteStrikethrough_verLight = textAfter
End Function
'==============================================================================================================================
'   NumberFormatやNumberFormatLocalで表示がおかしくなっているセルを修正するため
'   ※数が多いので全て対応はしきれないため、随時追加をお願いします
'
'   使い方：
'
'           If aTypeErrorIsNumberFormat(r) = 7 Then
'               ...
'               ...
'           End If
'
'   戻り値 : VarType : https://www.sejuku.net/blog/68632
'           No  引数に入れる値  実行結果
'           1   Integer         2
'           2   Double          5
'           3   String          8
'           4   Boolean         11
'           5   Date            7
'           6   Object          9
'           7   Variant         0
'           8   String()        8200
'           9   Integer()       8194
'
'   r : 対象セル
'==============================================================================================================================
Public Function LEGACY_aTypeErrorIsNumberFormat(ByRef r As range) As Long
    
    '--------------------------------------------------------------
    '   Date型の場合
    LEGACY_aTypeErrorIsNumberFormat = 7

    '日付＆時刻
    If r.NumberFormatLocal Like "m""月""d""日""" Then Exit Function
    If r.NumberFormatLocal = "yyyy/mm/dd" Then Exit Function
    If r.NumberFormatLocal = "yyyy年mm月dd日" Then Exit Function
    If r.NumberFormatLocal = "ggge年mm月dd日" Then Exit Function
    If r.NumberFormatLocal = "yyyy/mm/dd(aaa)" Then Exit Function
    If r.NumberFormatLocal = "yyyy/mm/dd hh:mm:ss" Then Exit Function
    '時刻
    If r.NumberFormatLocal = "hh時mm分dd秒" Then Exit Function
    If r.NumberFormatLocal = "hh:mm:dd" Then Exit Function

    '--------------------------------------------------------------
    '   Double型の場合
    LEGACY_aTypeErrorIsNumberFormat = 5

    '数値
    If r.NumberFormatLocal = "#0.000" Then Exit Function
    If r.NumberFormatLocal = "#,##0" Then Exit Function
    '通貨
    If r.NumberFormatLocal = "\#,##0" Then Exit Function
    If r.NumberFormatLocal = "#,##0円" Then Exit Function

    '--------------------------------------------------------------
    '   String型の場合
    LEGACY_aTypeErrorIsNumberFormat = 8

    '文字列
    If r.NumberFormatLocal = "G/標準" Then Exit Function
    If r.NumberFormatLocal = "@" Then Exit Function
    
    '--------------------------------------------------------------
    '見つからなかった場合は文字列と仮定する
    LEGACY_aTypeErrorIsNumberFormat = 8
End Function
