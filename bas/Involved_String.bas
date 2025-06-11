Attribute VB_Name = "Involved_String"
Option Explicit
'##############################################################################################################################
'
'   文字列(String)でVBAの標準機能だけでは足りない部分を追加する
'   ※ 2024/01/30：Involved_Otherから独立
'
'   新規作成日 : 2024/01/30
'   最終更新日 : 2024/01/30
'
'   新規作成エクセルバージョン : Microsoft 365 Apps for enterprise
'   最終更新エクセルバージョン : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'
'   数値を判定
'   戻り値 : はい(true),いいえ(false)
'
'   text  : 判定用の数値
'   value : 数数値の入った数値型(Long,Double)のどちらか、エラーの場合はEmptyが入る
'           最終的には型の判定が要ります。↓参考URL：例→ If VarType(value) = vbLong Then
'           http://officetanaka.net/excel/vba/function/VarType.htm
'
'==============================================================================================================================
Public Function LEGACY_checkNumericalValue(ByVal text As String, Optional ByRef value As Variant = Empty) As Boolean

    text = StrConv(text, vbNarrow)
    text = StrConv(text, vbLowerCase)
    text = LCase(text)
    If IsNumeric(text) Then
        value = Val(text)
        If StrComp(CStr(value), CStr(CLng(CStr(value))), vbBinaryCompare) = 0 Then
            value = CLng(CStr(value))
        End If
        LEGACY_checkNumericalValue = True
    Else
        value = Empty
        LEGACY_checkNumericalValue = False
    End If
End Function

'==============================================================================================================================
'
'   文字列の中から、数字のみを抜き出す。参考URL↓
'   https://vbabeginner.net/vba%E3%81%A7%E6%96%87%E5%AD%97%E5%88%97%E3%81%8B%E3%82%89%E6%95%B0%E5%AD%97%E3%81%AE%E3%81%BF%E3%82%92%E6%8A%BD%E5%87%BA%E3%81%99%E3%82%8B/
'
'   戻り値 : 抜き出した数字、エラーの場合は空の配列が返却されます。
'
'   text  : 数字が含まれる文字列
'
'==============================================================================================================================
Public Function LEGACY_findNumber(ByVal text As String) As Variant()
    Dim reg As Object     '正規表現クラスオブジェクト
    Dim matches As Object 'RegExp.Execute結果
    Dim match As Object   '検索結果オブジェクト
    Dim i As Long         'ループカウンタ
    
    Dim returnVariant() As Variant
    ReDim returnVariant(0)
    LEGACY_findNumber = returnVariant
    
    Set reg = CreateObject("VBScript.RegExp")
    
    '検索範囲＝文字列の最後まで検索
    reg.Global = True
    '検索条件＝数字を検索
    reg.Pattern = "[0-9]"
    '検索実行
    Set matches = reg.Execute(text)
    '検索一致件数だけループ
    For i = 0 To matches.count - 1
        'コレクションの現ループオブジェクトを取得
        Set match = matches.Item(i)
        '検索一致文字列
        ReDim Preserve returnVariant(i)
        returnVariant(i) = match.value
    Next
    LEGACY_findNumber = returnVariant
End Function

'==============================================================================================================================
'
'   改行コードのみを取り替えるプログラム
'   エクセルでは改行コードの種類が意外に多いため開発
'
'   戻り値 :　改行コードが消された文字列
'
'   text : 文字列
'   replaceText : 改行コードと取り替える文字列（任意）
'
'==============================================================================================================================
Public Function LEGACY_ReplaceEnter(ByVal text As String, Optional ByVal replaceText As String = "") As String

    text = Replace(text, vbCr, replaceText)
    text = Replace(text, vbLf, replaceText)
    text = Replace(text, vbCrLf, replaceText)
    text = Replace(text, vbNewLine, replaceText)
    
    LEGACY_ReplaceEnter = text
    
End Function

'==============================================================================================================================
'
'   文字列の中から、特定の文字より前後を抽出
'   インターネットに書かれた情報通りだと特定の文字がない場合エラーになりプログラムが止まってしまうので自作
'
'   戻り値 : 抜き出した数字、エラーの場合は空の配列が返却されます。
'
'   text : 文字列
'   deleteText : 削除する文字列
'
'
'==============================================================================================================================
Public Function LEGACY_LeftInStrString(ByVal text As String, ByVal deleteText As String) As String

    Dim r As String: r = text
    Dim i As Long: i = InStr(text, deleteText)
    If i >= 1 Then
        r = Left(text, i - 1)
    End If
    LEGACY_LeftInStrString = r
End Function

Public Function LEGACY_RigetInStrString(ByVal text As String, ByVal deleteText As String) As String

    Dim r As String: r = text
    Dim i As Long: i = InStr(text, deleteText)
    Dim l As Long: l = Len(deleteText)
    If i >= 1 And l >= 0 Then
        r = Mid(text, i + l)
    End If
    LEGACY_RigetInStrString = r
End Function
'==============================================================================================================================
'
'   文字列の中身がアルファベットのみで構成されているか？
'   https://vbabeginner.net/isalpha/
'
'   戻り値 : True(アルファベットのみ), False(それ以外も含む)
'
'==============================================================================================================================
Public Function LEGACY_IsAlphabets(text As String) As Boolean
    
    LEGACY_IsAlphabets = False
    If text = "" Then Exit Function
    
    LEGACY_IsAlphabets = Not text Like "*[!a-zA-Zａ-ｚＡ-Ｚ]*"
    
End Function
